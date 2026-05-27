const express = require("express");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const mammoth = require("mammoth");
const session = require("express-session");
const Anthropic = require("@anthropic-ai/sdk");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;
const PASSWORD = process.env.APP_PASSWORD || "Jasmine1!";
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const TRAINING_MAILBOX = process.env.TRAINING_MAILBOX || "training@risk2solution.com";
const DIANE_EMAIL = process.env.DIANE_EMAIL || "diane.k@risk2solution.com";
const KANDIA_EMAIL = process.env.KANDIA_EMAIL || "kandia@risk2solution.com";

const client = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

// In-memory stores (no database yet — same pattern as Ariel)
const reviewQueue = [];
const bookingsLog = [];
const emailLog = [];

let graphToken = null;
let graphTokenExpiry = null;

// ── KNOWN CLIENT DOMAINS ──────────────────────────────────────────────────────
const CLIENT_DOMAINS = [
  "@virginaustralia.com",
  "@vline.com.au",
  "@deeca.vic.gov.au",
  "@wayss.org.au",
  "@scentregroup.com",
  "kandia@risk2solution.com" // TEMP: remove after testing
];

function getClientFromEmail(email) {
  const from = (email || "").toLowerCase();
  if (from.includes("@virginaustralia.com")) return "Virgin Australia";
  if (from.includes("@vline.com.au")) return "V/Line";
  if (from.includes("@deeca.vic.gov.au")) return "DEECA";
  if (from.includes("@wayss.org.au")) return "Wayss";
  if (from.includes("@scentregroup.com")) return "Scentregroup";
  // Internal senders (staff forwarding emails for testing)
  if (from.includes("@risk2solution.com")) return "INTERNAL";
  return null;
}

function detectClientFromContent(bodyText) {
  const text = (bodyText || "").toLowerCase();
  if (text.includes("virginaustralia.com") || text.includes("virgin australia")) return "Virgin Australia";
  if (text.includes("vline.com.au") || text.includes("v/line")) return "V/Line";
  if (text.includes("deeca.vic.gov.au") || text.includes("deeca") || text.includes("conflict training request") || text.includes("request for service")) return "DEECA";
  if (text.includes("wayss.org.au") || text.includes("wayss")) return "Wayss";
  if (text.includes("scentregroup.com") || text.includes("scentre")) return "Scentregroup";
  return null;
}

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || "jasmine-r2s-secret-2026",
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 24 * 60 * 60 * 1000 }
}));
app.use(express.static("public"));

function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) return next();
  res.status(401).json({ error: "Not authenticated" });
}

// ── MICROSOFT GRAPH ───────────────────────────────────────────────────────────
async function getGraphToken() {
  if (graphToken && graphTokenExpiry && Date.now() < graphTokenExpiry - 60000) return graphToken;
  if (!AZURE_CLIENT_ID || !AZURE_TENANT_ID || !AZURE_CLIENT_SECRET) throw new Error("Azure credentials not configured");
  const url = "https://login.microsoftonline.com/" + AZURE_TENANT_ID + "/oauth2/v2.0/token";
  const body = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString()
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("Failed to get Graph token: " + JSON.stringify(data));
  graphToken = data.access_token;
  graphTokenExpiry = Date.now() + (data.expires_in * 1000);
  console.log("Graph token acquired.");
  return graphToken;
}

async function graphRequest(method, endpoint, body, token) {
  const t = token || await getGraphToken();
  const url = endpoint.startsWith("https") ? endpoint : "https://graph.microsoft.com/v1.0" + endpoint;
  const options = {
    method,
    headers: { "Authorization": "Bearer " + t, "Content-Type": "application/json" }
  };
  if (body) options.body = JSON.stringify(body);
  const res = await fetch(url, options);
  if (res.status === 204) return null;
  const text = await res.text();
  try { return JSON.parse(text); } catch { return text; }
}

// ── JASMINE SYSTEM PROMPT (PROC_SYS) ─────────────────────────────────────────
// Load from file if it exists, otherwise use inline
let JASMINE_PROC_SYS = "";
try {
  JASMINE_PROC_SYS = fs.readFileSync(path.join(__dirname, "jasmine_prompt.txt"), "utf8").trim();
  console.log("Loaded Jasmine system prompt from file (" + JASMINE_PROC_SYS.length + " chars)");
} catch {
  console.log("jasmine_prompt.txt not found — using default prompt");
  JASMINE_PROC_SYS = `RESPOND WITH VALID JSON ONLY. Start with { end with }. No text outside JSON. No preamble. No explanations. Never.

FORBIDDEN: Do not write "TBC" in date or session_start fields when dates/times are available in the source material.

BUSINESS: Risk 2 Solution Group | diane.k@risk2solution.com | 1300 459 970 | Brisbane, Australia
DEECA DISTANCE ORIGIN (only for DEECA travel costs): Leopold VIC

CLIENTS & DOMAINS:
@virginaustralia.com=Virgin Australia | @vline.com.au=V/Line | @deeca.vic.gov.au=DEECA | @wayss.org.au=Wayss | @scentregroup.com=Scentregroup

SESSION TYPES & TIMES:
1. Security Awareness Recurrent: 09:00-14:00 | VA only
2. Security Awareness Initial: 08:30-16:30 | VA only
3. OVA Virgin Australia: 09:00-13:00 | VA only | ALWAYS 1 trainer | NOT the same as Initial
4. Manage Conflict & Unlawful Behaviour: TBC | V/Line
5. Authorised Officer Training: TBC | V/Line
6. Emergency Management Training: 09:00-13:00 | Wayss/DEECA
7. OVA Training: 09:00-13:00 | Wayss
8. Chief Warden/Manage Self-Resilience: 09:00-13:00 | Scentregroup
9. Personal Safety & Conflict Mgmt: 09:00-13:00 | DEECA | full_day=true

VA PARTICIPANT RULE: FC + CC = TOTAL. Never use FC or CC alone. Total >=14 = 2 trainers. Total <14 = 1 trainer. SA = always 1 trainer.

VA TRAINER PAIRS (Security Awareness):
QLD: Lawrence Phillips + Chris Walsh | 2nd: Paul Johnston + Dan Du Plessis
NSW: Mark Edmonds + Mick Haran | 2nd: Marina Toailoa + Shane Garrett | 3rd: Dave Cohen + TBC
VIC: Ross Mackenzie + Grant McDonald | 2nd: Dirk McLean + Adam Stone
WA: Andrew Chan + Paul Johnston | 2nd: Dave Cohen + Warren Kotkis
SA: Dave Cohen (always 1 trainer)

VA OVA (always 1 trainer): QLD=Lawrence Phillips | NSW=Mark Edmonds | VIC=Ross Mackenzie | WA=Andrew Chan | SA=Dave Cohen

VA AIRPORT CODES: BNE=QLD | SYD=NSW | MEL=VIC | PER=WA | ADL=SA

OTHER TRAINER PREFERENCES:
DEECA VIC: 1st Ross Mackenzie | 2nd Grant McDonald
Scentregroup: SYD=Mark Edmonds | VIC=Ross Mackenzie | QLD/WA/ACT=Paul Johnston | SA/NZ=Dave Cohen

DEECA PRICING: Course $1,996.00 inc GST ($1,814.55 ex) | Materials $36.00 inc ($32.73 ex) | Travel $0.88/km return beyond 200km total | Accommodation $207/night | Meals $128.85/night
DEECA TRAVEL: use Google Maps from Leopold VIC to venue. If >100km one-way: charge return trip beyond 200km @ $0.88/km. Include accommodation+meals. Do NOT mention Leopold in quote.

EMAIL SIGN-OFF: Kind Regards\nDiane Kruger\nCorporate Operations Lead\nRisk 2 Solution Group\n1300 459 970

OUTPUT JSON:
{
  "bookings": [{
    "action_type": "booking OR deeca_quote OR cancellation OR rescheduling OR conflict OR unknown",
    "client": "", "session_type": "", "date": "REAL DATE never TBC",
    "state": "", "venue": "", "participants": 0,
    "trainers": [], "trainer_count": 0,
    "session_start": "09:00", "session_end": "13:00",
    "full_day": false, "rules_applied": [],
    "calendar_title": "", "calendar_note": "", "flags": []
  }],
  "client_email_to": "", "client_email_subject": "", "client_email_body": "",
  "trainer_emails": [{"trainer_names": [], "trainer_email_to": [], "trainer_email_subject": "", "trainer_email_body": ""}],
  "quote_section5": "", "quote_section5_fields": {
    "TRAINER_NAME": "", "CLASS_DATES": "", "START_TIME": "09:00", "FINISH_TIME": "13:00",
    "CLASS_COST_EX": "$0.00", "CLASS_COST_INC": "$0.00",
    "MATERIALS_EX": "$0.00", "MATERIALS_INC": "$0.00",
    "TRAVEL_KM": "$0.00", "ACCOM_MEALS": "$0.00",
    "TOTAL_EX": "$0.00", "TOTAL_INC": "$0.00", "DATE_PREPARED": ""
  },
  "quote_total": 0, "diane_summary": "", "overall_flags": []
}`;
}

// ── JASMINE CHAT SYSTEM PROMPT ────────────────────────────────────────────────
const JASMINE_CHAT_SYS = `You are Jasmine, the AI scheduling agent for Risk 2 Solution Group. You help Diane and Kandia manage training bookings, assign trainers, process client requests and answer questions about the business.

BUSINESS: Risk 2 Solution Group | training@risk2solution.com | 1300 459 970 | Brisbane, Australia

You know all the clients (Virgin Australia, V/Line, DEECA, Wayss, Scentregroup), all 16 trainers, all 9 session types and all booking rules. Be helpful, concise and professional. If asked to create a booking, trainer or session, guide the user through it conversationally.`;

// ── TWO-STEP EMAIL PROCESSING ─────────────────────────────────────────────────

// Check if we already sent an email in this conversation thread
async function alreadyRepliedToThread(conversationId, token) {
  if (!conversationId) return false;
  try {
    const sent = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/mailFolders/SentItems/messages?$filter=conversationId eq '" + conversationId + "'&$top=1&$select=id",
      null, token);
    return sent && sent.value && sent.value.length > 0;
  } catch (err) {
    return false; // If check fails, process anyway
  }
}

async function processEmailWithJasmine(emailContent, images = []) {
  // Step 1: Extract facts as plain text — include images if present (e.g. VA booking table)
  let step1Content;
  if (images && images.length > 0) {
    // Multi-modal: include images so Claude can read booking tables from screenshots
    step1Content = images.map(img => ({
      type: "image",
      source: { type: "base64", media_type: img.mediaType, data: img.base64 }
    }));
    step1Content.push({
      type: "text",
      text: `Extract ALL booking details from this email and the attached image(s). The image may contain a booking schedule table with dates, locations, course types and participant numbers. Return a numbered list — one entry per session row in the table.

For each session include:
CLIENT: (e.g. Virgin Australia)
SESSION TYPE: (Recurrent FC/CC = Security Awareness Recurrent | Initial Cabin Crew = Security Awareness Initial | OVA = OVA Virgin Australia)
DATE: (exact date from table — convert to readable e.g. 10 July 2026)
START TIME: (from table e.g. 09:00)
END TIME: (from table e.g. 14:00)
LOCATION/STATE: (BNE=QLD, SYD=NSW, MEL=VIC, PER=WA, ADL=SA)
PARTICIPANTS: (FC + CC combined — add both columns together)
NOTES:

Email context:
---
${emailContent.slice(0, 4000)}`
    });
  } else {
    step1Content = `Extract ALL booking details from this email. Return a numbered list — one entry per session.

For each session include:
CLIENT: 
SESSION TYPE: 
DATE: (use exact dates — convert DD/MM/YY to readable e.g. 5 March 2026 — NEVER write TBC if a date exists)
START TIME: (09:00 default)
END TIME:
VENUE: (full address)
PARTICIPANTS:
NOTES:

If DEECA form with multiple classes in section 2.2, list EACH class separately using Preferred date 1.

Source material:
---
${emailContent.slice(0, 20000)}`;
  }

  const step1Res = await client.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 3000,
    messages: [{ role: "user", content: step1Content }]
  });

  const extractedFacts = step1Res.content[0].text || "";

  // Step 2: Apply Jasmine's rules and produce JSON
  const step2Res = await client.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 8000,
    system: JASMINE_PROC_SYS,
    messages: [{ role: "user", content: extractedFacts }]
  });

  const text = step2Res.content[0].text || "{}";
  let clean = text.replace(/```json|```/g, "").trim();
  if (!clean.startsWith("{")) throw new Error("Jasmine returned text instead of JSON: " + clean.slice(0, 120));

  const j = JSON.parse(clean);
  return { jasmineParsed: j, extractedFacts };
}

// ── INBOX POLLING ─────────────────────────────────────────────────────────────
async function pollInbox() {
  if (!AZURE_CLIENT_ID) { console.log("Graph not configured — skipping poll."); return; }
  console.log("Jasmine polling inbox...");
  try {
    const token = await getGraphToken();
    const result = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/mailFolders/inbox/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,bodyPreview,body,receivedDateTime,hasAttachments,categories,webLink",
      null, token);

    if (!result || !result.value) { console.log("No emails found."); return; }
    console.log("Inbox has " + result.value.length + " email(s).");

    // Filter out calendar invites, auto-replies, out-of-office, delivery reports
    const shouldSkip = (email) => {
      const subj = (email.subject || "").toLowerCase();
      const ct = (email.contentType || "").toLowerCase();
      if (ct.includes("calendar")) return true;
      if (subj.startsWith("accepted:")) return true;
      if (subj.startsWith("declined:")) return true;
      if (subj.startsWith("tentative:")) return true;
      if (subj.startsWith("cancelled:")) return false; // DO process cancellations
      if (subj.startsWith("automatic reply:")) return true;
      if (subj.startsWith("out of office:")) return true;
      if (subj.includes("delivery failed")) return true;
      if (subj.includes("undeliverable")) return true;
      return false;
    };

    const unprocessed = result.value.filter(email =>
      (!email.categories || !email.categories.includes("Jasmine Processed")) &&
      !shouldSkip(email)
    );
    console.log("Unprocessed: " + unprocessed.length);

    for (const email of unprocessed) {
      try {
        const senderEmail = (email.from && email.from.emailAddress ? email.from.emailAddress.address : "").toLowerCase();
        let clientName = getClientFromEmail(senderEmail);

        // For internal senders (staff forwarding emails for testing/manual submission)
        // Always process — let Jasmine identify the client from content + attachments
        if (clientName === "INTERNAL") {
          // Try subject line first (fastest)
          const subject = (email.subject || "").toLowerCase();
          let detectedClient = detectClientFromContent(subject);

          // Try body preview
          if (!detectedClient) {
            const bodyPreview = (email.bodyPreview || "") + (email.body && email.body.content ? email.body.content.replace(/<[^>]+>/g, " ") : "");
            detectedClient = detectClientFromContent(bodyPreview);
          }

          if (detectedClient) {
            console.log("Internal forward — client detected from content: " + detectedClient);
            clientName = detectedClient;
          } else {
            // Could not detect from subject/body — still process, Jasmine will read the attachment
            console.log("Internal forward — client unknown from header/body, processing anyway (attachment may contain client info)");
            clientName = "Unknown — Internal Forward";
          }
        }

        if (!clientName) {
          // Truly unknown external sender
          console.log("Unknown sender: " + senderEmail);
          emailLog.push({
            id: Date.now(),
            timestamp: new Date().toISOString(),
            subject: email.subject,
            from: senderEmail,
            client: "Unknown",
            status: "unrecognised",
            summary: "Unrecognised sender — no action taken"
          });
        } else {
          // Check if this is a reply to a thread we already handled
          const alreadyHandled = await alreadyRepliedToThread(email.conversationId, token);
          if (alreadyHandled) {
            console.log("Reply to already-handled thread — tagging without reprocessing: " + email.subject);
            emailLog.push({
              id: Date.now(),
              timestamp: new Date().toISOString(),
              subject: email.subject,
              from: senderEmail,
              client: clientName,
              status: "reply_to_handled",
              summary: "Reply to an already-processed conversation — no new action required"
            });
          } else {
            await processInboundEmail(email, clientName, token);
          }
        }

        // Tag as processed
        await graphRequest("PATCH",
          "/users/" + TRAINING_MAILBOX + "/messages/" + email.id,
          { categories: [...(email.categories || []), "Jasmine Processed"] },
          token);
        console.log("Tagged: " + email.subject);
        await new Promise(r => setTimeout(r, 500)); // Brief pause between emails
      } catch (err) {
        console.error("Error processing email:", email.subject, err.message);
      }
    }
  } catch (err) {
    console.error("Poll error:", err.message);
  }
}

async function processInboundEmail(email, clientName, token) {
  console.log("Processing for " + clientName + ": " + email.subject);

  const bodyText = email.body && email.body.content
    ? email.body.content.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()
    : email.bodyPreview || "";

  // ── Fetch attachments — Word docs, images (critical for VA inline tables) ────
  let attachmentText = "";
  let originalDocxBuffer = null;
  let originalDocxName = "";
  const inlineImages = []; // { contentType, base64 } for Claude vision

  if (email.hasAttachments) {
    try {
      const attachments = await graphRequest("GET",
        "/users/" + TRAINING_MAILBOX + "/messages/" + email.id + "/attachments",
        null, token);

      if (attachments && attachments.value) {
        for (const att of attachments.value) {
          const name = (att.name || "").toLowerCase();
          const contentType = (att.contentType || "").toLowerCase();

          if (name.endsWith(".docx") || name.endsWith(".doc")) {
            if (att.contentBytes) {
              originalDocxBuffer = Buffer.from(att.contentBytes, "base64");
              originalDocxName = att.name || "attachment.docx";
              try {
                const result = await mammoth.extractRawText({ buffer: originalDocxBuffer });
                if (result.value) {
                  attachmentText += "\n\n=== ATTACHED DOCUMENT: " + att.name + " ===\n" + result.value.trim();
                  console.log("Read Word attachment: " + att.name + " (" + result.value.length + " chars)");
                }
              } catch (err) {
                console.error("mammoth error:", err.message);
              }
            }
          } else if (contentType.startsWith("image/") || name.match(/\.(png|jpg|jpeg|gif|bmp|webp)$/)) {
            // Inline image — send to Claude vision
            if (att.contentBytes) {
              const mediaType = contentType.startsWith("image/") ? contentType.split(";")[0] : "image/png";
              inlineImages.push({ mediaType, base64: att.contentBytes });
              console.log("Found inline image: " + (att.name || "image") + " (" + mediaType + ")");
            }
          } else if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) {
            attachmentText += "\n\n=== ATTACHED FILE: " + att.name + " (Excel/CSV — extract key data) ===";
          }
        }
      }
    } catch (err) {
      console.error("Attachment fetch error:", err.message);
    }
  }

  // If client was not detected from header, try attachment text
  if ((clientName === "Unknown — Internal Forward" || clientName === "INTERNAL") && attachmentText) {
    const detected = detectClientFromContent(attachmentText);
    if (detected) {
      clientName = detected;
      console.log("Client detected from attachment: " + clientName);
    }
  }

  const emailContent = [
    "From: " + (email.from && email.from.emailAddress ? email.from.emailAddress.address : ""),
    "Subject: " + (email.subject || ""),
    "Date received: " + (email.receivedDateTime || ""),
    "Client identified as: " + clientName,
    "Body: " + bodyText.slice(0, 6000),
    attachmentText.slice(0, 16000)
  ].join("\n");

  let jasmineParsed, extractedFacts;
  try {
    const result = await processEmailWithJasmine(emailContent, inlineImages);
    jasmineParsed = result.jasmineParsed;
    extractedFacts = result.extractedFacts;
  } catch (err) {
    console.error("Jasmine processing error:", err.message);
    reviewQueue.push({
      id: Date.now(),
      timestamp: new Date().toISOString(),
      email_id: email.id,
      subject: email.subject,
      from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
      client: clientName,
      action_type: "error",
      diane_summary: "Jasmine could not process this email: " + err.message,
      status: "pending",
      drafts: [],
      raw_email: emailContent.slice(0, 2000)
    });
    return;
  }

  // Build queue entry from Jasmine's output
  const bookings = jasmineParsed.bookings || [];
  // Determine overall action type and queue tag for Diane
  const overallActionType = bookings.length > 0 ? bookings[0].action_type : (jasmineParsed.action_type || "unknown");
  const bookingTagMap = {
    "booking": { label: "Booking Request", color: "#1d4ed8", bg: "#dbeafe" },
    "deeca_quote": { label: "DEECA Quote", color: "#15803d", bg: "#dcfce7" },
    "cancellation": { label: "Cancellation", color: "#b91c1c", bg: "#fef2f2" },
    "rescheduling": { label: "Rescheduling", color: "#d97706", bg: "#fef9c3" },
    "conflict": { label: "Conflict", color: "#7c3aed", bg: "#f3e8ff" },
    "unknown": { label: "Needs Review", color: "#64748b", bg: "#f1f5f9" },
    "error": { label: "Error", color: "#b91c1c", bg: "#fef2f2" }
  };
  const bookingTag = bookingTagMap[overallActionType] || bookingTagMap["unknown"];

  const entry = {
    id: Date.now(),
    timestamp: new Date().toISOString(),
    email_id: email.id,
    subject: email.subject,
    from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
    client: clientName,
    action_type: overallActionType,
    booking_tag: bookingTag,
    bookings,
    client_email_to: jasmineParsed.client_email_to || "",
    client_email_subject: jasmineParsed.client_email_subject || "",
    client_email_body: jasmineParsed.client_email_body || "",
    trainer_emails: jasmineParsed.trainer_emails || [],
    quote_section5: jasmineParsed.quote_section5 || "",
    quote_section5_fields: jasmineParsed.quote_section5_fields || {},
    quote_total: jasmineParsed.quote_total || 0,
    diane_summary: jasmineParsed.diane_summary || "",
    overall_flags: jasmineParsed.overall_flags || [],
    extracted_facts: extractedFacts,
    original_docx: originalDocxBuffer,       // original client Word doc binary
    original_docx_name: originalDocxName,    // original filename
    status: "pending",
    drafts: []
  };

  // Create Outlook draft emails (in training@risk2solution.com drafts)
  if (token && entry.client_email_body && entry.action_type !== "deeca_quote") {
    try {
      // Client confirmation draft
      const clientDraft = await graphRequest("POST",
        "/users/" + TRAINING_MAILBOX + "/messages",
        {
          subject: entry.client_email_subject || "Training Booking Confirmation",
          toRecipients: [{ emailAddress: { address: entry.client_email_to || entry.from_email } }],
          body: { contentType: "HTML", content: "<p>" + entry.client_email_body.replace(/\n/g, "<br>") + "</p>" },
          isDraft: true
        }, token);

      if (clientDraft && clientDraft.id) {
        entry.drafts.push({
          type: "client",
          draft_id: clientDraft.id,
          to: entry.client_email_to,
          subject: entry.client_email_subject,
          description: "Client confirmation email"
        });
        console.log("Client draft created.");
      }

      // Trainer draft(s)
      for (const te of (entry.trainer_emails || [])) {
        if (!te.trainer_email_body) continue;
        const trainerDraft = await graphRequest("POST",
          "/users/" + TRAINING_MAILBOX + "/messages",
          {
            subject: te.trainer_email_subject || "Session Booking — " + clientName,
            toRecipients: (te.trainer_email_to || []).map(addr => ({ emailAddress: { address: addr } })),
            body: { contentType: "HTML", content: "<p>" + te.trainer_email_body.replace(/\n/g, "<br>") + "</p>" },
            isDraft: true
          }, token);

        if (trainerDraft && trainerDraft.id) {
          entry.drafts.push({
            type: "trainer",
            draft_id: trainerDraft.id,
            to: (te.trainer_email_to || []).join(", "),
            subject: te.trainer_email_subject,
            description: "Trainer confirmation — " + (te.trainer_names || []).join(", ")
          });
          console.log("Trainer draft created for " + (te.trainer_names || []).join(", "));
        }
      }
    } catch (err) {
      console.error("Draft creation error:", err.message);
      entry.drafts_error = err.message;
    }
  }

  // Log to bookings
  for (const bk of bookings) {
    bookingsLog.push({
      ...bk,
      id: Date.now() + Math.floor(Math.random() * 9999),
      queue_id: entry.id,
      status: bk.action_type === "deeca_quote" ? "Tentative — Pending Quote Approval" : "Pending Approval",
      created: new Date().toISOString()
    });
  }

  emailLog.push({
    id: Date.now(),
    timestamp: new Date().toISOString(),
    subject: email.subject,
    from: entry.from_email,
    client: clientName,
    status: entry.action_type,
    summary: entry.diane_summary,
    booking_count: bookings.length
  });

  reviewQueue.push(entry);
  console.log("Queue entry added for " + clientName + ". Drafts: " + entry.drafts.length);
}

// ── APPROVE BOOKING — creates calendar event + sends drafts ──────────────────
async function approveQueueEntry(entry) {
  console.log("Approving entry:", entry.id, "| Bookings:", (entry.bookings||[]).length, "| Drafts:", (entry.drafts||[]).length);
  const token = await getGraphToken();
  const results = { calendar: [], drafts: [], errors: [] };

  for (const bk of (entry.bookings || [])) {
    if (!bk.date || bk.date.toLowerCase() === "tbc") {
      console.log("Skipping booking — date is TBC:", bk.session_type);
      results.errors.push("Skipped " + bk.session_type + " — date is TBC");
      continue;
    }

    // Build calendar event datetime
    const datePart = bk.date;
    const startTime = bk.session_start || "09:00";
    const endTime = bk.session_end || "13:00";
    console.log("Creating calendar event:", datePart, startTime, "-", endTime, bk.session_type);

    // Parse date — try multiple formats
    let startDT, endDT;
    try {
      // Try "5 March 2026 09:00" format
      let d = new Date(datePart + " " + startTime);
      let e = new Date(datePart + " " + endTime);
      // If that fails try adding year context
      if (isNaN(d.getTime())) {
        d = new Date(datePart.replace(/(\d+)\s+(\w+)\s+(\d+)/, "$2 $1 $3") + " " + startTime);
        e = new Date(datePart.replace(/(\d+)\s+(\w+)\s+(\d+)/, "$2 $1 $3") + " " + endTime);
      }
      if (isNaN(d.getTime())) throw new Error("Could not parse date: " + datePart + " " + startTime);
      startDT = d.toISOString().slice(0, 19);
      endDT = e.toISOString().slice(0, 19);
      console.log("Parsed datetime:", startDT, "to", endDT);
    } catch (err) {
      console.error("Date parse error:", err.message);
      results.errors.push("Date parse failed for " + bk.session_type + ": " + err.message);
      continue;
    }

    const isDeeca = bk.action_type === "deeca_quote";
    const isFullDay = bk.full_day || false;

    // Attendees — trainer emails (if we have them)
    const trainerAttendees = [];
    for (const te of (entry.trainer_emails || [])) {
      for (const addr of (te.trainer_email_to || [])) {
        if (addr && addr.includes("@")) {
          trainerAttendees.push({ emailAddress: { address: addr }, type: "required" });
        }
      }
    }

    const calendarEvent = {
      subject: bk.calendar_title || (entry.client + " — " + bk.session_type),
      body: {
        contentType: "HTML",
        content: "<p><strong>Client:</strong> " + entry.client + "</p>" +
          "<p><strong>Session:</strong> " + bk.session_type + "</p>" +
          "<p><strong>Trainers:</strong> " + (bk.trainers || []).join(", ") + "</p>" +
          "<p><strong>Venue:</strong> " + (bk.venue || "TBC") + "</p>" +
          "<p><strong>Participants:</strong> " + (bk.participants || "TBC") + "</p>" +
          (bk.calendar_note ? "<p><strong>Notes:</strong> " + bk.calendar_note + "</p>" : "") +
          "<p><em>Booked by Jasmine — Risk 2 Solution Group</em></p>"
      },
      start: isFullDay
        ? { date: new Date(datePart).toISOString().split("T")[0], timeZone: "Australia/Brisbane" }
        : { dateTime: startDT, timeZone: "Australia/Brisbane" },
      end: isFullDay
        ? { date: new Date(new Date(datePart).getTime() + 86400000).toISOString().split("T")[0], timeZone: "Australia/Brisbane" }
        : { dateTime: endDT, timeZone: "Australia/Brisbane" },
      isAllDay: isFullDay,
      showAs: isDeeca ? "tentative" : "busy",
      location: { displayName: bk.venue || "" },
      attendees: trainerAttendees
    };

    try {
      const created = await graphRequest("POST",
        "/users/" + TRAINING_MAILBOX + "/calendar/events",
        calendarEvent, token);
      if (created && created.id) {
        console.log("✓ Calendar event created:", calendarEvent.subject, "ID:", created.id);
        entry.calendar_event_id = created.id;
        results.calendar.push(calendarEvent.subject);
      } else if (created && created.error) {
        console.error("Calendar API error:", JSON.stringify(created.error));
        results.errors.push("Calendar error: " + JSON.stringify(created.error));
      } else {
        console.warn("Calendar: unexpected response:", JSON.stringify(created));
      }
    } catch (err) {
      console.error("Calendar exception:", err.message);
      results.errors.push("Calendar exception: " + err.message);
    }
  }

  // Create Outlook drafts if not already created during processing
  if (entry.drafts && entry.drafts.length > 0) {
    console.log("Drafts already created during processing:", entry.drafts.length, "drafts ready in Outlook");
  } else {
    // Drafts were not created during processing — create them now on approve
    console.log("No drafts found — creating now on approve...");
    if (entry.client_email_body && entry.client_email_to) {
      try {
        const clientDraft = await graphRequest("POST",
          "/users/" + TRAINING_MAILBOX + "/messages",
          {
            subject: entry.client_email_subject || "Training Booking Confirmation",
            toRecipients: [{ emailAddress: { address: entry.client_email_to } }],
            body: { contentType: "HTML", content: "<p>" + (entry.client_email_body||"").split("\n").join("<br>") + "</p>" },
            isDraft: true
          }, token);
        if (clientDraft && clientDraft.id) {
          console.log("✓ Client draft created on approve:", clientDraft.id);
          results.drafts.push("Client email to " + entry.client_email_to);
          if (!entry.drafts) entry.drafts = [];
          entry.drafts.push({ type: "client", draft_id: clientDraft.id, to: entry.client_email_to, description: "Client confirmation" });
        } else {
          console.error("Client draft failed:", JSON.stringify(clientDraft));
          results.errors.push("Client draft failed: " + JSON.stringify(clientDraft));
        }
      } catch (err) {
        console.error("Client draft exception:", err.message);
        results.errors.push("Client draft error: " + err.message);
      }
    }

    for (const te of (entry.trainer_emails || [])) {
      if (!te.trainer_email_body || !te.trainer_email_to || te.trainer_email_to.length === 0) continue;
      try {
        const trDraft = await graphRequest("POST",
          "/users/" + TRAINING_MAILBOX + "/messages",
          {
            subject: te.trainer_email_subject || "Session Booking",
            toRecipients: te.trainer_email_to.map(a => ({ emailAddress: { address: a } })),
            body: { contentType: "HTML", content: "<p>" + (te.trainer_email_body||"").split("\n").join("<br>") + "</p>" },
            isDraft: true
          }, token);
        if (trDraft && trDraft.id) {
          console.log("✓ Trainer draft created:", trDraft.id);
          results.drafts.push("Trainer email to " + te.trainer_email_to.join(", "));
        } else {
          console.error("Trainer draft failed:", JSON.stringify(trDraft));
          results.errors.push("Trainer draft failed: " + JSON.stringify(trDraft));
        }
      } catch (err) {
        console.error("Trainer draft exception:", err.message);
        results.errors.push("Trainer draft error: " + err.message);
      }
    }
  }

  entry.status = "approved";
  entry.approved_at = new Date().toISOString();
  entry.approve_results = results;

  for (const bk of bookingsLog) {
    if (bk.queue_id === entry.id) bk.status = "Confirmed";
  }

  console.log("Entry", entry.id, "approved. Calendar:", results.calendar.length, "Drafts:", results.drafts.length, "Errors:", results.errors.length);
  if (results.errors.length > 0) console.error("Approval errors:", results.errors);
}

// ── API ROUTES ────────────────────────────────────────────────────────────────

app.post("/api/login", (req, res) => {
  const { password } = req.body;
  if (password === PASSWORD) { req.session.authenticated = true; res.json({ success: true }); }
  else res.status(401).json({ error: "Incorrect password" });
});

app.post("/api/logout", (req, res) => { req.session.destroy(); res.json({ success: true }); });

app.get("/api/check-auth", (req, res) => {
  res.json({ authenticated: !!(req.session && req.session.authenticated) });
});



// ── REVIEW QUEUE ──────────────────────────────────────────────────────────────
app.get("/api/review-queue", requireAuth, (req, res) => {
  res.json(reviewQueue.filter(e => e.status === "pending").reverse());
});

app.get("/api/review-queue/count", requireAuth, (req, res) => {
  res.json({ count: reviewQueue.filter(e => e.status === "pending").length });
});

app.post("/api/review-queue/:id/approve", requireAuth, async (req, res) => {
  const entry = reviewQueue.find(e => e.id === parseInt(req.params.id));
  if (!entry) return res.status(404).json({ error: "Not found" });
  try {
    await approveQueueEntry(entry);
    res.json({ success: true, drafts: entry.drafts, calendar_event_id: entry.calendar_event_id });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/review-queue/:id/reject", requireAuth, (req, res) => {
  const entry = reviewQueue.find(e => e.id === parseInt(req.params.id));
  if (!entry) return res.status(404).json({ error: "Not found" });
  entry.status = "rejected";
  entry.rejected_at = new Date().toISOString();
  entry.reject_reason = req.body.reason || "";
  for (const bk of bookingsLog) {
    if (bk.queue_id === entry.id) bk.status = "Rejected";
  }
  res.json({ success: true });
});

app.post("/api/review-queue/:id/dismiss", requireAuth, (req, res) => {
  const idx = reviewQueue.findIndex(e => e.id === parseInt(req.params.id));
  if (idx === -1) return res.status(404).json({ error: "Not found" });
  reviewQueue[idx].status = "dismissed";
  reviewQueue[idx].dismissed_at = new Date().toISOString();
  res.json({ success: true });
});

// ── POLL NOW ──────────────────────────────────────────────────────────────────
app.post("/api/poll-now", requireAuth, async (req, res) => {
  try { await pollInbox(); res.json({ success: true }); }
  catch (err) { res.status(500).json({ error: err.message }); }
});

// Force reprocess — reads ALL emails ignoring the Jasmine Processed tag
// Use this when testing or if something was missed
app.post("/api/reprocess-all", requireAuth, async (req, res) => {
  try {
    const token = await getGraphToken();
    const result = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/mailFolders/inbox/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,bodyPreview,body,receivedDateTime,hasAttachments,categories",
      null, token);
    if (!result || !result.value) return res.json({ success: true, processed: 0 });
    let count = 0;
    for (const email of result.value) {
      const senderEmail = (email.from && email.from.emailAddress ? email.from.emailAddress.address : "").toLowerCase();
      let clientName = getClientFromEmail(senderEmail);
      if (clientName === "INTERNAL") {
        const bodyText = (email.bodyPreview || "") + (email.body && email.body.content ? email.body.content : "");
        clientName = detectClientFromContent(bodyText);
      }
      if (clientName === "INTERNAL") {
        const subject = (email.subject || "").toLowerCase();
        const bodyText = (email.bodyPreview || "");
        clientName = detectClientFromContent(subject) || detectClientFromContent(bodyText) || "Unknown — Internal Forward";
      }
      if (clientName) {
        await processInboundEmail(email, clientName, token);
        count++;
      }
    }
    res.json({ success: true, processed: count });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── BOOKINGS ──────────────────────────────────────────────────────────────────
app.get("/api/bookings", requireAuth, (req, res) => {
  res.json([...bookingsLog].reverse());
});

// ── EMAIL LOG ─────────────────────────────────────────────────────────────────
app.get("/api/email-log", requireAuth, (req, res) => {
  res.json([...emailLog].reverse());
});

// ── MANUAL PROCESS EMAIL (from Process Email tab) ─────────────────────────────
app.post("/api/process-email", requireAuth, async (req, res) => {
  const { content, images } = req.body;
  if (!content || !content.trim()) return res.status(400).json({ error: "No content provided" });
  try {
    const { jasmineParsed, extractedFacts } = await processEmailWithJasmine(content, images || []);
    res.json({ success: true, result: jasmineParsed, extracted_facts: extractedFacts });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── CHAT ──────────────────────────────────────────────────────────────────────
app.post("/api/chat", requireAuth, async (req, res) => {
  const { messages } = req.body;
  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      system: JASMINE_CHAT_SYS,
      messages
    });
    res.json({ message: response.content[0].text });
  } catch (err) {
    res.status(500).json({ error: "Something went wrong: " + err.message });
  }
});

// ── M365 STATUS ───────────────────────────────────────────────────────────────
app.get("/api/m365-status", requireAuth, async (req, res) => {
  if (!AZURE_CLIENT_ID) return res.json({ connected: false, reason: "Azure not configured" });
  try {
    await getGraphToken();
    res.json({ connected: true, mailbox: TRAINING_MAILBOX });
  } catch (err) {
    res.json({ connected: false, reason: err.message });
  }
});

// ── GENERATE SECTION 5 WORD DOCUMENT ─────────────────────────────────────────
app.get("/api/generate-section5/:id", requireAuth, (req, res) => {
  const entry = reviewQueue.find(e => e.id === parseInt(req.params.id));
  if (!entry) return res.status(404).json({ error: "Queue entry not found" });

  const fields = entry.quote_section5_fields;
  if (!fields || Object.keys(fields).length === 0) {
    return res.status(400).json({ error: "No quote fields available for this entry. Make sure Jasmine processed a DEECA email with a booking form attached." });
  }

  // Prefer the original client document — fall back to template if not available
  let docxContent;
  let outputFilename;

  if (entry.original_docx && entry.original_docx.length > 0) {
    // Use the original client Word document
    console.log("Using original client document: " + entry.original_docx_name);
    docxContent = entry.original_docx.toString("binary");
    outputFilename = (entry.original_docx_name || "DEECA_Request").replace(".docx", "") + "_Completed.docx";
  } else {
    // Fall back to standalone template
    console.log("No original document found — using standalone template");
    const templatePath = path.join(__dirname, "templates", "DEECA_Section5_Template.docx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "No original DEECA document was attached to the email, and the fallback template was not found on the server." });
    }
    docxContent = fs.readFileSync(templatePath, "binary");
    outputFilename = "DEECA_Section5_Completed.docx";
  }

  try {
    const zip = new PizZip(docxContent);
    let xmlStr = zip.file("word/document.xml").asText();

    // Find the Request for Service table by its heading text
    const tableStart = xmlStr.lastIndexOf("<w:tbl>", xmlStr.indexOf("Request for Service"));
    const tableEnd = xmlStr.indexOf("</w:tbl>", tableStart) + 8;
    let tableXml = xmlStr.slice(tableStart, tableEnd);

    // Extract all rows from the table
    const rowMatches = [...tableXml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)];

    // Helper: set text in a specific cell of a row
    function setCellText(rowXml, cellIndex, newText) {
      const cells = [...rowXml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)];
      if (cellIndex >= cells.length) return rowXml;
      const cell = cells[cellIndex];
      // Find the paragraph in this cell and replace all runs with a single new run
      const cellXml = cell[0];
      // Preserve cell properties (shading, borders, width)
      const tcPr = cellXml.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/) || [""];
      const rPr = cellXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""];
      const newCellXml = "<w:tc>" + tcPr[0] +
        "<w:p><w:r>" + rPr[0] +
        "<w:t xml:space=\"preserve\">" + newText.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;") + "</w:t></w:r></w:p></w:tc>";
      return rowXml.slice(0, cell.index) + newCellXml + rowXml.slice(cell.index + cell[0].length);
    }

    // Values to fill in
    const today = new Date().toLocaleDateString("en-AU", { day: "numeric", month: "long", year: "numeric" });

    // Format CLASS_DATES as: Class 1 – Venue – DD/MM/YY (one per line, dates only — no full address)
    // Jasmine returns CLASS_DATES as newline-separated readable dates
    // We format them simply as Class 1 – Date, Class 2 – Date etc.
    const rawDates = (fields.CLASS_DATES || "").split(/\n|,/).map(d => d.trim()).filter(Boolean);
    const formattedDates = rawDates.map((d, i) => "Class " + (i+1) + " – " + d).join("\n");

    const fillMap = {
      1: { 1: "Risk 2 Solution Pty Ltd" },
      2: { 1: "training@risk2solution.com", 2: "Ph: 1300 459 970" },
      3: { 1: "Personal Safety & Conflict Management" },
      4: { 1: fields.TRAINER_NAME || "" },
      5: { 1: formattedDates || fields.CLASS_DATES || "" },
      6: { 1: "Start time: " + (fields.START_TIME || "09:00"), 2: "Finish time: " + (fields.FINISH_TIME || "13:00") },
      10: { 1: fields.CLASS_COST_EX || "$0.00", 2: fields.CLASS_COST_INC || "$0.00" },
      11: { 1: fields.MATERIALS_EX || "$0.00", 2: fields.MATERIALS_INC || "$0.00" },
      13: { 2: fields.TRAVEL_KM || "$0.00" },
      14: { 2: fields.ACCOM_MEALS || "$0.00" },
      15: { 1: fields.TOTAL_EX || "$0.00", 2: fields.TOTAL_INC || "$0.00" }
    };

    // Apply fills row by row
    let newTableXml = tableXml;
    const updatedRows = [...rowMatches];

    for (const [rowIdxStr, cellFills] of Object.entries(fillMap)) {
      const rowIdx = parseInt(rowIdxStr);
      if (rowIdx >= updatedRows.length) continue;
      let rowXml = updatedRows[rowIdx][0];
      for (const [cellIdxStr, value] of Object.entries(cellFills)) {
        rowXml = setCellText(rowXml, parseInt(cellIdxStr), value);
      }
      updatedRows[rowIdx] = { 0: rowXml, index: updatedRows[rowIdx].index };
    }

    // Fix YES/NO row (row 16)
    if (updatedRows[16]) {
      updatedRows[16][0] = updatedRows[16][0].replace(">YES/NO<", ">YES<");
    }

    // Rebuild table with updated rows
    let rebuilt = tableXml;
    // Replace rows from last to first to preserve indices
    for (let i = rowMatches.length - 1; i >= 0; i--) {
      const orig = rowMatches[i];
      rebuilt = rebuilt.slice(0, orig.index) + updatedRows[i][0] + rebuilt.slice(orig.index + orig[0].length);
    }

    // Replace table in document
    xmlStr = xmlStr.slice(0, tableStart) + rebuilt + xmlStr.slice(tableEnd);

    // Add prepared date at end of document before </w:body>
    const prepNote = "<w:p><w:r><w:rPr><w:sz w:val=\"16\"/><w:szCs w:val=\"16\"/><w:i/><w:color w:val=\"666666\"/></w:rPr><w:t xml:space=\"preserve\">Section completed by Jasmine (Risk 2 Solution Group Scheduling Agent) on " + today + "</w:t></w:r></w:p>";
    xmlStr = xmlStr.replace("</w:body>", prepNote + "</w:body>");

    zip.file("word/document.xml", xmlStr);
    const buf = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="' + outputFilename + '"');
    res.send(buf);
    console.log("Section completed and document generated for entry", entry.id);
  } catch (err) {
    console.error("Document generation error:", err);
    res.status(500).json({ error: "Could not generate document: " + err.message });
  }
});

// ── CATCH-ALL — serve index.html for any non-API route ───────────────────────
// Must be LAST — after all API routes are defined
app.get("*", (req, res) => {
  if (!req.path.startsWith("/api")) {
    res.sendFile(path.join(__dirname, "public", "index.html"));
  } else {
    res.status(404).json({ error: "Route not found: " + req.path });
  }
});

// ── START ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log("Jasmine is running on port " + PORT);
  if (AZURE_CLIENT_ID) {
    pollInbox();
    setInterval(pollInbox, 10 * 60 * 1000); // Poll every 10 minutes
    console.log("Email polling started — every 10 minutes.");
  } else {
    console.log("Azure not configured — email polling disabled. Set AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET to enable.");
  }
});
