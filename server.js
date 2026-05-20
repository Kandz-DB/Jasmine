const express = require("express");
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
    "action_type": "booking OR deeca_quote OR conflict OR unknown",
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
async function processEmailWithJasmine(emailContent) {
  // Step 1: Extract facts as plain text
  const step1Res = await client.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 3000,
    messages: [{
      role: "user",
      content: `Extract ALL booking details from this email. Return a numbered list — one entry per session.

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
${emailContent.slice(0, 20000)}`
    }]
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

    const unprocessed = result.value.filter(email =>
      !email.categories || !email.categories.includes("Jasmine Processed")
    );
    console.log("Unprocessed: " + unprocessed.length);

    for (const email of unprocessed) {
      try {
        const senderEmail = (email.from && email.from.emailAddress ? email.from.emailAddress.address : "").toLowerCase();
        const clientName = getClientFromEmail(senderEmail);

        if (!clientName) {
          // Unknown sender — log it but don't process
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
          await processInboundEmail(email, clientName, token);
        }

        // Tag as processed
        await graphRequest("PATCH",
          "/users/" + TRAINING_MAILBOX + "/messages/" + email.id,
          { categories: [...(email.categories || []), "Jasmine Processed"] },
          token);
        console.log("Tagged: " + email.subject);
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

  const emailContent = [
    "From: " + (email.from && email.from.emailAddress ? email.from.emailAddress.address : ""),
    "Subject: " + (email.subject || ""),
    "Date received: " + (email.receivedDateTime || ""),
    "Body: " + bodyText.slice(0, 8000)
  ].join("\n");

  let jasmineParsed, extractedFacts;
  try {
    const result = await processEmailWithJasmine(emailContent);
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
  const entry = {
    id: Date.now(),
    timestamp: new Date().toISOString(),
    email_id: email.id,
    subject: email.subject,
    from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
    client: clientName,
    action_type: bookings.length > 0 ? bookings[0].action_type : "unknown",
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
  const token = await getGraphToken();

  for (const bk of (entry.bookings || [])) {
    if (!bk.date || bk.date.toLowerCase() === "tbc") continue;

    // Build calendar event datetime
    const datePart = bk.date; // e.g. "5 March 2026"
    const startTime = bk.session_start || "09:00";
    const endTime = bk.session_end || "13:00";

    // Parse date for Graph API (needs ISO format)
    let startDT, endDT;
    try {
      const d = new Date(datePart + " " + startTime);
      const e = new Date(datePart + " " + endTime);
      if (isNaN(d.getTime())) throw new Error("Invalid date: " + datePart);
      startDT = d.toISOString().replace(/\.\d{3}Z$/, "");
      endDT = e.toISOString().replace(/\.\d{3}Z$/, "");
    } catch (err) {
      console.error("Date parse error for calendar:", err.message);
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
        console.log("Calendar event created: " + calendarEvent.subject);
        entry.calendar_event_id = created.id;
      }
    } catch (err) {
      console.error("Calendar error:", err.message);
    }
  }

  // Move drafts from Drafts to Sent (or just log that they exist)
  // Drafts are already in training@risk2solution.com — Diane just needs to review and send
  // We don't auto-send — Diane clicks Send in Outlook

  entry.status = "approved";
  entry.approved_at = new Date().toISOString();

  // Update bookings log status
  for (const bk of bookingsLog) {
    if (bk.queue_id === entry.id) {
      bk.status = "Confirmed";
    }
  }

  console.log("Entry " + entry.id + " approved.");
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

// Serve index.html for all non-API routes
app.get("*", (req, res) => {
  if (!req.path.startsWith("/api")) {
    res.sendFile(path.join(__dirname, "public", "index.html"));
  }
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

// ── POLL NOW ──────────────────────────────────────────────────────────────────
app.post("/api/poll-now", requireAuth, async (req, res) => {
  try { await pollInbox(); res.json({ success: true }); }
  catch (err) { res.status(500).json({ error: err.message }); }
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
  const { content } = req.body;
  if (!content || !content.trim()) return res.status(400).json({ error: "No content provided" });
  try {
    const { jasmineParsed, extractedFacts } = await processEmailWithJasmine(content);
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
