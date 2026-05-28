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

// ── TEST MODE ─────────────────────────────────────────────────────────────────
// When TEST_MODE=true, ALL outgoing emails (client + trainer) and calendar
// attendees are redirected to TEST_EMAIL (defaults to the training mailbox).
// No real trainer or client emails are ever touched during testing.
// Set TEST_MODE=false (or remove it) once you are happy with the flow and
// are ready to supply real trainer email addresses.
const TEST_MODE  = (process.env.TEST_MODE || "true").toLowerCase() !== "false";
const TEST_EMAIL = process.env.TEST_EMAIL || TRAINING_MAILBOX;

// Wrap every outgoing address through this function.
// In TEST_MODE it returns the internal test mailbox; in production it returns
// the real address as-is.
function safeEmail(realAddress) {
  if (TEST_MODE) return TEST_EMAIL;
  return realAddress;
}

if (TEST_MODE) {
  console.log("⚠  TEST MODE ACTIVE — all emails/invites will be sent to: " + TEST_EMAIL);
  console.log("   Set TEST_MODE=false in your environment to use real addresses.");
}

// ── IN-MEMORY STORES + FILE PERSISTENCE ──────────────────────────────────────
// Data lives in memory for speed and survives restarts via jasmine_store.json.
// Atomic write (tmp → rename) means a crash mid-save never corrupts the file.
const reviewQueue = [];
const bookingsLog = [];
const emailLog    = [];

const DATA_DIR  = path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR,  "jasmine_store.json");

function persistLoad() {
  try {
    if (!fs.existsSync(DATA_FILE)) {
      console.log("No existing data file — starting fresh.");
      return;
    }
    const data = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    // Restore reviewQueue — convert stored base64 docx back to Buffers
    for (const e of (data.reviewQueue || [])) {
      if (e._docx_b64) {
        e.original_docx = Buffer.from(e._docx_b64, "base64");
        delete e._docx_b64;
      }
      reviewQueue.push(e);
    }
    bookingsLog.push(...(data.bookingsLog || []));
    emailLog.push(...(data.emailLog || []));
    console.log("Loaded persisted data — queue: " + reviewQueue.length +
      ", bookings: " + bookingsLog.length + ", emails: " + emailLog.length);
  } catch (err) {
    console.error("persistLoad error:", err.message);
  }
}

let _saveTimer = null;
function persistSave() {
  // Debounce: coalesce rapid back-to-back writes into one disk write after 500ms
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(() => {
    try {
      if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
      const snapshot = {
        reviewQueue: reviewQueue.map(e => {
          const copy = { ...e };
          // Binary docx buffers can't go into JSON — store as base64 instead
          if (Buffer.isBuffer(copy.original_docx)) {
            copy._docx_b64 = copy.original_docx.toString("base64");
            delete copy.original_docx;
          }
          return copy;
        }),
        bookingsLog: [...bookingsLog],
        emailLog:    [...emailLog]
      };
      // Atomic write: write to .tmp first, then rename over the real file.
      // If the process dies mid-write, the old file is still intact.
      const tmp = DATA_FILE + ".tmp";
      fs.writeFileSync(tmp, JSON.stringify(snapshot, null, 2));
      fs.renameSync(tmp, DATA_FILE);
    } catch (err) {
      console.error("persistSave error:", err.message);
    }
  }, 500);
}

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
// ── CALENDAR SLOT CHECK ───────────────────────────────────────────────────────
// Checks the training mailbox calendar for:
//   (a) confirmed/busy events overlapping the time slot
//   (b) any event on the same day that mentions one of the given trainer names
//       — catches cases where the same trainer is already booked for a different
//         session on the same day even if the clock times differ
// Returns { available: bool, conflicts: [], trainerConflicts: [] }
async function checkCalendarSlot(startDT, endDT, token, trainerNames = []) {
  try {
    const t = token || await getGraphToken();

    // (a) Check the exact time slot for busy events
    const slotResult = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/calendarView" +
      "?startDateTime=" + encodeURIComponent(startDT) +
      "&endDateTime=" + encodeURIComponent(endDT) +
      "&$select=subject,start,end,showAs&$top=20",
      null, t
    );
    const timeConflicts = (slotResult && slotResult.value || []).filter(e =>
      e.showAs === "busy" || e.showAs === "oof" || e.showAs === "workingElsewhere"
    );

    // (b) If trainers provided, check the full calendar day for same-trainer clashes
    let trainerConflicts = [];
    if (trainerNames.length > 0) {
      // Widen the window to the full day (midnight-to-midnight) of the startDT date
      const dayStart = startDT.slice(0, 10) + "T00:00:00";
      const dayEnd   = startDT.slice(0, 10) + "T23:59:59";
      const dayResult = await graphRequest("GET",
        "/users/" + TRAINING_MAILBOX + "/calendarView" +
        "?startDateTime=" + encodeURIComponent(dayStart) +
        "&endDateTime=" + encodeURIComponent(dayEnd) +
        "&$select=subject,body,start,end,showAs&$top=20",
        null, t
      );
      const dayEvents = (dayResult && dayResult.value || []).filter(e =>
        e.showAs === "busy" || e.showAs === "tentative"
      );
      for (const ev of dayEvents) {
        const haystack = ((ev.subject || "") + " " +
          (ev.body && ev.body.content ? ev.body.content.replace(/<[^>]+>/g, " ") : "")).toLowerCase();
        const clashingTrainer = trainerNames.find(n => n && haystack.includes(n.toLowerCase()));
        if (clashingTrainer) {
          trainerConflicts.push({ trainer: clashingTrainer, event: ev.subject, date: startDT.slice(0, 10) });
        }
      }
    }

    return {
      available: timeConflicts.length === 0 && trainerConflicts.length === 0,
      conflicts: timeConflicts,
      trainerConflicts
    };
  } catch (err) {
    console.error("Calendar slot check error:", err.message);
    return { available: true, conflicts: [], trainerConflicts: [] };
  }
}

// ── TRAINER AVAILABILITY — find open weekdays in next N days ─────────────────
// Used to suggest alternative dates in conflict client emails.
// Returns up to 5 human-readable weekday dates where the named trainers are free.
async function getTrainerAvailableDates(trainerNames, token, daysAhead = 35) {
  try {
    const t = token || await getGraphToken();
    const fromDT = new Date().toISOString().slice(0, 19);
    const toDT   = new Date(Date.now() + daysAhead * 86400000).toISOString().slice(0, 19);
    const result = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/calendarView" +
      "?startDateTime=" + encodeURIComponent(fromDT) +
      "&endDateTime=" + encodeURIComponent(toDT) +
      "&$select=subject,start,showAs&$top=200",
      null, t
    );
    const events = (result && result.value) || [];
    // Collect dates that already have a trainer-specific booking
    const busyDates = new Set();
    for (const ev of events) {
      if (ev.showAs === "busy" || ev.showAs === "tentative" || ev.showAs === "oof") {
        const dateStr = (ev.start.dateTime || ev.start.date || "").slice(0, 10);
        if (trainerNames.some(n => n && (ev.subject || "").toLowerCase().includes(n.toLowerCase()))) {
          busyDates.add(dateStr);
        }
      }
    }
    // Walk forward day-by-day and collect free weekdays
    const available = [];
    for (let i = 1; i <= daysAhead && available.length < 5; i++) {
      const day  = new Date(Date.now() + i * 86400000);
      const dow  = day.getUTCDay();
      if (dow === 0 || dow === 6) continue;  // skip weekends
      const dateStr = day.toISOString().slice(0, 10);
      if (!busyDates.has(dateStr)) {
        available.push(day.toLocaleDateString("en-AU", {
          weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "Australia/Brisbane"
        }));
      }
    }
    return available;
  } catch (err) {
    console.error("getTrainerAvailableDates error:", err.message);
    return [];
  }
}

// ── CALENDAR PRE-BOOKING ──────────────────────────────────────────────────────
// After Jasmine parses an email:
// 1. Checks the training calendar for each booking slot
// 2. If free → creates a TENTATIVE calendar event immediately
// 3. Rewrites trainer emails from "please confirm availability" → "you have been booked"
// 4. Flags any conflicts for Diane
// On Diane's approval, tentative events are upgraded to confirmed (busy).
// Read-only calendar availability check — no events created, no emails sent.
// Called during email processing so Jasmine can flag conflicts for Diane.
async function checkCalendarAvailability(entry, jasmineParsed, token) {
  const results = { calendar_events: [], conflicts: [] };
  const bookings = jasmineParsed.bookings || [];

  for (const bk of bookings) {
    if (!bk.date || bk.date.toLowerCase() === "tbc") {
      console.log("Skipping calendar check — date is TBC for:", bk.session_type);
      continue;
    }

    // Parse dates to ISO
    let startDT, endDT;
    try {
      const d = new Date(bk.date + " " + (bk.session_start || "09:00"));
      const e = new Date(bk.date + " " + (bk.session_end || "13:00"));
      if (isNaN(d.getTime())) { console.warn("Could not parse date for calendar check:", bk.date); continue; }
      startDT = d.toISOString().slice(0, 19);
      endDT   = e.toISOString().slice(0, 19);
    } catch (err) { console.error("Date parse error in calendar check:", err.message); continue; }

    // Collect trainer names and emails for conflict checking and later approval
    const trainerNamesForCheck = bk.trainers || [];
    const trainerEmailsForApprove = [];
    for (const te of (jasmineParsed.trainer_emails || [])) {
      for (const addr of (te.trainer_email_to || [])) {
        if (addr && addr.includes("@")) trainerEmailsForApprove.push(addr);
      }
    }

    // Check slot AND same-day trainer conflicts in the training calendar
    const slotCheck = await checkCalendarSlot(startDT, endDT, token, trainerNamesForCheck);

    if (!slotCheck.available) {
      bk.calendar_conflict = true;
      const timeMsg    = slotCheck.conflicts.map(c => c.subject || "existing booking").join(", ");
      const trainerMsg = slotCheck.trainerConflicts.map(c => c.trainer + " already booked on " + c.date).join(", ");
      const allMsg     = [timeMsg, trainerMsg].filter(Boolean).join(" | ");
      bk.flags = [...(bk.flags || []), "⚠ Calendar conflict on " + bk.date + ": " + allMsg];
      results.conflicts.push({
        session_type: bk.session_type, date: bk.date,
        conflicts: slotCheck.conflicts, trainerConflicts: slotCheck.trainerConflicts
      });
      console.log("Calendar conflict for", bk.session_type, "on", bk.date, ":", allMsg);
      continue;
    }

    // Slot is available — create a TENTATIVE calendar event to hold it.
    // attendees: [] means no meeting-request emails are fired at this point.
    // Attendees are added (and invites sent) only when Diane hits Confirm.
    bk.trainer_emails_for_invite = trainerEmailsForApprove;

    const isFullDay = bk.full_day || false;

    // For all-day events (DEECA full_day=true) Graph API requires dateTime+timeZone at midnight.
    // The { date, timeZone } combo is invalid and causes a silent 400 — events never appear.
    const tentFullDayEnd = (() => {
      const [y, m, d] = startDT.slice(0, 10).split("-").map(Number);
      return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10) + "T00:00:00";
    })();
    const tentativeEvent = {
      subject: (bk.calendar_title || (entry.client + " — " + bk.session_type)) + " [Tentative]",
      body: {
        contentType: "HTML",
        content:
          "<p><strong>⏳ TENTATIVE — Pending Diane's Confirmation</strong></p>" +
          "<p><strong>Client:</strong> " + entry.client + "</p>" +
          "<p><strong>Session:</strong> " + bk.session_type + "</p>" +
          "<p><strong>Trainers:</strong> " + (bk.trainers || []).join(", ") + "</p>" +
          "<p><strong>Venue:</strong> " + (bk.venue || "TBC") + "</p>" +
          "<p><strong>Participants:</strong> " + (bk.participants || "TBC") + "</p>" +
          (bk.calendar_note ? "<p><strong>Notes:</strong> " + bk.calendar_note + "</p>" : "") +
          "<p><em>Slot held by Jasmine — no invites sent yet. Diane must approve before trainers are notified.</em></p>"
      },
      start: isFullDay
        ? { dateTime: startDT.slice(0, 10) + "T00:00:00", timeZone: "Australia/Brisbane" }
        : { dateTime: startDT, timeZone: "Australia/Brisbane" },
      end: isFullDay
        ? { dateTime: tentFullDayEnd, timeZone: "Australia/Brisbane" }
        : { dateTime: endDT, timeZone: "Australia/Brisbane" },
      isAllDay: isFullDay,
      showAs: "tentative",
      location: { displayName: bk.venue || "" },
      attendees: []  // ← no attendees = no invite emails until Diane confirms
    };

    try {
      const created = await graphRequest("POST",
        "/users/" + TRAINING_MAILBOX + "/calendar/events", tentativeEvent, token);
      if (created && created.id) {
        bk.calendar_event_id   = created.id;
        bk.calendar_pre_booked = true;
        bk.calendar_available  = true;
        results.calendar_events.push({ session_type: bk.session_type, date: bk.date, event_id: created.id });
        console.log("✓ Tentative slot held:", bk.session_type, "on", bk.date,
          "| no invites sent | ID:", created.id);
      }
    } catch (err) {
      console.error("Tentative event creation error:", err.message);
      // Non-fatal — booking still goes to review queue, Diane can approve manually
    }
  }

  return results;
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

CONFIRMATION REQUESTS: If the client is asking whether a session is already booked/confirmed (e.g. "can you confirm this is scheduled", "I have it noted as confirmed", "just checking this is still on"), use action_type=confirmation_check. Do NOT create a new booking. Set diane_summary to explain what the client is asking to confirm, and list the session details they mentioned in the bookings array with their existing dates/times. Do not generate trainer emails or client confirmation emails — set those fields to empty strings.

EMAIL SIGN-OFF: Kind Regards\nDiane Kruger\nCorporate Operations Lead\nRisk 2 Solution Group\n1300 459 970

OUTPUT JSON:
{
  "bookings": [{
    "action_type": "booking OR confirmation_check OR deeca_quote OR cancellation OR rescheduling OR conflict OR unknown",
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
${emailContent.slice(0, 6000)}`;
  }

  const step1Res = await client.messages.create({
    model: "claude-haiku-4-5-20251001",  // Haiku — fast, cheap, perfect for extraction
    max_tokens: 800,   // extraction output is short — list of facts, not prose
    messages: [{ role: "user", content: step1Content }]
  });

  const extractedFacts = step1Res.content[0].text || "";

  // Step 2: Apply Jasmine's rules and produce JSON
  const step2Res = await client.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 4000,  // 4000 covers most emails; truncation guard below catches large VA batches
    system: [{ type: "text", text: JASMINE_PROC_SYS, cache_control: { type: "ephemeral" } }],
    messages: [{ role: "user", content: extractedFacts }]
  });

  if (step2Res.stop_reason === "max_tokens") {
    // Retry once with 8000 tokens for large VA batch emails before giving up
    console.warn("Step 2 hit max_tokens (4000) — retrying with 8000 for large batch...");
    const retryRes = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 8000,
      system: [{ type: "text", text: JASMINE_PROC_SYS, cache_control: { type: "ephemeral" } }],
      messages: [{ role: "user", content: extractedFacts }]
    });
    if (retryRes.stop_reason === "max_tokens") {
      throw new Error(
        "Jasmine's JSON response was cut off even at 8000 tokens. " +
        "The email likely contains more sessions than usual. " +
        "Try splitting the email or processing sessions in smaller batches."
      );
    }
    const retryText = retryRes.content[0].text || "{}";
    let retryClean = retryText.replace(/```json|```/g, "").trim();
    if (!retryClean.startsWith("{")) throw new Error("Jasmine returned text instead of JSON: " + retryClean.slice(0, 120));
    const j = JSON.parse(retryClean);
    return { jasmineParsed: j, extractedFacts };
  }

  const text = step2Res.content[0].text || "{}";
  let clean = text.replace(/```json|```/g, "").trim();
  if (!clean.startsWith("{")) throw new Error("Jasmine returned text instead of JSON: " + clean.slice(0, 120));

  const j = JSON.parse(clean);
  return { jasmineParsed: j, extractedFacts };
}

// ── BOOKING RELEVANCE PRE-SCAN ────────────────────────────────────────────────
// Quick keyword check BEFORE calling Jasmine — costs zero API credits.
// Returns true if the email looks like it could be a booking-related request.
// Errs on the side of inclusion (false positives go to Jasmine; false negatives are silently skipped).
const BOOKING_KEYWORDS = [
  "booking", "book ", "session", "training", "course", "schedule", "scheduled",
  "confirm", "cancel", "reschedule", "dates", "availability", "participants",
  "attendees", "quote", "request for service", "ova", "security awareness",
  "de-escalation", "deescalation", "conflict", "warden", "emergency management",
  "initial", "recurrent", "fc ", "cc ", " fc\n", " cc\n", "cabin crew",
  "flight crew", "personal safety", "authorised officer", "rp7", "rp8", "rp9"
];

// If any of these appear in the body, it's an operational admin email (exam paperwork etc.)
// and should be skipped regardless of booking keywords in the subject.
const EXCLUSION_KEYWORDS = [
  // VA exam admin
  "exam sheet", "re-scan", "rescan", "scan back", "scan through",
  "re-scan back", "scanned back", "please scan", "scan image from va",
  // Feedback / surveys
  "feedback", "survey", "course feedback", "training feedback",
  // Results / records admin (not bookings)
  "class list", "attendance record", "attendance sheet", "summary report",
  "course summary", "send results", "please send results", "missing results",
  "has been actioned", "been actioned",
  // IT / account issues
  "password", "log in", "login", "unable to complete", "unable to access",
  // General admin noise
  "monthly reporting", "please send through", "still missing"
];

function looksLikeBookingEmail(subject, bodyText) {
  const haystack = ((subject || "") + " " + (bodyText || "")).toLowerCase();
  // Hard exclusion — operational threads always blocked regardless of attachments
  if (EXCLUSION_KEYWORDS.some(k => haystack.includes(k))) return false;
  return BOOKING_KEYWORDS.some(k => haystack.includes(k));
}


// Returns Monday 00:00 AEST of the current week as a UTC ISO string.
// AEST = UTC+10. Used to restrict inbox polling to this week only.
function getWeekStartUTC() {
  const AEST_OFFSET_MS = 10 * 60 * 60 * 1000;
  const nowAEST = new Date(Date.now() + AEST_OFFSET_MS);
  const dow = nowAEST.getUTCDay(); // 0=Sun, 1=Mon … 6=Sat
  const daysToMonday = dow === 0 ? 6 : dow - 1;
  const mondayAEST = new Date(nowAEST);
  mondayAEST.setUTCDate(mondayAEST.getUTCDate() - daysToMonday);
  mondayAEST.setUTCHours(0, 0, 0, 0);
  return new Date(mondayAEST.getTime() - AEST_OFFSET_MS).toISOString();
}

async function pollInbox() {
  if (!AZURE_CLIENT_ID) { console.log("Graph not configured — skipping poll."); return; }
  console.log("Jasmine polling inbox...");
  try {
    const token = await getGraphToken();

    // Fetch only this week's unprocessed emails from Graph.
    // Date filter stops old emails from April/March flooding in.
    // Category filter means $top=50 applies only to untagged emails, not total inbox.
    // ConsistencyLevel: eventual required for advanced $filter on categories.
    const weekStart = getWeekStartUTC();
    console.log("Polling for unprocessed emails since (UTC):", weekStart);
    const result = await fetch(
      "https://graph.microsoft.com/v1.0/users/" + TRAINING_MAILBOX +
      "/mailFolders/inbox/messages" +
      "?$filter=not categories/any(c:c eq 'Jasmine Processed')" +
      " and receivedDateTime ge " + weekStart +
      "&$orderby=receivedDateTime desc" +
      "&$top=50" +
      "&$count=true" +
      "&$select=id,subject,from,bodyPreview,body,receivedDateTime,hasAttachments,categories,webLink,conversationId,toRecipients,ccRecipients",
      {
        headers: {
          "Authorization": "Bearer " + token,
          "ConsistencyLevel": "eventual",   // required for advanced $filter on categories
          "Content-Type": "application/json"
        }
      }
    ).then(r => r.json());

    if (!result || !result.value) { console.log("No unprocessed emails found."); return; }
    console.log("Unprocessed emails from Graph: " + result.value.length);

    // shouldSkip: filter out calendar noise, auto-replies etc.
    const shouldSkip = (email) => {
      const subj = (email.subject || "").toLowerCase();
      const ct = (email.contentType || "").toLowerCase();
      if (ct.includes("calendar")) return true;
      if (subj.startsWith("accepted:")) return true;
      if (subj.startsWith("declined:")) return true;
      if (subj.startsWith("tentative:")) return true;
      if (subj.startsWith("cancelled:")) return false;
      if (subj.startsWith("automatic reply:")) return true;
      if (subj.startsWith("out of office:")) return true;
      if (subj.includes("delivery failed")) return true;
      if (subj.includes("undeliverable")) return true;
      return false;
    };

    const unprocessed = result.value.filter(email => !shouldSkip(email));
    console.log("After skip-filter: " + unprocessed.length + " to process");

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
            await processInboundEmail(email, clientName, token);
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

  // ── Quick relevance check — skip before spending any API credits ─────────
  // Internal forwards (FW:/Fwd: from staff) are always intentional — bypass pre-scan.
  const isInternalForward = (clientName === "INTERNAL" || (email.from && email.from.emailAddress &&
    email.from.emailAddress.address.toLowerCase().includes("@risk2solution.com"))) &&
    /^(fw|fwd|re:\s*fw|re:\s*fwd):/i.test(email.subject || "");

  if (!isInternalForward && !looksLikeBookingEmail(email.subject, bodyText)) {
    console.log("Pre-scan: no booking keywords found — skipping without API call: " + email.subject);
    emailLog.push({
      id: Date.now(),
      timestamp: new Date().toISOString(),
      subject: email.subject,
      from: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
      client: clientName,
      status: "skipped_not_booking",
      summary: "Pre-scan: no booking keywords — skipped without API call"
    });
    persistSave();
    return;
  }

  // ── Fetch attachments — Word docs, images (critical for VA inline tables) ────
  let attachmentText = "";
  let originalDocxBuffer = null;
  let originalDocxName = "";
  const inlineImages = []; // { contentType, base64 } for Claude vision

  // Fetch raw MIME email to extract inline images and attachments reliably
  // This is the same format as an .eml file — most reliable way to get inline images
  try {
    const token2 = token || await getGraphToken();
    const mimeUrl = "https://graph.microsoft.com/v1.0/users/" + TRAINING_MAILBOX + "/messages/" + email.id + "/$value";
    const mimeRes = await fetch(mimeUrl, {
      headers: { "Authorization": "Bearer " + token2, "Accept": "text/plain" }
    });

    if (mimeRes.ok) {
      const rawMime = await mimeRes.text();
      console.log("Raw MIME fetched: " + rawMime.length + " chars");

      // Parse MIME parts to find images and Word docs
      const boundaryMatch = rawMime.match(/boundary="?([^"\r\n;]+)"?/i);
      if (boundaryMatch) {
        const boundary = boundaryMatch[1];
        const parts = rawMime.split("--" + boundary);

        for (const part of parts) {
          if (!part || part.trim() === "--") continue;
          const headerEnd = part.indexOf("\r\n\r\n") !== -1 ? part.indexOf("\r\n\r\n") + 4 : part.indexOf("\n\n") + 2;
          const headers = part.slice(0, headerEnd).toLowerCase();
          const partBody = part.slice(headerEnd);

          const ctMatch = part.match(/Content-Type:\s*([^\r\n;]+)/i);
          const nameMatch = part.match(/(?:filename|name)="?([^"\r\n;]+)"?/i);
          const encMatch = part.match(/Content-Transfer-Encoding:\s*(\S+)/i);

          const ct = ctMatch ? ctMatch[1].trim().toLowerCase() : "";
          const fname = nameMatch ? nameMatch[1].trim().toLowerCase() : "";
          const enc = encMatch ? encMatch[1].trim().toLowerCase() : "";

          const isImage = ct.startsWith("image/");
          const isWord = fname.endsWith(".docx") || fname.endsWith(".doc");
          const isExcel = fname.endsWith(".xlsx") || fname.endsWith(".xls") || fname.endsWith(".csv");

          if (enc === "base64") {
            const b64 = partBody.replace(/[\r\n\s]/g, "");

            if (isImage && b64.length > 100) {
              const mediaType = ct.split(";")[0].trim() || "image/png";
              inlineImages.push({ mediaType, base64: b64 });
              console.log("Extracted inline image from MIME: " + (fname || ct) + " (" + b64.length + " chars b64)");
            } else if (isWord && b64.length > 100) {
              try {
                originalDocxBuffer = Buffer.from(b64, "base64");
                originalDocxName = fname || "attachment.docx";
                const result = await mammoth.extractRawText({ buffer: originalDocxBuffer });
                if (result.value) {
                  attachmentText += "\n\n=== ATTACHED DOCUMENT: " + fname + " ===\n" + result.value.trim();
                  console.log("Extracted Word doc from MIME: " + fname + " (" + result.value.length + " chars)");
                }
              } catch (err) {
                console.error("mammoth MIME error:", err.message);
              }
            } else if (isExcel) {
              attachmentText += "\n\n=== ATTACHED FILE: " + fname + " (Excel/CSV) ===";
            }
          }
        }
      } else {
        console.log("No MIME boundary found — simple email, no attachments to parse");
      }
    } else {
      console.warn("Could not fetch raw MIME:", mimeRes.status, mimeRes.statusText);
    }
  } catch (err) {
    console.error("MIME fetch error:", err.message);
  }

  console.log("MIME parse complete — images: " + inlineImages.length + ", docx: " + (originalDocxBuffer ? "yes" : "no") + ", attachmentText: " + attachmentText.length + " chars");

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
    "Body: " + bodyText.slice(0, 3000),
    attachmentText.slice(0, 8000)
  ].join("\n");

  let jasmineParsed, extractedFacts;
  try {
    const result = await processEmailWithJasmine(emailContent, inlineImages);
    jasmineParsed = result.jasmineParsed;
    extractedFacts = result.extractedFacts;

    // ── Calendar availability + tentative pre-booking ──────────────────────────
    // checkCalendarAvailability checks for slot/trainer conflicts and, if clear,
    // immediately creates a TENTATIVE calendar event to hold the slot.
    // Conflicted slots are skipped — no tentative event, just flags for Diane.
    if (token && jasmineParsed.bookings && jasmineParsed.bookings.length > 0) {
      const calResults = await checkCalendarAvailability(
        { client: clientName }, jasmineParsed, token
      );
      if (calResults.conflicts.length > 0) {
        console.log("Calendar conflicts found:", calResults.conflicts.length, "slots — fetching alternatives for all.");

        // ── Step 1: Gather alternatives for every conflicted date in parallel ─
        const conflictedDates = [];   // { date, session_type, trainerMsg, alts[] }
        for (const conflict of calResults.conflicts) {
          const conflictBk = jasmineParsed.bookings.find(
            bk => bk.date === conflict.date && bk.session_type === conflict.session_type
          );
          const trainerNames = (conflictBk && conflictBk.trainers) || [];
          const availDates   = await getTrainerAvailableDates(trainerNames, token);
          if (conflictBk) conflictBk.available_alternatives = availDates;
          const trainerMsg   = (conflict.trainerConflicts || []).map(tc => tc.trainer + " already booked").join(", ");
          conflictedDates.push({
            date: conflict.date,
            session_type: conflict.session_type,
            trainerMsg,
            alts: availDates
          });
        }

        // ── Step 2: Identify confirmed (available) dates ──────────────────────
        const confirmedDates = jasmineParsed.bookings
          .filter(bk => !calResults.conflicts.some(c => c.date === bk.date && c.session_type === bk.session_type))
          .map(bk => bk.date + " (" + bk.session_type + ")")
          .filter(Boolean);

        // ── Step 3: ONE Haiku call rewrites the client email with full context ─
        // This produces a single coherent email covering all confirmed dates AND
        // all unavailable dates with alternatives — not a series of partial rewrites.
        if (jasmineParsed.client_email_body) {
          try {
            const confirmedSection = confirmedDates.length > 0
              ? "CONFIRMED dates (proceed as planned): " + confirmedDates.join(", ")
              : "No dates were confirmed.";

            const conflictSection = conflictedDates.map(c => {
              const altStr = c.alts.length > 0
                ? "Available alternatives: " + c.alts.slice(0, 5).join(", ")
                : "Please contact us to arrange an alternative.";
              return "UNAVAILABLE: " + c.date + " (" + c.session_type + ")" +
                (c.trainerMsg ? " — " + c.trainerMsg : "") + ". " + altStr;
            }).join("\n");

            const rewriteRes = await client.messages.create({
              model: "claude-haiku-4-5-20251001",
              max_tokens: 800,
              messages: [{ role: "user", content:
                "Rewrite this client email as a single, coherent reply that covers ALL of the following in one email:\n\n" +
                confirmedSection + "\n\n" + conflictSection + "\n\n" +
                "Instructions:\n" +
                "- Open by confirming the dates that ARE proceeding as planned (if any).\n" +
                "- Then address each unavailable date: apologise, state it is unavailable, and list the specific alternative dates provided.\n" +
                "- Ask the client to reply confirming which alternative date suits for each unavailable slot.\n" +
                "- Keep it professional, warm and concise. Use the same sign-off as the original.\n" +
                "- Do NOT invent dates or details not provided above.\n" +
                "- Return ONLY the rewritten email body, no preamble.\n\n" +
                "Original email (for tone/sign-off reference only):\n" +
                jasmineParsed.client_email_body
              }]
            });
            jasmineParsed.client_email_body = rewriteRes.content[0].text || jasmineParsed.client_email_body;
            console.log("Client email rewritten — confirmed:", confirmedDates.length, "| conflicts:", conflictedDates.length);
          } catch (err) {
            console.error("Conflict email rewrite error:", err.message);
          }
        }

        // ── Step 4: Update Diane's summary with full picture ──────────────────
        const conflictNote = "\n\n⚠ MIXED RESULT — " + confirmedDates.length + " date(s) confirmed, " +
          conflictedDates.length + " unavailable:\n" +
          (confirmedDates.length ? "✓ Confirmed: " + confirmedDates.join(", ") + "\n" : "") +
          conflictedDates.map(c =>
            "✗ Unavailable: " + c.session_type + " on " + c.date +
            (c.trainerMsg ? " (" + c.trainerMsg + ")" : "") +
            (c.alts.length ? "\n  Alternatives in draft: " + c.alts.slice(0, 3).join(" | ") : "")
          ).join("\n") +
          "\n\nClient draft updated — covers confirmed and unavailable dates in one email. Please review before sending.";
        jasmineParsed.diane_summary = (jasmineParsed.diane_summary || "") + conflictNote;
        jasmineParsed.has_conflicts = true;
      }
    }
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

  // ── Auto-dismiss if Jasmine found nothing actionable ─────────────────────
  // If action_type is unknown AND there are no bookings, this email isn't a
  // booking request. Log it and skip adding to the review queue entirely.
  const overallActionType = bookings.length > 0 ? bookings[0].action_type : (jasmineParsed.action_type || "unknown");
  if (overallActionType === "unknown" && bookings.length === 0) {
    console.log("Jasmine: no actionable booking found — auto-dismissing: " + email.subject);
    emailLog.push({
      id: Date.now(),
      timestamp: new Date().toISOString(),
      subject: email.subject,
      from: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
      client: clientName,
      status: "auto_dismissed",
      summary: "Jasmine found no booking content — auto-dismissed"
    });
    persistSave();
    return;
  }

  const bookingTagMap = {
    "booking": { label: "Booking Request", color: "#1d4ed8", bg: "#dbeafe" },
    "confirmation_check": { label: "Confirmation Required", color: "#7c3aed", bg: "#f3e8ff" },
    "deeca_quote": { label: "DEECA Quote", color: "#15803d", bg: "#dcfce7" },
    "cancellation": { label: "Cancellation", color: "#b91c1c", bg: "#fef2f2" },
    "rescheduling": { label: "Rescheduling", color: "#d97706", bg: "#fef9c3" },
    "conflict": { label: "Conflict", color: "#7c3aed", bg: "#f3e8ff" },
    "unknown": { label: "Needs Review", color: "#64748b", bg: "#f1f5f9" },
    "error": { label: "Error", color: "#b91c1c", bg: "#fef2f2" }
  };

  // ── Rescheduling deduplication ──────────────────────────────────────────────
  // When a client replies with a new date, Jasmine returns action_type=rescheduling.
  // Find the original pending/approved booking for the same client+session_type
  // and cancel its tentative calendar event so no duplicate appears in Outlook.
  if (overallActionType === "rescheduling") {
    try {
      const t = token || await getGraphToken();
      for (const bk of bookings) {
        // Find the most recent queue entry for the same client + session type
        const priorEntry = [...reviewQueue]
          .reverse()
          .find(e =>
            e.client === clientName &&
            e.id !== undefined &&
            (e.status === "pending" || e.status === "approved") &&
            (e.bookings || []).some(pb => pb.session_type === bk.session_type)
          );
        if (priorEntry) {
          for (const priorBk of (priorEntry.bookings || [])) {
            if (priorBk.session_type === bk.session_type && priorBk.calendar_event_id && priorBk.calendar_pre_booked) {
              // Delete the old tentative calendar event — new date will be pre-booked fresh
              try {
                await graphRequest("DELETE",
                  "/users/" + TRAINING_MAILBOX + "/calendar/events/" + priorBk.calendar_event_id,
                  null, t);
                console.log("Rescheduling: deleted old tentative event for", bk.session_type, "on", priorBk.date);
                priorBk.calendar_event_id = null;
                priorBk.calendar_pre_booked = false;
              } catch (delErr) {
                console.error("Rescheduling: calendar delete error:", delErr.message);
              }
            }
          }
          // Mark the prior entry as superseded so it doesn't clutter the queue
          if (priorEntry.status === "pending") {
            priorEntry.status = "superseded_by_reschedule";
            console.log("Rescheduling: prior queue entry", priorEntry.id, "marked superseded.");
          }
        }
      }
    } catch (err) {
      console.error("Rescheduling dedup error:", err.message);
    }
  }

  const bookingTag = bookingTagMap[overallActionType] || bookingTagMap["unknown"];

  const entry = {
    id: Date.now(),
    timestamp: new Date().toISOString(),
    email_id: email.id,
    email_web_link: email.webLink || "",          // ← Outlook web URL for "View Original" button
    received_datetime: email.receivedDateTime || "",
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

  // ── NO DRAFTS CREATED HERE ──────────────────────────────────────────────────
  // All email content (client_email_body, trainer_emails, etc.) is stored on
  // the queue entry above. Drafts are only created when Diane hits Confirm
  // in the review queue — see approveQueueEntry(). This ensures nothing is
  // staged in Outlook until Diane has reviewed and approved the booking.

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
  persistSave(); // Write queue + bookings to disk
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

    // Resolve booking type flags up-front — used in both the pre-booked upgrade
    // path AND the new-event creation path below (previously isDeeca was declared
    // only after the upgrade block, causing a temporal dead zone ReferenceError).
    const isDeeca = bk.action_type === "deeca_quote";

    // If this slot was pre-booked as tentative during processing, upgrade to confirmed
    // PATCH adds attendees → Outlook fires the meeting-request invite to trainers.
    if (bk.calendar_event_id && bk.calendar_pre_booked) {
      try {
        const inviteAttendees = [];
        // Prefer the saved addresses from processing; fall back to trainer_emails
        const savedAddrs = bk.trainer_emails_for_invite || [];
        if (savedAddrs.length > 0) {
          savedAddrs.forEach(addr =>
            inviteAttendees.push({ emailAddress: { address: safeEmail(addr) }, type: "required" }));
        } else {
          for (const te of (entry.trainer_emails || [])) {
            for (const addr of (te.trainer_email_to || [])) {
              if (addr && addr.includes("@"))
                inviteAttendees.push({ emailAddress: { address: safeEmail(addr) }, type: "required" });
            }
          }
        }
        await graphRequest("PATCH",
          "/users/" + TRAINING_MAILBOX + "/calendar/events/" + bk.calendar_event_id,
          {
            showAs: isDeeca ? "tentative" : "busy",
            subject: (bk.calendar_title || (entry.client + " — " + bk.session_type))
              .replace(" [Tentative]", ""),
            attendees: inviteAttendees  // adding attendees now triggers invite emails
          }, token);
        results.calendar.push(bk.session_type + " confirmed — trainer invites sent");
        entry.calendar_event_id = bk.calendar_event_id;
        console.log("✓ Tentative → confirmed:", bk.session_type,
          "| invites sent to:", inviteAttendees.map(a => a.emailAddress.address).join(", ") || "none (TEST_MODE or no trainers)");
      } catch (err) {
        console.error("Calendar confirm/upgrade error:", err.message);
        results.errors.push("Calendar confirm error for " + bk.session_type + ": " + err.message);
      }
      continue; // skip new-event creation below
    }

    // Build calendar event datetime (for bookings not pre-booked)
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

    const isFullDay = bk.full_day || false;
    // For all-day events (e.g. DEECA) the Graph API requires dateTime+timeZone at midnight.
    // Using { date: "...", timeZone: "..." } is invalid and causes a silent 400 error.
    const fullDayEndStr = (() => {
      const [y, m, d] = startDT.slice(0, 10).split("-").map(Number);
      return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10) + "T00:00:00";
    })();

    // Attendees — trainer emails routed through safeEmail() for TEST_MODE support
    const trainerAttendees = [];
    for (const te of (entry.trainer_emails || [])) {
      for (const addr of (te.trainer_email_to || [])) {
        if (addr && addr.includes("@")) {
          trainerAttendees.push({ emailAddress: { address: safeEmail(addr) }, type: "required" });
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
        ? { dateTime: startDT.slice(0, 10) + "T00:00:00", timeZone: "Australia/Brisbane" }
        : { dateTime: startDT, timeZone: "Australia/Brisbane" },
      end: isFullDay
        ? { dateTime: fullDayEndStr, timeZone: "Australia/Brisbane" }
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
            subject: (TEST_MODE ? "[TEST] " : "") + (entry.client_email_subject || "Training Booking Confirmation"),
            toRecipients: [{ emailAddress: { address: safeEmail(entry.client_email_to) } }],
            body: { contentType: "HTML", content: (TEST_MODE ? "<div style=\"background:#fef9c3;border:1px solid #ca8a04;padding:10px 14px;border-radius:6px;margin-bottom:12px;font-size:13px;\"><strong>⚠ TEST MODE</strong> — Real recipient: <em>" + entry.client_email_to + "</em> · Redirected to " + TEST_EMAIL + "</div>" : "") + "<p>" + (entry.client_email_body||"").split("\n").join("<br>") + "</p>" },
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
            subject: (TEST_MODE ? "[TEST] " : "") + (te.trainer_email_subject || "Session Booking"),
            toRecipients: te.trainer_email_to.map(a => ({ emailAddress: { address: safeEmail(a) } })),
            body: { contentType: "HTML", content: (TEST_MODE ? "<div style=\"background:#fef9c3;border:1px solid #ca8a04;padding:10px 14px;border-radius:6px;margin-bottom:12px;font-size:13px;\"><strong>⚠ TEST MODE</strong> — Real recipient(s): <em>" + te.trainer_email_to.join(", ") + "</em> · Redirected to " + TEST_EMAIL + "</div>" : "") + "<p>" + (te.trainer_email_body||"").split("\n").join("<br>") + "</p>" },
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
  persistSave(); // Persist confirmed status

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
  const pending = reviewQueue.filter(e => e.status === "pending");
  pending.sort((a, b) => {
    const da = new Date(a.received_datetime || a.timestamp);
    const db = new Date(b.received_datetime || b.timestamp);
    return db - da; // most recent first
  });
  res.json(pending);
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
  // Delete any tentative calendar events that were pre-booked during processing
  // so the slot is freed up and no phantom events remain in the calendar.
  (async () => {
    try {
      const t = await getGraphToken();
      for (const bk of (entry.bookings || [])) {
        if (bk.calendar_event_id && bk.calendar_pre_booked) {
          await graphRequest("DELETE",
            "/users/" + TRAINING_MAILBOX + "/calendar/events/" + bk.calendar_event_id,
            null, t);
          console.log("Deleted tentative event on reject:", bk.session_type, bk.calendar_event_id);
        }
      }
    } catch (err) {
      console.error("Reject — calendar cleanup error:", err.message);
    }
  })();
  // Delete Outlook drafts that were staged during processing
  (async () => {
    try {
      const t = await getGraphToken();
      for (const draft of (entry.drafts || [])) {
        if (draft.draft_id) {
          await graphRequest("DELETE",
            "/users/" + TRAINING_MAILBOX + "/messages/" + draft.draft_id,
            null, t);
          console.log("Deleted draft on reject:", draft.description);
        }
      }
    } catch (err) {
      console.error("Reject — draft cleanup error:", err.message);
    }
  })();
  persistSave();
  res.json({ success: true });
});

app.post("/api/review-queue/:id/dismiss", requireAuth, (req, res) => {
  const idx = reviewQueue.findIndex(e => e.id === parseInt(req.params.id));
  if (idx === -1) return res.status(404).json({ error: "Not found" });
  const dismissedEntry = reviewQueue[idx];
  dismissedEntry.status = "dismissed";
  dismissedEntry.dismissed_at = new Date().toISOString();
  // Clean up calendar events and drafts staged for this entry
  (async () => {
    try {
      const t = await getGraphToken();
      for (const bk of (dismissedEntry.bookings || [])) {
        if (bk.calendar_event_id && bk.calendar_pre_booked) {
          await graphRequest("DELETE",
            "/users/" + TRAINING_MAILBOX + "/calendar/events/" + bk.calendar_event_id,
            null, t);
          console.log("Deleted tentative event on dismiss:", bk.session_type);
        }
      }
      for (const draft of (dismissedEntry.drafts || [])) {
        if (draft.draft_id) {
          await graphRequest("DELETE",
            "/users/" + TRAINING_MAILBOX + "/messages/" + draft.draft_id,
            null, t);
          console.log("Deleted draft on dismiss:", draft.description);
        }
      }
    } catch (err) {
      console.error("Dismiss — cleanup error:", err.message);
    }
  })();
  persistSave();
  res.json({ success: true });
});

// ── POLL NOW ──────────────────────────────────────────────────────────────────
app.post("/api/poll-now", requireAuth, async (req, res) => {
  try { await pollInbox(); res.json({ success: true }); }
  catch (err) { res.status(500).json({ error: err.message }); }
});

// Force reprocess — scans recent 50 emails but only calls Jasmine on up to 20
// that pass the booking keyword pre-scan and aren't already in the queue.
app.post("/api/reprocess-all", requireAuth, async (req, res) => {
  try {
    const token = await getGraphToken();
    const result = await graphRequest("GET",
      "/users/" + TRAINING_MAILBOX + "/mailFolders/inbox/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,bodyPreview,body,receivedDateTime,hasAttachments,categories,toRecipients,ccRecipients",
      null, token);
    if (!result || !result.value) return res.json({ success: true, processed: 0, skipped: 0 });
    let count = 0;
    let skipped = 0;
    const API_CALL_LIMIT = 20; // never burn more than 20 Jasmine calls per Force Reprocess
    for (const email of result.value) {
      if (count >= API_CALL_LIMIT) {
        console.log("Reprocess: API call limit (" + API_CALL_LIMIT + ") reached — stopping.");
        break;
      }
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
        // Skip if already in queue
        const alreadyQueued = reviewQueue.some(e => e.email_id === email.id);
        if (alreadyQueued) {
          console.log("Skipping reprocess — already in queue: " + email.subject);
          skipped++;
          continue;
        }
        // Pre-scan: skip non-booking emails without using any API credits
        if (!email.hasAttachments && !looksLikeBookingEmail(email.subject, email.bodyPreview || "")) {
          console.log("Reprocess pre-scan: skipping non-booking email: " + email.subject);
          skipped++;
          continue;
        }
        await processInboundEmail(email, clientName, token);
        count++;
      }
    }
    res.json({ success: true, processed: count, skipped });
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
      model: "claude-haiku-4-5-20251001",  // Haiku — chat is conversational Q&A, no need for Sonnet
      max_tokens: 800,
      system: [{ type: "text", text: JASMINE_CHAT_SYS, cache_control: { type: "ephemeral" } }],
      messages: messages.slice(-10)   // sliding window — last 10 messages only
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
  persistLoad(); // Restore reviewQueue, bookingsLog, emailLog from disk
  if (AZURE_CLIENT_ID) {
    pollInbox();
    setInterval(pollInbox, 5 * 60 * 1000); // Poll every 5 minutes — keeps system prompt cache warm (5-min TTL)
    console.log("Email polling started — every 5 minutes.");
  } else {
    console.log("Azure not configured — email polling disabled. Set AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET to enable.");
  }
});
