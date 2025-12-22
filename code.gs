/**
 * Church Tech Scheduler — One Single File (Option B + Confirm UX + Auto-create Event + Tech Sub RSVP Buttons)
 *
 * ✅ Schedule is source of truth (emails)
 * ✅ Monthly sheets show NAMES (dropdown), store EMAILS in hidden columns
 * ✅ Monthly edit updates Schedule + stores a pending change (NO email yet)
 * ✅ User confirms in sidebar -> calendar updated + emails sent
 * ✅ Auto-create calendar event if missing Event Id
 * ✅ Saturday 5PM reminder emails
 * ✅ Batch scheduler creates missing Sundays/events
 *
 * ✅ NEW: Tech Sub Emails + YES/NO buttons
 * - When a role changes (oldEmail -> newEmail), sends:
 *   (1) To Person A (old): "Tech Sub Confirmation"
 *   (2) To Person B (new): "Tech Sub Invite" with YES/NO buttons
 * - YES/NO buttons call Apps Script Web App (doGet) to update:
 *   - Schedule RSVP column for that role/date
 *   - Monthly RSVP column for that date
 *   - Calendar event description (adds Sub RSVP line)
 *   - (optional) guest add/remove
 *
 * REQUIRED SCRIPT PROPERTIES:
 * - SPREADSHEET_ID = <your Google Sheet ID>
 *
 * REQUIRED CONFIG:
 * - CONFIG.WEBAPP_URL = deployed Web App URL (Deploy > New deployment > Web app)
 * - CONFIG.WEBAPP_SECRET = long random string
 */

// =====================
// CONFIG
// =====================
const TEST_EMAIL_ONLY = "tubigangelito981@gmail.com"; // set null when LIVE
const TEST_SMS_ONLY = "639157213563"; // set to "639XXXXXXXXX" to force SMS to one number when testing

const CONFIG = {
  TIMEZONE: "Asia/Manila",
  CALENDAR_NAME: "Church Tech Schedule",

  ROSTER_SHEET_NAME: "Roster",
  SCHEDULE_SHEET_NAME: "Schedule",

  START_DATE: "2025-12-19",
  END_DATE: "2026-12-31",

  MINISTRIES: ["LiveStream", "Audio", "PPT"],

  MAX_ASSIGNMENTS_PER_PERSON_PER_MONTH_PER_MINISTRY: 2,
  AVOID_CONSECUTIVE_SUNDAYS_SAME_MINISTRY: true,

  // batch + speed
  BATCH_SUNDAYS: 10,
  AUTO_BATCH_EVERY_MINUTES: 1,
  RSVP_SYNC_EVERY_MINUTES: 30,

  // branding
  CHURCH_NAME: "Scripture Alone Baptist Church",
  TECH_TEAM_NAME: "Tech Team",

  LOGO_URL:
    "https://scontent.fcrk1-4.fna.fbcdn.net/v/t39.30808-6/581931274_1143165951345548_7970612996274936410_n.jpg?_nc_cat=109&ccb=1-7&_nc_sid=6ee11a&_nc_eui2=AeE-BXRETeSyNlfd6wkbF8AebLI5KjsVKv5ssjkqOxUq_p1zeegMthA0fPsa1h9ASXPjkOKEsmRG66ukeyrTr17C&_nc_ohc=Yypu83H4jmkQ7kNvwEX9Mf_&_nc_oc=AdlfRmpC8r7svvlS32DVqmfIP9t1kLBGNr5COZcXy4QqY9NYDPDfduve81j1PfuVjKA&_nc_zt=23&_nc_ht=scontent.fcrk1-4.fna&_nc_gid=o1INtoXGXhSnC9pvLMQh6A&oh=00_Aflhn3UZGn6KXluKDdnP8CwpwegDi4KSYE6X6VEbgdw-9g&oe=69498F4A",
  SHOW_LOGO: true,

  // SMS (PhilSMS)
  // Store your token in Script Properties: PHILSMS_TOKEN
  SMS_ENABLED: true,
  SMS_API_BASE: "https://dashboard.philsms.com/api/v3",
  SMS_SENDER_ID: "PhilSMS",
  SMS_TYPE: "plain",
  SMS_TOKEN: "", // optional fallback if Script Properties are not set
  COORDINATOR_MOBILE: "639760446193", // optional: SMS number for Tech Sub NO alerts

  // monthly sheet date format
  MONTH_DATE_FORMAT: "ddd, mmm d, yyyy",

  // ===== NEW: Web App for YES/NO buttons =====
  // Deploy > New deployment > Web app
  // Execute as: Me
  // Who has access: Anyone with the link (works for personal emails)
  WEBAPP_URL: "https://script.google.com/macros/s/AKfycbw1FsNt5az_3-Dh_YAYIBmTGtPUJGVxoTfBEn_cnQA6irVCkdFn8lXlg13Lkj2pC4p2/exec", // <-- paste your deployed web app URL here
  WEBAPP_SECRET: "AKfycbw1FsNt5az_3-Dh_YAYIBmTGtPUJGVxoTfBEn_cnQA6irVCkdFn8lXlg13Lkj2pC4p2",
};

// =====================
// MONTHLY SHEETS 
// =====================
const MONTH_SHEETS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

// Monthly Sheet Layout 
// A: Date (Date object, formatted)
// B: Audio (NAME dropdown)
// C: Livestream (NAME dropdown)
// D: PPT (NAME dropdown)
// E: Audio Email (hidden)
// F: Livestream Email (hidden)
// G: PPT Email (hidden)
// H: Audio RSVP
// I: Livestream RSVP
// J: PPT RSVP
const MONTH_HEADERS = [
  "Date", "Audio", "Livestream", "PPT",
  "Audio Email", "Livestream Email", "PPT Email",
  "Audio RSVP", "Livestream RSVP", "PPT RSVP",
];

// =====================
// PENDING CONFIRM STATE
// =====================
const PENDING_KEY = "TECH_SCHED_PENDING_CHANGE";
const SUB_PENDING_KEY = "TECH_SUB_PENDING_CONFIRMATIONS";

// =====================
// MENU
// =====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Tech Scheduler");

  menu
    .addItem("Start Auto Batch Scheduler", "startAutoBatchScheduler")
    .addSeparator()
    .addItem("Setup Auto-Sync on Schedule Edit", "setupAutoSyncOnEdit")
    .addItem("Setup Monthly Edit Sync ", "setupMonthlyEditSyncOptionB")
    .addSeparator()
    .addItem("Generate Monthly Sheets ", "generateMonthlySheetsOptionB")
    .addItem("Sync Monthly Sheets from Schedule ", "syncMonthlySheetsFromScheduleOptionB")
    .addSeparator()
    .addItem("Setup Monthly Reminder (Last Day)", "setupMonthlyScheduleReminderTrigger")
    .addItem("Send Monthly Reminder Now", "sendMonthlyScheduleReminder_")
    .addSeparator()
    .addItem("Check Email Quota", "logEmailQuota")
    .addSeparator()
    .addItem("Confirm & Send Emails", "openConfirmSendSidebar")
    .addItem("Cancel Pending Change", "cancelPendingChange")
    .addSeparator()
    .addItem("Setup RSVP Sync (ICS)", "setupRsvpSyncTrigger")
    .addItem("Sync RSVPs Now", "syncRsvpStatuses")    
    .addSeparator()
    .addItem("Send Reminder Now", "sendSaturdayReminder")
    .addItem("Setup Wednesday 7PM Trigger", "setupSaturday5pmTrigger")
    .addSeparator()
    .addItem("Show WebApp URL Help", "showWebAppHelp_")
    .addToUi();
}

// Optional installer (one-time) to create both triggers quickly
function install() {
  setupAutoSyncOnEdit();
  setupMonthlyEditSyncOptionB();
  setupRsvpSyncTrigger();
  SpreadsheetApp.getUi().alert(
    "Installed triggers:\n" +
    "- onScheduleEdit\n" +
    "- onMonthlyEditOptionB\n" +
    "- syncRsvpStatuses\n\n" +
    "Monthly changes now require Confirm & Send Emails."
  );
}

function showWebAppHelp_() {
  SpreadsheetApp.getUi().alert(
    "Quick Start (from scratch):\n" +
    "1) Tech Scheduler > Start Auto Batch Scheduler\n" +
    "2) Wait until the Schedule tab is filled up to END_DATE\n" +
    "3) Tech Scheduler > Generate Monthly Sheets (Option B)\n" +
    "4) Tech Scheduler > Sync Monthly Sheets from Schedule (Option B)\n\n" +
    "Install triggers (required):\n" +
    "5) Tech Scheduler > install (sets: onScheduleEdit, onMonthlyEditOptionB, syncRsvpStatuses)\n\n" +
    "Enable YES/NO buttons:\n" +
    "6) Deploy > New deployment > Web app\n" +
    "7) Execute as: Me\n" +
    "8) Who has access: Anyone with the link\n" +
    "9) Copy the Web App URL and paste into CONFIG.WEBAPP_URL\n\n" +
    "Security is enforced by signed links + roster + assignment checks."
  );
}

// =====================
// WEB APP ENDPOINT (YES/NO buttons)
// =====================

/**
 * Secure RSVP endpoint.
 * Only works if:
 * 1) signature is valid (HMAC)
 * 2) email exists in Roster
 * 3) email matches current Schedule assignee for role+date (prevents forwarding)
 *
 * URL format:
 *   <WEBAPP_URL>?a=YES&role=LiveStream&date=2026-01-04&email=x@y.com&eid=...&sig=...
 */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.a || "").toUpperCase(); // YES / NO
    const role = String(p.role || "");
    const dateKey = String(p.date || "");
    const email = String(p.email || "").trim().toLowerCase();
    const eventId = String(p.eid || "").trim();
    const sig = String(p.sig || "");

    if (!["YES", "NO"].includes(action)) return html_(uiMsg_("Invalid action."));
    if (!role || !dateKey || !email || !eventId || !sig) return html_(uiMsg_("Missing parameters."));

    // 1) Signature check (prevents guessing/tampering)
    const expected = signRsvp_({ action, role, dateKey, email, eventId });
    if (sig !== expected) return html_(uiMsg_("Invalid link. Please contact the Tech Team."));

    const ss = SpreadsheetApp.openById(getSpreadsheetId_());

    // 2) Roster whitelist check
    const roster = buildRosterMaps_(ss);
    if (!roster.emailToName.has(email)) {
      return html_(uiMsg_("You are not listed in the Tech Team roster."));
    }

    // 3) Assignment binding check (prevents forwarding)
    const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!schedule) return html_(uiMsg_("Schedule sheet missing."));
    ensureScheduleHeaders_(schedule);

    const row = findOrCreateScheduleRowByDate_(schedule, dateKey);
    const header = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
    const H = headerIndex_(header);

    if (H[role] == null) return html_(uiMsg_("Unknown role column."));
    const expectedAssignee = String(schedule.getRange(row, H[role] + 1).getValue() || "").trim().toLowerCase();

    if (expectedAssignee !== email) {
      return html_(uiMsg_("This link is not for your assignment."));
    }

    applySubRsvpUpdate_({ action, role, dateKey, email, eventId });
    return html_(uiSuccess_(action));
  } catch (err) {
    console.error(err);
    return html_(uiMsg_("Something went wrong. Please message the Tech Team."));
  }
}

function applySubRsvpUpdate_({ action, role, dateKey, email, eventId }) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedule) throw new Error("Schedule sheet missing.");
  ensureScheduleHeaders_(schedule);

  // Update schedule RSVP
  const row = findOrCreateScheduleRowByDate_(schedule, dateKey);
  const header = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);

  const rsvpColName =
    role === "Audio" ? "Audio RSVP" :
    role === "LiveStream" ? "LiveStream RSVP" :
    role === "PPT" ? "PPT RSVP" :
    "";

  if (!rsvpColName || H[rsvpColName] == null) throw new Error("RSVP column missing for role: " + role);

  const newVal = action === "YES" ? "Yes" : "No";
  schedule.getRange(row, H[rsvpColName] + 1).setValue(newVal);

  // Update monthly RSVP
  updateMonthlyRsvpForDate_(ss, dateKey, {
    Audio: role === "Audio" ? newVal : undefined,
    LiveStream: role === "LiveStream" ? newVal : undefined,
    PPT: role === "PPT" ? newVal : undefined,
  });

  // Update calendar event description (and optionally guest list)
  const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);
  const event = safeGetEventById_(cal, eventId);
  if (!event) throw new Error("Calendar event not found.");

  const name = getNameFromEmail_(ss, email) || email;
  const marker = `SUB_RSVP|${role}|${email}`;
  const lines = String(event.getDescription() || "").split("\n");
  const filtered = lines.filter(l => !String(l).includes(marker));

  const statusText = action === "YES" ? "YES ✅" : "NO ❌";
  filtered.push(`${marker} :: ${name} responded ${statusText}`);

  event.setDescription(filtered.join("\n"));

  if (!TEST_EMAIL_ONLY) {
    if (action === "YES") {
      ensureGuestOnEvent_(event, email);
    } else {
      try { event.removeGuest(email); } catch (_) {}
    }
  }

  if (action === "YES") {
    sendPendingSubConfirmationIfAny_(ss, dateKey, role, email);
  }

  // Optional: notify coordinator when NO
  if (action === "NO") {
    const notifyEmail = TEST_EMAIL_ONLY || "johncyrilflaviano@gmail.com";
    GmailApp.sendEmail(
      notifyEmail,
      `[Tech Sub] NO - ${role} - ${dateKey}`,
      `${email} declined to cover ${role} on ${dateKey}. Please coordinate in the Tech Group chat.`
    );
    if (CONFIG.COORDINATOR_MOBILE) {
      const sms = buildSmsTechSubNoNotify_({ role, dateKey, email });
      sendSms_({ recipient: CONFIG.COORDINATOR_MOBILE, message: sms });
    }
  }
}

function getPendingSubConfirmations_() {
  const s = PropertiesService.getDocumentProperties().getProperty(SUB_PENDING_KEY);
  const list = s ? JSON.parse(s) : [];
  return Array.isArray(list) ? list : [];
}

function savePendingSubConfirmation_(item) {
  const list = getPendingSubConfirmations_();
  const key = `${item.dateKey}|${item.role}|${String(item.newEmail || "").toLowerCase()}`;
  const filtered = list.filter(x =>
    `${x.dateKey}|${x.role}|${String(x.newEmail || "").toLowerCase()}` !== key
  );
  filtered.push(item);
  PropertiesService.getDocumentProperties().setProperty(SUB_PENDING_KEY, JSON.stringify(filtered));
}

function removePendingSubConfirmation_(dateKey, role, newEmail) {
  const list = getPendingSubConfirmations_();
  const key = `${dateKey}|${role}|${String(newEmail || "").toLowerCase()}`;
  const filtered = list.filter(x =>
    `${x.dateKey}|${x.role}|${String(x.newEmail || "").toLowerCase()}` !== key
  );
  PropertiesService.getDocumentProperties().setProperty(SUB_PENDING_KEY, JSON.stringify(filtered));
}

function sendPendingSubConfirmationIfAny_(ss, dateKey, role, newEmail) {
  const list = getPendingSubConfirmations_();
  const key = `${dateKey}|${role}|${String(newEmail || "").toLowerCase()}`;
  const match = list.find(x =>
    `${x.dateKey}|${x.role}|${String(x.newEmail || "").toLowerCase()}` === key
  );
  if (!match || !match.oldEmail) return;

  const roster = buildRosterMaps_(ss);
  const personAEmail = String(match.oldEmail || "").trim().toLowerCase();
  const personBEmail = String(match.newEmail || "").trim().toLowerCase();
  const personAName = roster.emailToName.get(personAEmail) || personAEmail;
  const personBName = roster.emailToName.get(personBEmail) || personBEmail;
  const prettyDate = Utilities.formatDate(new Date(dateKey + "T00:00:00"), CONFIG.TIMEZONE, "EEE, MMM d, yyyy");

  const confHtml = buildTechSubConfirmationEmail_({
    personAName,
    personBName,
    role,
    prettyDate,
  });
  GmailApp.sendEmail(TEST_EMAIL_ONLY || personAEmail, "Tech Sub Confirmation", "Please view in HTML.", { htmlBody: confHtml });
  const personAMobile = roster.emailToMobile.get(personAEmail) || "";
  if (personAMobile) {
    const sms = buildSmsTechSubConfirmation_({ personBName, role, prettyDate });
    sendSms_({ recipient: personAMobile, message: sms });
  }

  removePendingSubConfirmation_(dateKey, role, newEmail);
}

function buildWebAppRsvpLink_({ action, role, dateKey, assigneeEmail, eventId }) {
  if (!CONFIG.WEBAPP_URL) return "#";
  const email = String(assigneeEmail || "").trim().toLowerCase();
  const sig = signRsvp_({ action, role, dateKey, email, eventId });
  const q = [
    `a=${encodeURIComponent(action)}`,
    `role=${encodeURIComponent(role)}`,
    `date=${encodeURIComponent(dateKey)}`,
    `email=${encodeURIComponent(email)}`,
    `eid=${encodeURIComponent(eventId)}`,
    `sig=${encodeURIComponent(sig)}`,
  ].join("&");
  return `${CONFIG.WEBAPP_URL}?${q}`;
}

function signRsvp_({ action, role, dateKey, email, eventId }) {
  const base = [action, role, dateKey, String(email || "").trim().toLowerCase(), eventId].join("|");
  const raw = Utilities.computeHmacSha256Signature(base, CONFIG.WEBAPP_SECRET);
  return Utilities.base64EncodeWebSafe(raw);
}

function html_(body) {
  return HtmlService.createHtmlOutput(body)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function uiSuccess_(action) {
  return `
    <div style="font-family:Arial,Helvetica,sans-serif;padding:18px;">
      <div style="font-size:18px;font-weight:800;margin-bottom:6px;">Response recorded</div>
      <div style="font-size:14px;color:#334155;">
        You answered <b>${escapeHtml_(action)}</b>. Thank you!
      </div>
      <div style="margin-top:14px;font-size:12px;color:#64748b;">
        You may now close this tab.
      </div>
    </div>
  `;
}
function uiMsg_(msg) {
  return `
    <div style="font-family:Arial,Helvetica,sans-serif;padding:18px;">
      <div style="font-size:14px;color:#334155;">${escapeHtml_(msg)}</div>
    </div>
  `;
}

// =====================
// BATCH SCHEDULER (admin)
// =====================
function startAutoBatchScheduler() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runSchedulerBatch")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("runSchedulerBatch")
    .timeBased()
    .everyMinutes(CONFIG.AUTO_BATCH_EVERY_MINUTES)
    .create();

  SpreadsheetApp.getActive().toast("Auto batch scheduler started.", "Tech Scheduler", 5);
}

function stopAutoBatchScheduler() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runSchedulerBatch")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function resetBatchCursor() {
  PropertiesService.getScriptProperties().deleteProperty("SCHED_CURSOR_SUNDAY");
}

// =====================
// MAIN BATCH CREATION
// =====================
function runSchedulerBatch() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId_());
    const roster = readRosterFromSheet_(ss, CONFIG.ROSTER_SHEET_NAME);

    const sh = getOrCreateSheet_(ss, CONFIG.SCHEDULE_SHEET_NAME, getScheduleHeaders_());
    ensureScheduleHeaders_(sh);

    const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);
    const sundays = getSundays_(CONFIG.START_DATE, CONFIG.END_DATE);

    const props = PropertiesService.getScriptProperties();
    let cursor = Number(props.getProperty("SCHED_CURSOR_SUNDAY") || "0");

    const existing = loadExistingScheduleByDate_(sh);
    const countsByMonth = rebuildCountsByMonthAllRolesFromSheet_(sh);
    const emailToDates = rebuildEmailToDatesFromSheet_(sh);

    const rows = [];
    let created = 0;

    while (cursor < sundays.length && created < CONFIG.BATCH_SUNDAYS) {
      const date = sundays[cursor];
      const dateKey = Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd");
      cursor++;

      if (existing.has(dateKey)) continue;

      const monthKey = dateKey.slice(0, 7);
      if (!countsByMonth.has(monthKey)) countsByMonth.set(monthKey, new Map());

      const assignments = {};
      const assignedEmailsForDate = new Set();

      for (const ministry of shuffle_([...CONFIG.MINISTRIES])) {
        const m = countsByMonth.get(monthKey);
        if (!m.has("ALL")) m.set("ALL", new Map());

        const picked = pickCandidate_({
          roster,
          ministry,
          countsMap: m.get("ALL"),
          maxPerMonth: CONFIG.MAX_ASSIGNMENTS_PER_PERSON_PER_MONTH_PER_MINISTRY,
          avoidConsecutive: true,
          assignedEmailsForDate,
          emailToDates,
          dateKey,
        });

        if (!picked) throw new Error(`No eligible volunteer for ${ministry} on ${dateKey}`);

        assignments[ministry] = picked;
        assignedEmailsForDate.add(picked);
        m.get("ALL").set(picked, (m.get("ALL").get(picked) || 0) + 1);
        if (!emailToDates.has(picked)) emailToDates.set(picked, new Set());
        emailToDates.get(picked).add(dateKey);
      }

      const event = cal.createAllDayEvent("Tech Duty", date, {
        description: buildSundayDescription_(dateKey, assignments),
      });

      if (TEST_EMAIL_ONLY) {
        event.addGuest(TEST_EMAIL_ONLY);
      } else {
        new Set(Object.values(assignments)).forEach(e => event.addGuest(e));
      }

      rows.push([
        dateKey,
        assignments.Audio,
        assignments.LiveStream,
        assignments.PPT,
        event.getId(),
        CONFIG.CALENDAR_NAME,
        assignments.Audio,
        assignments.LiveStream,
        assignments.PPT,
        "",
        "",
        "",
        "",
        "",
      ]);

      existing.set(dateKey, true);
      created++;
    }

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    props.setProperty("SCHED_CURSOR_SUNDAY", cursor);
    if (cursor >= sundays.length) {
      props.deleteProperty("SCHED_CURSOR_SUNDAY");
      stopAutoBatchScheduler();
    }
  } finally {
    lock.releaseLock();
  }
}

// =====================
// SYNC CHANGES (manual scan all)
// =====================
function syncScheduleChanges() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId_());
    const sh = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!sh) throw new Error(`Missing sheet: ${CONFIG.SCHEDULE_SHEET_NAME}`);

    ensureScheduleHeaders_(sh);

    for (let r = 2; r <= sh.getLastRow(); r++) {
      syncScheduleChangesForRow_(ss, sh, r);
    }
  } finally {
    lock.releaseLock();
  }
}

// =====================
// RSVP SYNC (ICS responses)
// =====================
function setupRsvpSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "syncRsvpStatuses")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("syncRsvpStatuses")
    .timeBased()
    .everyMinutes(CONFIG.RSVP_SYNC_EVERY_MINUTES)
    .create();

  SpreadsheetApp.getUi().alert("RSVP sync trigger installed (ICS responses).");
}

function syncRsvpStatuses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId_());
    const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!schedule) throw new Error(`Missing sheet: ${CONFIG.SCHEDULE_SHEET_NAME}`);

    ensureScheduleHeaders_(schedule);

    const header = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
    const H = headerIndex_(header);
    const required = [
      "Date", "Event Id",
      "Audio", "LiveStream", "PPT",
      "Audio RSVP", "LiveStream RSVP", "PPT RSVP",
    ];
    required.forEach(k => {
      if (H[k] == null) throw new Error(`Missing column in Schedule: ${k}`);
    });

    const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);
    const lastRow = schedule.getLastRow();
    for (let r = 2; r <= lastRow; r++) {
      const dateKey = normalizeDateKey_(schedule.getRange(r, H["Date"] + 1).getValue());
      if (!dateKey) continue;

      const eventId = String(schedule.getRange(r, H["Event Id"] + 1).getValue() || "").trim();
      if (!eventId) continue;

      const event = safeGetEventById_(cal, eventId);
      if (!event) continue;

      const emailA = String(schedule.getRange(r, H["Audio"] + 1).getValue() || "").trim();
      const emailL = String(schedule.getRange(r, H["LiveStream"] + 1).getValue() || "").trim();
      const emailP = String(schedule.getRange(r, H["PPT"] + 1).getValue() || "").trim();

      const newA = getRsvpForEmail_(event, emailA);
      const newL = getRsvpForEmail_(event, emailL);
      const newP = getRsvpForEmail_(event, emailP);

      const curA = String(schedule.getRange(r, H["Audio RSVP"] + 1).getValue() || "").trim();
      const curL = String(schedule.getRange(r, H["LiveStream RSVP"] + 1).getValue() || "").trim();
      const curP = String(schedule.getRange(r, H["PPT RSVP"] + 1).getValue() || "").trim();

      if (newA !== curA) schedule.getRange(r, H["Audio RSVP"] + 1).setValue(newA);
      if (newL !== curL) schedule.getRange(r, H["LiveStream RSVP"] + 1).setValue(newL);
      if (newP !== curP) schedule.getRange(r, H["PPT RSVP"] + 1).setValue(newP);

      updateMonthlyRsvpForDate_(ss, dateKey, { Audio: newA, LiveStream: newL, PPT: newP });
    }
  } finally {
    lock.releaseLock();
  }
}

function fixMonthlyRsvpColumns() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  MONTH_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    ensureMonthHeadersOptionB_(sh);
    applyPrettyMonthFormattingOptionB_(sh);
  });
  SpreadsheetApp.getUi().alert("Monthly RSVP columns fixed/visible.");
}

function safeGetEventById_(cal, eventId) {
  try {
    return cal.getEventById(eventId);
  } catch (_) {
    return null;
  }
}

function getRsvpForEmail_(event, email) {
  if (!email) return "";
  const target = String(email).trim().toLowerCase();
  const guest = event.getGuestList().find(g => String(g.getEmail() || "").trim().toLowerCase() === target);
  if (!guest) return "";
  return normalizeGuestStatus_(guest.getGuestStatus());
}

function normalizeGuestStatus_(status) {
  if (status === CalendarApp.GuestStatus.YES) return "Yes";
  if (status === CalendarApp.GuestStatus.NO) return "No";
  if (status === CalendarApp.GuestStatus.MAYBE) return "Maybe";
  if (status === CalendarApp.GuestStatus.INVITED) return "Invited";
  return "";
}

function updateMonthlyRsvpForDate_(ss, dateKey, rsvpByRole) {
  if (!dateKey) return;
  const monthKey = dateKey.slice(0, 7);
  const sheetName = monthNameFromKey_(monthKey);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  ensureMonthHeadersOptionB_(sh);
  const row = findRowByDateKeyInMonthSheet_(sh, dateKey);
  if (!row) return;

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);

  const roleToCol = {
    Audio: "Audio RSVP",
    LiveStream: "Livestream RSVP",
    PPT: "PPT RSVP",
  };

  Object.keys(roleToCol).forEach(role => {
    const colName = roleToCol[role];
    const colIdx = H[colName];
    if (colIdx == null) return;
    const newVal = rsvpByRole[role];
    if (newVal == null) return; // allow partial updates
    const cell = sh.getRange(row, colIdx + 1);
    const curVal = String(cell.getValue() || "").trim();
    if (String(newVal) !== curVal) cell.setValue(newVal);
  });
}

function findRowByDateKeyInMonthSheet_(sh, dateKey) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  for (let r = 2; r <= lastRow; r++) {
    const dateVal = sh.getRange(r, 1).getValue();
    if (normalizeDateKey_(dateVal) === dateKey) return r;
  }
  return null;
}

// =====================
// AUTO-SYNC ON EDIT (Schedule)
// =====================
function setupAutoSyncOnEdit() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onScheduleEdit")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onScheduleEdit")
    .forSpreadsheet(SpreadsheetApp.openById(getSpreadsheetId_()))
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert("Auto-sync on Schedule edit installed.");
}

function onScheduleEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SCHEDULE_SHEET_NAME) return;

    const row = e.range.getRow();
    if (row <= 1) return;

    const props = PropertiesService.getScriptProperties();
    if (props.getProperty("SCHED_EDIT_GUARD") === "1") return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return;

    try {
      ensureScheduleHeaders_(sh);
      const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const H = headerIndex_(header);
      const roleColumns = getRoleColumns_();
      const col = e.range.getColumn();

      const editedRole = roleColumns.find(r => H[r] + 1 === col);
      const ss = SpreadsheetApp.openById(getSpreadsheetId_());

      if (editedRole) {
        const email = String(e.range.getValue() || "").trim();
        const reason = validateAssignmentForScheduleRow_(sh, row, editedRole, email, {
          enforceConsecutive: true,
        });
        if (reason) {
          props.setProperty("SCHED_EDIT_GUARD", "1");
          e.range.setValue(e.oldValue || "");
          props.deleteProperty("SCHED_EDIT_GUARD");
          (e.source || SpreadsheetApp.getActive()).toast(reason, "Tech Scheduler", 8);
          return;
        }

        const oldVal = String(e.oldValue || "").trim();
        const newVal = String(e.range.getValue() || "").trim();
        if (oldVal !== newVal) {
          const dateKey = normalizeDateKey_(sh.getRange(row, H["Date"] + 1).getValue());
          props.setProperty("SCHED_EDIT_GUARD", "1");
          props.setProperty("MONTH_EDIT_GUARD", "1");
          try {
            resetScheduleRsvpForRole_(sh, row, editedRole);
            resetMonthlyRsvpForRoleByDate_(ss, dateKey, editedRole);
          } finally {
            props.deleteProperty("SCHED_EDIT_GUARD");
            props.deleteProperty("MONTH_EDIT_GUARD");
          }
        }
      }

      syncScheduleChangesForRow_(ss, sh, row);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error("onScheduleEdit error:", err);
  }
}

/**
 * Sync a single Schedule row:
 * - ensures calendar event exists (auto-create if missing)
 * - compares new vs last synced
 * - updates event description
 * - sends emails:
 *    - if oldEmail+newEmail -> send Tech Sub Confirmation + Tech Sub Invite
 *    - else -> send general schedule update
 * - updates last synced fields
 */
function syncScheduleChangesForRow_(ss, sh, row) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  const rosterMaps = buildRosterMaps_(ss);
  const H = headerIndex_(data[0]);

  const required = [
    "Date", "Event Id",
    "Audio", "LiveStream", "PPT",
    "Last Synced Audio", "Last Synced LiveStream", "Last Synced PPT",
  ];
  required.forEach(k => {
    if (H[k] == null) throw new Error(`Missing column in Schedule: ${k}`);
  });

  const dateKey = normalizeDateKey_(data[row - 1][H["Date"]]);
  if (!dateKey) return;

  // Auto-create calendar event if missing / deleted
  ensureCalendarEventForScheduleRow_(sh, row);

  // Re-read event id after ensure
  const eventId = String(sh.getRange(row, H["Event Id"] + 1).getValue() || "").trim();
  if (!eventId) return;

  const newA = String(sh.getRange(row, H["Audio"] + 1).getValue() || "").trim();
  const newL = String(sh.getRange(row, H["LiveStream"] + 1).getValue() || "").trim();
  const newP = String(sh.getRange(row, H["PPT"] + 1).getValue() || "").trim();

  const oldA = String(sh.getRange(row, H["Last Synced Audio"] + 1).getValue() || "").trim();
  const oldL = String(sh.getRange(row, H["Last Synced LiveStream"] + 1).getValue() || "").trim();
  const oldP = String(sh.getRange(row, H["Last Synced PPT"] + 1).getValue() || "").trim();

  const changes = [];
  if (newA !== oldA) changes.push({ role: "Audio", oldEmail: oldA, newEmail: newA });
  if (newL !== oldL) changes.push({ role: "LiveStream", oldEmail: oldL, newEmail: newL });
  if (newP !== oldP) changes.push({ role: "PPT", oldEmail: oldP, newEmail: newP });

  if (H["Changed At"] != null) sh.getRange(row, H["Changed At"] + 1).setValue(new Date());
  if (!changes.length) return;

  const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);
  const event = cal.getEventById(eventId);
  if (!event) throw new Error(`Event not found for row ${row}. Event Id: ${eventId}`);

  // Always update description to reflect latest assignments
  event.setDescription(buildSundayDescription_(dateKey, { Audio: newA, LiveStream: newL, PPT: newP }));

  // Send emails (sub emails when old+new)
  sendChangeEmailsSmart_(ss, dateKey, changes, eventId);

  // Update last synced
  sh.getRange(row, H["Last Synced Audio"] + 1).setValue(newA);
  sh.getRange(row, H["Last Synced LiveStream"] + 1).setValue(newL);
  sh.getRange(row, H["Last Synced PPT"] + 1).setValue(newP);

  if (H["Notified At"] != null) sh.getRange(row, H["Notified At"] + 1).setValue(new Date());
}

// =====================
// WEEKLY REMINDER
// =====================
function setupSaturday5pmTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendSaturdayReminder")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendSaturdayReminder")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(19)
    .create();

  SpreadsheetApp.getUi().alert("Wednesday 7PM reminder trigger installed.");
}

function sendSaturdayReminder() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const sh = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  const rosterMaps = buildRosterMaps_(ss);
  const H = headerIndex_(data[0]);
  const nextSunday = getNextSunday_(new Date(), CONFIG.TIMEZONE);
  const dateKey = Utilities.formatDate(nextSunday, CONFIG.TIMEZONE, "yyyy-MM-dd");
  const pretty = Utilities.formatDate(nextSunday, CONFIG.TIMEZONE, "EEE, MMM d, yyyy");
  const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);

  for (let r = 1; r < data.length; r++) {
    if (normalizeDateKey_(data[r][H["Date"]]) !== dateKey) continue;

    ensureCalendarEventForScheduleRow_(sh, r + 1);
    const eventId = String(data[r][H["Event Id"]] || "").trim();
    if (!eventId) return;

    const event = cal.getEventById(eventId);
    if (!event) return;

    event.setDescription(buildSundayDescription_(dateKey, {
      Audio: data[r][H["Audio"]],
      LiveStream: data[r][H["LiveStream"]],
      PPT: data[r][H["PPT"]],
    }));

    const rolesByEmail = new Map();
    addRole_(rolesByEmail, data[r][H["Audio"]], "Audio");
    addRole_(rolesByEmail, data[r][H["LiveStream"]], "LiveStream");
    addRole_(rolesByEmail, data[r][H["PPT"]], "PPT");

    for (const [email, roles] of rolesByEmail) {
      const to = TEST_EMAIL_ONLY || email;
      const displayName = rosterMaps.emailToName.get(String(email || "").trim().toLowerCase()) || email;
      const subject = `Sunday Tech Duty Reminder | ${CONFIG.TECH_TEAM_NAME}`;
      const htmlBody = buildPrettyReminderEmail_({ prettyDate: pretty, roles, displayName });
      const summary = `Tech Duty - Your role: ${roles.join(", ")}`;
      ensureGuestOnEvent_(event, to);
      sendCalendarInviteEmail_({ toEmail: to, subject, htmlBody, event, summaryOverride: summary });

      const mobile = rosterMaps.emailToMobile.get(String(email || "").trim().toLowerCase()) || "";
      if (mobile) {
        const sms = buildSmsSaturdayReminder_({ prettyDate: pretty, roles });
        sendSms_({ recipient: mobile, message: sms });
      }
    }
    break;
  }
}

// =====================
// EMAIL QUOTA
// =====================
function logEmailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  Logger.log(`Remaining daily email quota: ${remaining}`);
  try {
    SpreadsheetApp.getActive().toast(`Remaining daily email quota: ${remaining}`, "Tech Scheduler", 6);
  } catch (_) {}
}

// =====================
// MONTHLY REMINDER (last day of month)
// =====================
function setupMonthlyScheduleReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendMonthlyScheduleReminderIfLastDay")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendMonthlyScheduleReminderIfLastDay")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  SpreadsheetApp.getUi().alert("Monthly reminder trigger installed (runs daily, sends only on last day).");
}

function sendMonthlyScheduleReminderIfLastDay() {
  const today = new Date();
  if (!isLastDayOfMonth_(today, CONFIG.TIMEZONE)) return;
  sendMonthlyScheduleReminder_();
}

function sendMonthlyScheduleReminder_() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const sh = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  const H = headerIndex_(data[0]);
  const roles = getRoleColumns_();
  const rosterMaps = buildRosterMaps_(ss);
  const eventIdMap = buildEventIdMap_(sh);

  const now = new Date(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd") + "T00:00:00");
  const nextMonth = new Date(now);
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const monthKey = Utilities.formatDate(nextMonth, CONFIG.TIMEZONE, "yyyy-MM");
  const monthLabel = Utilities.formatDate(nextMonth, CONFIG.TIMEZONE, "MMMM yyyy");

  const byEmail = new Map();

  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey || !dateKey.startsWith(monthKey)) continue;

    roles.forEach(role => {
      const email = String(data[r][H[role]] || "").trim();
      if (!email) return;
      if (!byEmail.has(email)) byEmail.set(email, new Map());
      const dateMap = byEmail.get(email);
      if (!dateMap.has(dateKey)) dateMap.set(dateKey, []);
      dateMap.get(dateKey).push(role);
    });
  }

  if (!byEmail.size) return;

  byEmail.forEach((dateMap, email) => {
    const dateKeys = Array.from(dateMap.keys()).sort();
    const items = dateKeys.map(k => {
      const pretty = Utilities.formatDate(new Date(k + "T00:00:00"), CONFIG.TIMEZONE, "EEE, MMM d, yyyy");
      const roleBlocks = dateMap.get(k).map(role => {
        const eventId = eventIdMap.get(k) || "";
        const yesLink = eventId
          ? buildWebAppRsvpLink_({ action: "YES", role, dateKey: k, assigneeEmail: email, eventId })
          : "#";
        const noLink = eventId
          ? buildWebAppRsvpLink_({ action: "NO", role, dateKey: k, assigneeEmail: email, eventId })
          : "#";

        return `
          <div style="margin:10px 0;padding:12px;border:1px solid #e5e7eb;border-radius:12px;">
            <div style="font-weight:800;margin-bottom:6px;">${escapeHtml_(pretty)}</div>
            <div style="margin:0 0 10px 0;">Your role: <b>${escapeHtml_(role)}</b></div>
            <div style="display:flex;gap:10px;margin:0;">
              <a href="${yesLink}"
                style="display:inline-block;background:#111827;color:#ffffff;text-decoration:none;
                  padding:10px 14px;border-radius:10px;font-weight:800;">
                Yes
              </a>
              <a href="${noLink}"
                style="display:inline-block;background:#ffffff;color:#111827;text-decoration:none;
                  padding:10px 14px;border-radius:10px;font-weight:800;border:1px solid #e5e7eb;">
                No
              </a>
            </div>
          </div>
        `;
      }).join("");
      return roleBlocks;
    }).join("");

    const displayName = rosterMaps.emailToName.get(email) || "there";
    const html = buildPrettyEmail_({
      title: `Your Volunteer Schedule for the Month of ${monthLabel}`,
      subtitle: "",
      bodyHtml: `
        <div style="margin:0 0 10px 0;">
          Hi ${escapeHtml_(displayName)},<br/>
          Thank you for serving. Here is your schedule for next month:
        </div>
        ${items}
        <div style="margin:12px 0 0 0;color:#475569;">
          Please tap Yes or No for each assignment above.
        </div>
        <div style="margin:12px 0 0 0;color:#475569;">
          Please confirm your availability by clicking “Yes.” If you have a conflict, click “No” and kindly inform us via the Tech group chat as soon as possible.
          Thank you for your continued service in our church’s tech ministry.
        </div>
      `,
    });

    const subject = `Next Month Schedule | ${CONFIG.TECH_TEAM_NAME}`;
    GmailApp.sendEmail(TEST_EMAIL_ONLY || email, subject, "Please view this email in HTML.", { htmlBody: html });

    const mobile = rosterMaps.emailToMobile.get(String(email || "").trim().toLowerCase()) || "";
    if (mobile) {
      const sms = buildSmsMonthlyReminder_({ monthLabel, count: dateKeys.length });
      sendSms_({ recipient: mobile, message: sms });
    }
  });
}

// =========================================================
// OPTION B — MONTHLY SHEETS (Name dropdown + hidden email)
// =========================================================

// 1) Generate month sheets (with hidden email columns)
function generateMonthlySheetsOptionB() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const schedule = getOrCreateSheet_(ss, CONFIG.SCHEDULE_SHEET_NAME, getScheduleHeaders_());
  ensureScheduleHeaders_(schedule);

  const scheduleMap = buildScheduleMap_(schedule); // dateKey -> emails
  const rosterMaps = buildRosterMaps_(ss);         // name<->email + eligible names per role

  const sundays = getSundays_(CONFIG.START_DATE, CONFIG.END_DATE);
  const monthsUsed = new Set(sundays.map(d => Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM")));

  Array.from(monthsUsed).sort().forEach(monthKey => {
    const sheetName = monthNameFromKey_(monthKey);
    buildOneMonthSheetOptionB_(ss, sheetName, monthKey, scheduleMap, rosterMaps.emailToName);
  });

  applyMonthlyDropdownsOptionB();
  SpreadsheetApp.getUi().alert("Monthly sheets generated .");
}

// 2) Apply dropdowns from Roster (names by eligibility)
function applyMonthlyDropdownsOptionB() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const rosterMaps = buildRosterMaps_(ss);

  MONTH_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    ensureMonthHeadersOptionB_(sh);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const baseCol = sh.getLastColumn() + 1;
    sh.getRange(1, baseCol).setValue("_dropdown_sources_");
    sh.getRange(2, baseCol, Math.max(lastRow, 60), 3).clearContent();

    writeList_(sh, baseCol, rosterMaps.namesByRole.Audio);
    writeList_(sh, baseCol + 1, rosterMaps.namesByRole.LiveStream);
    writeList_(sh, baseCol + 2, rosterMaps.namesByRole.PPT);

    sh.hideColumns(baseCol, 3);

    const audioRange = sh.getRange(2, baseCol, Math.max(rosterMaps.namesByRole.Audio.length, 1), 1);
    const liveRange  = sh.getRange(2, baseCol + 1, Math.max(rosterMaps.namesByRole.LiveStream.length, 1), 1);
    const pptRange   = sh.getRange(2, baseCol + 2, Math.max(rosterMaps.namesByRole.PPT.length, 1), 1);

    const audioRule = SpreadsheetApp.newDataValidation().requireValueInRange(audioRange, true).setAllowInvalid(true).build();
    const liveRule  = SpreadsheetApp.newDataValidation().requireValueInRange(liveRange, true).setAllowInvalid(true).build();
    const pptRule   = SpreadsheetApp.newDataValidation().requireValueInRange(pptRange, true).setAllowInvalid(true).build();

    sh.getRange(2, 2, lastRow - 1, 1).setDataValidation(audioRule);
    sh.getRange(2, 3, lastRow - 1, 1).setDataValidation(liveRule);
    sh.getRange(2, 4, lastRow - 1, 1).setDataValidation(pptRule);

    applyPrettyMonthFormattingOptionB_(sh);
  });

  try {
    SpreadsheetApp.getUi().alert("Monthly dropdowns applied .");
  } catch (_) {}
}

// 3) Sync month sheets from Schedule
function syncMonthlySheetsFromScheduleOptionB() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedule) throw new Error(`Missing sheet: ${CONFIG.SCHEDULE_SHEET_NAME}`);

  ensureScheduleHeaders_(schedule);

  const scheduleMap = buildScheduleMap_(schedule);
  const rosterMaps = buildRosterMaps_(ss);

  MONTH_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    ensureMonthHeadersOptionB_(sh);
    syncOneMonthSheetOptionB_(sh, scheduleMap, rosterMaps.emailToName);
    applyPrettyMonthFormattingOptionB_(sh);
  });

  SpreadsheetApp.getUi().alert("Monthly sheets synced from Schedule .");
}

// 4) Setup monthly onEdit sync trigger
function setupMonthlyEditSyncOptionB() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onMonthlyEditOptionB")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onMonthlyEditOptionB")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "Monthly edit sync trigger installed.\n\nNote: Pop-up confirmation is not allowed on edit.\nUse: Tech Scheduler → Confirm & Send Emails"
  );
}

/**
 * Monthly edit handler:
 * - updates hidden email (E/F/G) and Schedule (email)
 * - stores a pending change (for sidebar confirmation)
 * - DOES NOT send emails automatically
 */
function onMonthlyEditOptionB(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    const ss = e.source || SpreadsheetApp.openById(getSpreadsheetId_());
    const sheetName = sh.getName();
    if (!MONTH_SHEETS.includes(sheetName)) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Only respond to dropdown columns: B/C/D
    if (row <= 1) return;
    if (col < 2 || col > 4) return;

    const props = PropertiesService.getScriptProperties();
    if (props.getProperty("MONTH_EDIT_GUARD") === "1") return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return;

    try {
      const dateVal = sh.getRange(row, 1).getValue();
      if (!dateVal) return;

      const dateKey = normalizeDateKey_(dateVal);
      const prettyDate = Utilities.formatDate(new Date(dateKey + "T00:00:00"), CONFIG.TIMEZONE, "EEE, MMM d, yyyy");

      const roleLabel = col === 2 ? "Audio" : col === 3 ? "LiveStream" : "PPT";
      const newName = String(e.range.getValue() || "").trim();
      const oldName = String(e.oldValue || "").trim();

      const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
      if (!schedule) throw new Error("Schedule sheet not found.");
      ensureScheduleHeaders_(schedule);

      const scheduleRow = findOrCreateScheduleRowByDate_(schedule, dateKey);

      const header = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
      const H = headerIndex_(header);
      const oldEmail = String(schedule.getRange(scheduleRow, H[roleLabel] + 1).getValue() || "").trim();

      const rosterMaps = buildRosterMaps_(ss);
      const pickedEmail = rosterMaps.nameToEmail.get(newName) || "";
      if (newName && !pickedEmail) {
        ss.toast(`No email found for "${newName}" in Roster.`, "Tech Scheduler", 6);
      }

      if (pickedEmail) {
        const reason = validateAssignmentForScheduleRow_(schedule, scheduleRow, roleLabel, pickedEmail, {
          enforceConsecutive: false,
          ignoreMaxPerMonth: true,
        });
        if (reason) {
          props.setProperty("MONTH_EDIT_GUARD", "1");
          sh.getRange(row, col).setValue(oldName || "");
          sh.getRange(row, col + 3).setValue(oldEmail || "");
          props.deleteProperty("MONTH_EDIT_GUARD");
          ss.toast(reason, "Tech Scheduler", 8);
          return;
        }
        const warn = getConsecutiveWarning_(schedule, scheduleRow, pickedEmail);
        if (warn) ss.toast(warn, "Tech Scheduler", 8);
      }

      // Hidden email (B→E, C→F, D→G)
      props.setProperty("MONTH_EDIT_GUARD", "1");
      sh.getRange(row, col + 3).setValue(pickedEmail);
      props.deleteProperty("MONTH_EDIT_GUARD");

      // Update Schedule (source of truth)
      schedule.getRange(scheduleRow, H[roleLabel] + 1).setValue(pickedEmail);

      // Reset RSVP for this role (monthly + schedule) to avoid stale status/color
      resetMonthlyRsvpForRole_(sh, row, roleLabel);
      resetScheduleRsvpForRole_(schedule, scheduleRow, roleLabel);

      // Save pending change for confirmation UX
      addPendingChange_({
        sheetName,
        row,
        col,
        dateKey,
        prettyDate,
        role: roleLabel,
        fromName: oldName || "",
        toName: newName || "",
        fromEmail: oldEmail || "",
        toEmail: pickedEmail || "",
        createdAt: new Date().toISOString(),
      });

      ss.toast(
        `Saved: ${roleLabel} updated for ${prettyDate}.\nAction needed: Tech Scheduler → Confirm & Send Emails`,
        "Tech Scheduler",
        8
      );
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    try {
      SpreadsheetApp.getActive().toast(String(err), "Monthly Sync Error", 10);
    } catch (_) {}
    console.error("onMonthlyEditOptionB error:", err);
  }
}

function resetMonthlyRsvpForRole_(monthSheet, row, roleLabel) {
  const header = monthSheet.getRange(1, 1, 1, monthSheet.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);

  const roleCol =
    roleLabel === "Audio" ? 2 :
    roleLabel === "LiveStream" ? 3 :
    roleLabel === "PPT" ? 4 :
    null;
  const rsvpColName =
    roleLabel === "Audio" ? "Audio RSVP" :
    roleLabel === "LiveStream" ? "Livestream RSVP" :
    roleLabel === "PPT" ? "PPT RSVP" :
    "";
  const rsvpCol = rsvpColName && H[rsvpColName] != null ? H[rsvpColName] + 1 : null;

  if (rsvpCol) {
    monthSheet.getRange(row, rsvpCol).setValue("");
    monthSheet.getRange(row, rsvpCol).setBackground(null).setFontColor(null);
  }
  if (roleCol) {
    monthSheet.getRange(row, roleCol).setBackground(null).setFontColor(null);
  }
}

function resetScheduleRsvpForRole_(scheduleSheet, row, roleLabel) {
  const header = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);
  const rsvpColName =
    roleLabel === "Audio" ? "Audio RSVP" :
    roleLabel === "LiveStream" ? "LiveStream RSVP" :
    roleLabel === "PPT" ? "PPT RSVP" :
    "";
  if (!rsvpColName || H[rsvpColName] == null) return;
  scheduleSheet.getRange(row, H[rsvpColName] + 1).setValue("");
}

function resetMonthlyRsvpForRoleByDate_(ss, dateKey, roleLabel) {
  if (!dateKey) return;
  const monthKey = dateKey.slice(0, 7);
  const sheetName = monthNameFromKey_(monthKey);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  ensureMonthHeadersOptionB_(sh);
  const row = findRowByDateKeyInMonthSheet_(sh, dateKey);
  if (!row) return;
  resetMonthlyRsvpForRole_(sh, row, roleLabel);
}

// =====================
// CONFIRM UX (Sidebar)
// =====================
function openConfirmSendSidebar() {
  const pending = getPendingChanges_();
  if (!pending.length) {
    SpreadsheetApp.getActive().toast("No pending change to confirm.", "Tech Scheduler", 5);
    return;
  }

  const safe = s => escapeHtml_(s || "");
  const items = pending.map(p => `
    <div style="padding:10px 0;border-bottom:1px solid #e5e7eb;">
      <div style="font-size:12px;color:#64748b;margin-bottom:4px;">
        ${safe(p.prettyDate || p.dateKey)} • ${safe(p.role)}
      </div>
      <div style="font-size:13px;line-height:1.5;">
        <span style="color:#64748b;">From:</span> <b>${safe(p.fromName || "(blank)")}</b><br/>
        <span style="color:#64748b;">To:</span> <b>${safe(p.toName || "(blank)")}</b>
      </div>
    </div>
  `).join("");

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,Helvetica,sans-serif;padding:14px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
        <div style="width:10px;height:10px;border-radius:999px;background:#f59e0b;"></div>
        <div style="font-size:14px;color:#0f172a;font-weight:800;">Confirm & Send Emails</div>
      </div>

      <div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:12px;padding:12px;margin-bottom:12px;">
        <div style="font-size:13px;color:#475569;margin-bottom:8px;">
          You made ${pending.length} schedule change(s). Sending will:
        </div>
        <ul style="margin:0;padding-left:18px;font-size:13px;color:#0f172a;line-height:1.6;">
          <li>Update the Calendar event description</li>
          <li>Send notification emails to affected volunteers</li>
        </ul>
      </div>

      <div style="border:1px solid #e5e7eb;border-radius:12px;padding:12px;max-height:240px;overflow:auto;">
        ${items}
      </div>

      <div style="display:flex;gap:10px;margin-top:14px;">
        <button
          onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).confirmAndSendEmails()"
          style="flex:1;background:#111827;color:#fff;border:none;border-radius:10px;padding:10px 12px;font-weight:800;cursor:pointer;">
          Send Emails
        </button>

        <button
          onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).cancelPendingChange()"
          style="flex:1;background:#fff;color:#111827;border:1px solid #e5e7eb;border-radius:10px;padding:10px 12px;font-weight:800;cursor:pointer;">
          Cancel
        </button>
      </div>

      <div style="margin-top:10px;font-size:12px;color:#94a3b8;">
        Tip: This prevents accidental email notifications when editing dropdowns.
      </div>
    </div>
  `).setTitle("Tech Scheduler");

  SpreadsheetApp.getUi().showSidebar(html);
}

function confirmAndSendEmails() {
  const pending = getPendingChanges_();
  if (!pending.length) {
    SpreadsheetApp.getActive().toast("No pending change.", "Tech Scheduler", 5);
    return;
  }

  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedule) throw new Error("Schedule sheet not found.");
  ensureScheduleHeaders_(schedule);

  const dateKeys = new Set(pending.map(p => p.dateKey).filter(Boolean));
  dateKeys.forEach(dateKey => {
    const scheduleRow = findOrCreateScheduleRowByDate_(schedule, dateKey);
    syncScheduleChangesForRow_(ss, schedule, scheduleRow);
  });

  clearPendingChanges_();
  ss.toast("Confirmed ✅ Calendar updated and emails sent.", "Tech Scheduler", 6);
}

function cancelPendingChange() {
  const pending = getPendingChanges_();
  if (!pending.length) {
    SpreadsheetApp.getActive().toast("No pending change.", "Tech Scheduler", 5);
    return;
  }

  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const cells = new Map();
  pending.forEach(p => {
    const key = `${p.sheetName}|${p.row}|${p.col}`;
    if (!cells.has(key)) cells.set(key, p);
  });

  const props = PropertiesService.getScriptProperties();
  props.setProperty("MONTH_EDIT_GUARD", "1");
  cells.forEach(p => {
    const sh = ss.getSheetByName(p.sheetName);
    if (!sh) return;
    sh.getRange(p.row, p.col).setValue(p.fromName || "");
    // restore hidden email too
    try {
      sh.getRange(p.row, p.col + 3).setValue(p.fromEmail || "");
    } catch (_) {}
  });
  props.deleteProperty("MONTH_EDIT_GUARD");

  clearPendingChanges_();
  ss.toast("Cancelled. No emails sent.", "Tech Scheduler", 6);
}

function addPendingChange_(obj) {
  const list = getPendingChanges_();
  list.push(obj);
  PropertiesService.getDocumentProperties().setProperty(PENDING_KEY, JSON.stringify(list));
}
function getPendingChanges_() {
  const s = PropertiesService.getDocumentProperties().getProperty(PENDING_KEY);
  const list = s ? JSON.parse(s) : [];
  return Array.isArray(list) ? list : [];
}
function clearPendingChanges_() {
  PropertiesService.getDocumentProperties().deleteProperty(PENDING_KEY);
}

// ---------------------
// Month sheet builders 
// ---------------------
function buildOneMonthSheetOptionB_(ss, sheetName, monthKey, scheduleMap, emailToName) {
  const sh = getOrCreateSheet_(ss, sheetName, MONTH_HEADERS);
  sh.clear();

  sh.getRange(1, 1, 1, MONTH_HEADERS.length).setValues([MONTH_HEADERS]);

  const sundays = getSundaysForMonthKey_(monthKey);

  const rows = sundays.map(d => {
    const dateKey = Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd");
    const entry = scheduleMap.get(dateKey) || {
      Audio: "",
      LiveStream: "",
      PPT: "",
      AudioRsvp: "",
      LiveStreamRsvp: "",
      PptRsvp: "",
    };

    const aEmail = entry.Audio || "";
    const lEmail = entry.LiveStream || "";
    const pEmail = entry.PPT || "";

    return [
      d,
      emailToName.get(aEmail) || "",
      emailToName.get(lEmail) || "",
      emailToName.get(pEmail) || "",
      aEmail,
      lEmail,
      pEmail,
      entry.AudioRsvp || "",
      entry.LiveStreamRsvp || "",
      entry.PptRsvp || "",
    ];
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, MONTH_HEADERS.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat(CONFIG.MONTH_DATE_FORMAT);
  }

  sh.hideColumns(5, 3);
  applyPrettyMonthFormattingOptionB_(sh);
  sh.setFrozenRows(1);
}

function syncOneMonthSheetOptionB_(sh, scheduleMap, emailToName) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  for (let r = 2; r <= lastRow; r++) {
    const dateVal = sh.getRange(r, 1).getValue();
    const dateKey = normalizeDateKey_(dateVal);
    if (!dateKey) continue;

    const entry = scheduleMap.get(dateKey) || {
      Audio: "",
      LiveStream: "",
      PPT: "",
      AudioRsvp: "",
      LiveStreamRsvp: "",
      PptRsvp: "",
    };

    const aEmail = entry.Audio || "";
    const lEmail = entry.LiveStream || "";
    const pEmail = entry.PPT || "";

    sh.getRange(r, 2).setValue(emailToName.get(aEmail) || "");
    sh.getRange(r, 3).setValue(emailToName.get(lEmail) || "");
    sh.getRange(r, 4).setValue(emailToName.get(pEmail) || "");

    sh.getRange(r, 5).setValue(aEmail);
    sh.getRange(r, 6).setValue(lEmail);
    sh.getRange(r, 7).setValue(pEmail);

    sh.getRange(r, 8).setValue(entry.AudioRsvp || "");
    sh.getRange(r, 9).setValue(entry.LiveStreamRsvp || "");
    sh.getRange(r, 10).setValue(entry.PptRsvp || "");
  }

  sh.getRange(2, 1, lastRow - 1, 1).setNumberFormat(CONFIG.MONTH_DATE_FORMAT);
  sh.hideColumns(5, 3);
}

function ensureMonthHeadersOptionB_(sh) {
  if (sh.getLastColumn() < MONTH_HEADERS.length) {
    sh.insertColumnsAfter(sh.getLastColumn(), MONTH_HEADERS.length - sh.getLastColumn());
  }

  const firstRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), MONTH_HEADERS.length)).getValues()[0];
  const isBlank = firstRow.every(v => String(v || "").trim() === "");
  if (isBlank) {
    sh.getRange(1, 1, 1, MONTH_HEADERS.length).setValues([MONTH_HEADERS]);
    return;
  }

  const H = headerIndex_(firstRow);
  const hasDate = H["Date"] != null;
  const hasAudioEmail = H["Audio Email"] != null;

  if (!hasDate || !hasAudioEmail) {
    sh.getRange(1, 1, 1, MONTH_HEADERS.length).setValues([MONTH_HEADERS]);
    return;
  }

  const hasAll = MONTH_HEADERS.every(h => H[h] != null);
  if (!hasAll) {
    sh.getRange(1, 1, 1, MONTH_HEADERS.length).setValues([MONTH_HEADERS]);
  }
}

function applyPrettyMonthFormattingOptionB_(sh) {
  sh.setRowHeight(1, 34);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(8, 120);
  sh.setColumnWidth(9, 120);
  sh.setColumnWidth(10, 120);

  const header = sh.getRange(1, 1, 1, MONTH_HEADERS.length);
  header
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#f8fafc");

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const body = sh.getRange(2, 1, lastRow - 1, MONTH_HEADERS.length);
    body.setVerticalAlignment("middle");
    body.setWrap(true);
    body.setBorder(true, true, true, true, true, true);

    for (let r = 2; r <= lastRow; r++) {
      sh.getRange(r, 1, 1, MONTH_HEADERS.length).setBackground((r % 2) === 0 ? "#ffffff" : "#fbfdff");
    }
    sh.getRange(2, 1, lastRow - 1, 1).setFontWeight("bold");
  }

  sh.setFrozenRows(1);
  sh.hideColumns(5, 3);
  sh.hideColumns(8, 3);

  if (lastRow < 2) return;

  const audioRoleRange = sh.getRange(2, 2, lastRow - 1, 1);
  const liveRoleRange = sh.getRange(2, 3, lastRow - 1, 1);
  const pptRoleRange = sh.getRange(2, 4, lastRow - 1, 1);
  const audioRsvpRange = sh.getRange(2, 8, lastRow - 1, 1);
  const liveRsvpRange = sh.getRange(2, 9, lastRow - 1, 1);
  const pptRsvpRange = sh.getRange(2, 10, lastRow - 1, 1);

  const greenBg = "#dcfce7";
  const greenText = "#166534";
  const redBg = "#fee2e2";
  const redText = "#991b1b";

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="Yes"')
      .setBackground(greenBg)
      .setFontColor(greenText)
      .setRanges([audioRsvpRange, audioRoleRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="No"')
      .setBackground(redBg)
      .setFontColor(redText)
      .setRanges([audioRsvpRange, audioRoleRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="Yes"')
      .setBackground(greenBg)
      .setFontColor(greenText)
      .setRanges([liveRsvpRange, liveRoleRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="No"')
      .setBackground(redBg)
      .setFontColor(redText)
      .setRanges([liveRsvpRange, liveRoleRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="Yes"')
      .setBackground(greenBg)
      .setFontColor(greenText)
      .setRanges([pptRsvpRange, pptRoleRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2="No"')
      .setBackground(redBg)
      .setFontColor(redText)
      .setRanges([pptRsvpRange, pptRoleRange])
      .build(),
  ];

  sh.setConditionalFormatRules(rules);
}

// =====================
// Roster maps (Name <-> Email/Mobile) + eligible names per ministry
// Roster headers required: Name | Email | Audio | LiveStream | PPT
// Optional: Mobile (for SMS)
// If no role is TRUE, the person is treated as backup (dropdowns only).
// =====================
function buildRosterMaps_(ss) {
  const sh = ss.getSheetByName(CONFIG.ROSTER_SHEET_NAME);
  if (!sh) throw new Error(`Missing sheet: ${CONFIG.ROSTER_SHEET_NAME}`);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error(`Roster sheet "${CONFIG.ROSTER_SHEET_NAME}" is empty.`);

  const H = headerIndex_(data[0]);
  if (H["Name"] == null) throw new Error(`Roster must have column: Name`);
  if (H["Email"] == null) throw new Error(`Roster must have column: Email`);

  const nameToEmail = new Map();
  const emailToName = new Map();
  const nameToMobile = new Map();
  const emailToMobile = new Map();

  const namesByRole = {
    Audio: [],
    LiveStream: [],
    PPT: [],
  };

  for (let r = 1; r < data.length; r++) {
    const name = String(data[r][H["Name"]] || "").trim();
    const email = String(data[r][H["Email"]] || "").trim().toLowerCase();
    const mobileRaw = H["Mobile"] != null ? data[r][H["Mobile"]] : "";
    const mobile = normalizePhone_(mobileRaw);
    if (!name || !email) continue;

    nameToEmail.set(name, email);
    emailToName.set(email, name);
    if (mobile) {
      nameToMobile.set(name, mobile);
      emailToMobile.set(email, mobile);
    }

    const isAudio = isTruthy_(data[r][H["Audio"]]);
    const isLive = isTruthy_(data[r][H["LiveStream"]]);
    const isPpt = isTruthy_(data[r][H["PPT"]]);
    const isBackup = !isAudio && !isLive && !isPpt;

    if (isAudio || isBackup) namesByRole.Audio.push(name);
    if (isLive || isBackup) namesByRole.LiveStream.push(name);
    if (isPpt || isBackup) namesByRole.PPT.push(name);
  }

  Object.keys(namesByRole).forEach(k => {
    namesByRole[k] = Array.from(new Set(namesByRole[k])).sort();
  });

  return { nameToEmail, emailToName, nameToMobile, emailToMobile, namesByRole };
}

function getNameFromEmail_(ss, email) {
  try {
    const roster = buildRosterMaps_(ss);
    return roster.emailToName.get(String(email || "").trim().toLowerCase()) || "";
  } catch (_) {
    return "";
  }
}

function writeList_(sh, col, list) {
  if (!list || !list.length) {
    sh.getRange(2, col, 1, 1).setValue("");
    return;
  }
  sh.getRange(2, col, list.length, 1).setValues(list.map(x => [x]));
}

// ---------------------
// Find or create Schedule row by dateKey
// ---------------------
function findOrCreateScheduleRowByDate_(schedule, dateKey) {
  const values = schedule.getDataRange().getValues();
  const H = headerIndex_(values[0]);
  if (H["Date"] == null) throw new Error(`Schedule missing column: Date`);

  for (let r = 1; r < values.length; r++) {
    const k = normalizeDateKey_(values[r][H["Date"]]);
    if (k === dateKey) return r + 1;
  }

  const newRow = new Array(values[0].length).fill("");
  newRow[H["Date"]] = dateKey;

  schedule.getRange(schedule.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);
  return schedule.getLastRow();
}

// =====================
// AUTO-CREATE EVENT IF MISSING
// =====================
function ensureCalendarEventForScheduleRow_(scheduleSheet, row) {
  ensureScheduleHeaders_(scheduleSheet);

  const header = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);

  const dateKey = normalizeDateKey_(scheduleSheet.getRange(row, H["Date"] + 1).getValue());
  if (!dateKey) throw new Error(`Row ${row}: Missing Date`);

  const eventIdCell = scheduleSheet.getRange(row, H["Event Id"] + 1);
  let eventId = String(eventIdCell.getValue() || "").trim();

  const calNameCell = scheduleSheet.getRange(row, H["Calendar Name"] + 1);
  if (!String(calNameCell.getValue() || "").trim()) calNameCell.setValue(CONFIG.CALENDAR_NAME);

  const cal = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);

  let event = null;
  if (eventId) {
    try {
      event = cal.getEventById(eventId);
    } catch (_) {
      event = null;
    }
  }

  if (!event) {
    const d = new Date(dateKey + "T00:00:00");
    const newA = String(scheduleSheet.getRange(row, H["Audio"] + 1).getValue() || "").trim();
    const newL = String(scheduleSheet.getRange(row, H["LiveStream"] + 1).getValue() || "").trim();
    const newP = String(scheduleSheet.getRange(row, H["PPT"] + 1).getValue() || "").trim();

    const assignments = { Audio: newA, LiveStream: newL, PPT: newP };

    event = cal.createAllDayEvent("Tech Duty", d, {
      description: buildSundayDescription_(dateKey, assignments),
    });

    if (TEST_EMAIL_ONLY) {
      event.addGuest(TEST_EMAIL_ONLY);
    } else {
      [newA, newL, newP].filter(Boolean).forEach(email => event.addGuest(email));
    }

    eventId = event.getId();
    eventIdCell.setValue(eventId);
  }

  return eventId;
}

// =====================
// HELPERS (core)
// =====================
function getSpreadsheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("Missing Script Property: SPREADSHEET_ID");
  return id;
}

function getScheduleHeaders_() {
  return [
    "Date", "Audio", "LiveStream", "PPT",
    "Event Id", "Calendar Name",
    "Last Synced Audio", "Last Synced LiveStream", "Last Synced PPT",
    "Changed At", "Notified At",
    "Audio RSVP", "LiveStream RSVP", "PPT RSVP"
  ];
}

function ensureScheduleHeaders_(scheduleSheet) {
  const headers = getScheduleHeaders_();
  if (scheduleSheet.getLastRow() === 0) {
    scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const firstRow = scheduleSheet
    .getRange(1, 1, 1, Math.max(scheduleSheet.getLastColumn(), headers.length))
    .getValues()[0];

  const isBlank = firstRow.every(v => String(v || "").trim() === "");
  if (isBlank) {
    scheduleSheet.clear();
    scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const H = headerIndex_(firstRow);
  const missing = headers.filter(h => H[h] == null);
  if (missing.length) {
    const newHeaders = firstRow.slice();
    missing.forEach(h => newHeaders.push(h));
    scheduleSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }
}

function normalizeDateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, "yyyy-MM-dd");
  return String(v).trim();
}

function headerIndex_(row) {
  const H = {};
  row.forEach((h, i) => (H[String(h).trim()] = i));
  return H;
}

function addRole_(map, email, role) {
  if (!email) return;
  if (!map.has(email)) map.set(email, []);
  map.get(email).push(role);
}

function getNextSunday_(d, tz) {
  const base = new Date(Utilities.formatDate(d, tz, "yyyy-MM-dd") + "T00:00:00");
  base.setDate(base.getDate() + ((7 - base.getDay()) % 7));
  return base;
}

function shuffle_(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function getOrCreateCalendar_(name) {
  const c = CalendarApp.getCalendarsByName(name);
  return c.length ? c[0] : CalendarApp.createCalendar(name);
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function loadExistingScheduleByDate_(sh) {
  const out = new Map();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return out;
  const H = headerIndex_(values[0]);
  for (let r = 1; r < values.length; r++) {
    const k = normalizeDateKey_(values[r][H["Date"]]);
    if (k) out.set(k, true);
  }
  return out;
}

function rebuildCountsByMonthAllRolesFromSheet_(sh) {
  const countsByMonth = new Map();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return countsByMonth;

  const H = headerIndex_(data[0]);
  const roles = getRoleColumns_();

  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey) continue;
    const monthKey = dateKey.slice(0, 7);

    if (!countsByMonth.has(monthKey)) countsByMonth.set(monthKey, new Map());
    const monthMap = countsByMonth.get(monthKey);
    if (!monthMap.has("ALL")) monthMap.set("ALL", new Map());

    roles.forEach(role => {
      const email = String(data[r][H[role]] || "").trim();
      if (!email) return;
      const mMap = monthMap.get("ALL");
      mMap.set(email, (mMap.get(email) || 0) + 1);
    });
  }
  return countsByMonth;
}

function rebuildEmailToDatesFromSheet_(sh) {
  const emailToDates = new Map();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return emailToDates;

  const H = headerIndex_(data[0]);
  const roles = getRoleColumns_();

  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey) continue;

    roles.forEach(role => {
      const email = String(data[r][H[role]] || "").trim();
      if (!email) return;
      if (!emailToDates.has(email)) emailToDates.set(email, new Set());
      emailToDates.get(email).add(dateKey);
    });
  }
  return emailToDates;
}

function buildSundayDescription_(dateKey, assignments) {
  return (
    `Church Tech Assignments\n` +
    `Date: ${dateKey}\n\n` +
    `• Audio: ${assignments.Audio || "-"}\n` +
    `• LiveStream: ${assignments.LiveStream || "-"}\n` +
    `• PPT: ${assignments.PPT || "-"}\n\n` +
    `Timezone: ${CONFIG.TIMEZONE}\n` +
    (TEST_EMAIL_ONLY ? `\nTEST MODE: only "${TEST_EMAIL_ONLY}" is invited.\n` : "")
  );
}

// =====================
// SCHEDULE MAP (dateKey -> emails)
// =====================
function buildScheduleMap_(scheduleSheet) {
  const map = new Map();
  const data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) return map;

  const H = headerIndex_(data[0]);
  ["Date", "Audio", "LiveStream", "PPT"].forEach(k => {
    if (H[k] == null) throw new Error(`Schedule missing column: ${k}`);
  });

  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey) continue;

    map.set(dateKey, {
      Audio: String(data[r][H["Audio"]] || "").trim(),
      LiveStream: String(data[r][H["LiveStream"]] || "").trim(),
      PPT: String(data[r][H["PPT"]] || "").trim(),
      AudioRsvp: H["Audio RSVP"] != null ? String(data[r][H["Audio RSVP"]] || "").trim() : "",
      LiveStreamRsvp: H["LiveStream RSVP"] != null ? String(data[r][H["LiveStream RSVP"]] || "").trim() : "",
      PptRsvp: H["PPT RSVP"] != null ? String(data[r][H["PPT RSVP"]] || "").trim() : "",
    });
  }
  return map;
}

function buildEventIdMap_(scheduleSheet) {
  const map = new Map();
  const data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) return map;

  const H = headerIndex_(data[0]);
  if (H["Date"] == null || H["Event Id"] == null) return map;

  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey) continue;
    const eventId = String(data[r][H["Event Id"]] || "").trim();
    if (eventId) map.set(dateKey, eventId);
  }

  return map;
}

// =====================
// MONTH HELPERS
// =====================
function monthNameFromKey_(monthKey) {
  const m = Number(monthKey.split("-")[1]);
  return MONTH_SHEETS[m - 1];
}

function getSundaysForMonthKey_(monthKey) {
  const first = parseMonthKeyToDate_(monthKey);
  const end = endOfMonth_(first);

  const cur = new Date(first);
  while (cur.getDay() !== 0) cur.setDate(cur.getDate() + 1);

  const out = [];
  while (cur <= end) {
    out.push(new Date(cur));
    cur.setDate(cur.getDate() + 7);
  }
  return out;
}

function parseMonthKeyToDate_(monthKey) {
  const [y, m] = monthKey.split("-").map(Number);
  return new Date(`${y}-${String(m).padStart(2, "0")}-01T00:00:00`);
}

function endOfMonth_(dateInMonth) {
  const d = new Date(dateInMonth);
  d.setMonth(d.getMonth() + 1);
  d.setDate(0);
  return d;
}

function isLastDayOfMonth_(d, tz) {
  const key = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const next = new Date(d);
  next.setDate(next.getDate() + 1);
  const keyNext = Utilities.formatDate(next, tz, "yyyy-MM-dd");
  return key.slice(0, 7) !== keyNext.slice(0, 7);
}

// =====================
// ROSTER + PICKING
// =====================
function readRosterFromSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Missing sheet: ${sheetName}`);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error(`Roster sheet "${sheetName}" is empty.`);

  const H = headerIndex_(data[0]);
  if (H["Email"] == null) throw new Error(`Roster must have column: Email`);

  const roster = [];
  for (let r = 1; r < data.length; r++) {
    const email = String(data[r][H["Email"]] || "").trim().toLowerCase();
    if (!email) continue;

    const eligible = {};
    CONFIG.MINISTRIES.forEach(m => {
      const v = data[r][H[m]];
      eligible[m] = isTruthy_(v);
    });

    roster.push({ email, eligible });
  }
  return roster;
}

function normalizePhone_(raw) {
  const digits = String(raw || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("0") && digits.length === 11) return "63" + digits.slice(1);
  if (digits.startsWith("63")) return digits;
  if (digits.length === 10) return "63" + digits;
  return digits;
}

function isTruthy_(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "1" || s === "y";
}

function getSundays_(startDateStr, endDateStr) {
  const start = new Date(startDateStr + "T00:00:00");
  const end = new Date(endDateStr + "T00:00:00");

  while (start.getDay() !== 0) start.setDate(start.getDate() + 1);

  const out = [];
  const cur = new Date(start);
  while (cur <= end) {
    out.push(new Date(cur));
    cur.setDate(cur.getDate() + 7);
  }
  return out;
}

function pickCandidate_({ roster, ministry, countsMap, maxPerMonth, avoidConsecutive, assignedEmailsForDate, emailToDates, dateKey }) {
  const eligible = roster.filter(p => p.eligible[ministry]).map(p => p.email);
  if (!eligible.length) return null;

  let candidates = eligible
    .map(email => ({
      email,
      count: countsMap.get(email) || 0,
      isConsecutive: avoidConsecutive && isConsecutiveAssignment_(email, dateKey, emailToDates),
      isSameDay: assignedEmailsForDate && assignedEmailsForDate.has(email),
    }))
    .filter(x => x.count < maxPerMonth && !x.isSameDay);

  if (!candidates.length) return null;

  if (avoidConsecutive) {
    const nonConsecutive = candidates.filter(x => !x.isConsecutive);
    if (nonConsecutive.length) candidates = nonConsecutive;
  }

  candidates.sort((a, b) => a.count - b.count);

  const min = candidates[0].count;
  const best = candidates.filter(x => x.count === min);
  return best[Math.floor(Math.random() * best.length)].email;
}

// =====================
// SMS (PhilSMS)
// =====================
function getSmsToken_() {
  return PropertiesService.getScriptProperties().getProperty("PHILSMS_TOKEN") || CONFIG.SMS_TOKEN;
}

function sendSms_({ recipient, message }) {
  if (!CONFIG.SMS_ENABLED) return;
  if (!recipient || !message) return;
  if (TEST_EMAIL_ONLY && !TEST_SMS_ONLY) return;

  const token = getSmsToken_();
  if (!token) throw new Error("Missing PHILSMS_TOKEN in Script Properties.");

  const finalRecipient = TEST_SMS_ONLY || recipient;
  if (!finalRecipient) return;

  const url = CONFIG.SMS_API_BASE.replace(/\/$/, "") + "/sms/send";
  const payload = {
    recipient: finalRecipient,
    sender_id: CONFIG.SMS_SENDER_ID,
    type: CONFIG.SMS_TYPE,
    message,
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() >= 400) {
    Logger.log("SMS send failed: " + res.getContentText());
  }
}

function buildSmsSaturdayReminder_({ prettyDate, roles }) {
  const roleText = Array.isArray(roles) ? roles.join(", ") : String(roles || "");
  return `Tech Reminder: We appreciate your support. This Sunday (${prettyDate}) role: ${roleText}. - ${CONFIG.TECH_TEAM_NAME}`;
}

function buildSmsTechSubInvite_({ personAName, role, prettyDate }) {
  return `Tech Sub Invite: We appreciate your support. ${personAName} asked you to cover ${role} on ${prettyDate}. Check email to respond. - ${CONFIG.TECH_TEAM_NAME}`;
}

function buildSmsTechSubConfirmation_({ personBName, role, prettyDate }) {
  return `Tech Sub Confirmed: We appreciate your support. ${personBName} will cover ${role} on ${prettyDate}. - ${CONFIG.TECH_TEAM_NAME}`;
}

function buildSmsTechSubNoNotify_({ role, dateKey, email }) {
  return `Tech Sub Declined: ${email} declined ${role} on ${dateKey}. - ${CONFIG.TECH_TEAM_NAME}`;
}

function buildSmsMonthlyReminder_({ monthLabel, count }) {
  const n = Number(count || 0);
  return (
    `Good day!\n` +
    `We're pleased to inform you that your ${monthLabel} tech schedule is available.\n` +
    `You have ${n} assigned duty this month.\n\n` +
    `Kindly check your email for the details.\n` +
    `Thank you for your continued service.\n` +
    `- ${CONFIG.TECH_TEAM_NAME}`
  );
}

// =====================
// PRETTY EMAILS
// =====================
function escapeHtml_(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function pill_(text) {
  return `
    <span style="
      display:inline-block;
      padding:6px 10px;
      border-radius:999px;
      background:#f1f5f9;
      border:1px solid #e2e8f0;
      color:#0f172a;
      font-size:12px;
      line-height:1;
      margin:6px 8px 0 0;
      white-space:nowrap;
    ">${escapeHtml_(text)}</span>
  `;
}

function buildPrettyEmail_({ title, subtitle, bodyHtml }) {
  const safeTitle = escapeHtml_(title);
  const safeSubtitle = escapeHtml_(subtitle || "");

  const brandBlock = CONFIG.SHOW_LOGO
    ? `
      <img src="${CONFIG.LOGO_URL}" alt="${escapeHtml_(CONFIG.CHURCH_NAME)}"
        style="height:36px;width:auto;display:block;margin:0 0 10px 0;" />
      <div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#64748b;">
        ${escapeHtml_(CONFIG.CHURCH_NAME)}
      </div>
    `
    : `
      <div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#64748b;">
        ${escapeHtml_(CONFIG.CHURCH_NAME)}
      </div>
    `;

  return `
  <div style="margin:0;padding:0;background:#f6f7f9;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f7f9;">
      <tr>
        <td align="center" style="padding:14px 10px;">
          <table role="presentation" width="520" cellpadding="0" cellspacing="0"
            style="width:100%;max-width:520px;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e5e7eb;">

            <tr>
              <td style="padding:14px 16px 8px 16px;text-align:left;">
                ${brandBlock}
              </td>
            </tr>

            <tr>
              <td style="padding:0 16px 8px 16px;">
                <div style="
                  font-family:Arial,Helvetica,sans-serif;
                  font-size:22px;
                  line-height:1.25;
                  color:#0f172a;
                  font-weight:800;
                  margin:0;
                ">
                  ${safeTitle}
                </div>

                ${subtitle ? `
                  <div style="
                    font-family:Arial,Helvetica,sans-serif;
                    font-size:14px;
                    line-height:1.4;
                    color:#64748b;
                    margin-top:8px;
                  ">
                    ${safeSubtitle}
                  </div>
                ` : ``}
              </td>
            </tr>

            <tr>
              <td style="padding:10px 16px 6px 16px;">
                <div style="
                  font-family:Arial,Helvetica,sans-serif;
                  font-size:15px;
                  line-height:1.65;
                  color:#111827;
                ">
                  ${bodyHtml}
                </div>
              </td>
            </tr>

            <tr>
              <td style="padding:10px 16px 14px 16px;">
                <div style="border-top:1px solid #e5e7eb;margin-top:8px;padding-top:10px;
                  font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#94a3b8;">
                  - ${escapeHtml_(CONFIG.TECH_TEAM_NAME)}
                </div>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </div>
  `;
}

function buildPrettyReminderEmail_({ prettyDate, roles, displayName }) {
  const roleText = Array.isArray(roles) ? roles.join(", ") : String(roles || "");

  return buildPrettyEmail_({
    title: "Tech Schedule Reminder",
    subtitle: "",
    bodyHtml: `
      <div style="margin:0 0 12px 0;">Hi <b>${escapeHtml_(displayName || "there")}</b>!</div>

      <div style="margin:0 0 12px 0;">You are scheduled to serve tomorrow,</div>

      <div style="font-size:26px;font-weight:800;margin:6px 0 12px 0;">
        ${escapeHtml_(prettyDate)}
      </div>

      <div style="margin:0 0 12px 0;">
        Your role is: <b>${escapeHtml_(roleText)}</b>
      </div>

      <div style="margin:0 0 12px 0;color:#334155;">
        If you have conflicts please notify us Tech group chat as soon as possible.
      </div>

      <div style="margin:0;">Thank you for serving in our church’s tech ministry.</div>
    `,
  });
}

function sendCalendarInviteEmail_({ toEmail, subject, htmlBody, event, summaryOverride }) {
  const ics = buildIcsInvite_(event, toEmail, summaryOverride);
  GmailApp.sendEmail(toEmail, subject, "Please view this email in HTML.", {
    htmlBody,
    attachments: [
      {
        fileName: "invite.ics",
        mimeType: "text/calendar; charset=UTF-8; method=REQUEST",
        content: ics,
      },
    ],
  });
}

function buildIcsInvite_(event, attendeeEmail, summaryOverride) {
  const now = new Date();
  const uid = event.getId();
  const summary = escapeIcsText_(summaryOverride || event.getTitle() || "Tech Duty");
  const description = escapeIcsText_(event.getDescription() || "");
  const location = escapeIcsText_(event.getLocation() || "");
  const attendee = escapeIcsText_(attendeeEmail);

  const dtStamp = formatIcsDateTimeUtc_(now);

  let dtStart = "";
  let dtEnd = "";
  if (event.isAllDayEvent()) {
    dtStart = `DTSTART;VALUE=DATE:${formatIcsDate_(event.getStartTime())}`;
    dtEnd = `DTEND;VALUE=DATE:${formatIcsDate_(event.getEndTime())}`;
  } else {
    dtStart = `DTSTART:${formatIcsDateTimeUtc_(event.getStartTime())}`;
    dtEnd = `DTEND:${formatIcsDateTimeUtc_(event.getEndTime())}`;
  }

  const organizerEmail = Session.getActiveUser().getEmail() || "";
  const organizerLine = organizerEmail
    ? `ORGANIZER;CN=${escapeIcsText_(CONFIG.TECH_TEAM_NAME)}:MAILTO:${escapeIcsText_(organizerEmail)}`
    : "";

  return [
    "BEGIN:VCALENDAR",
    "PRODID:-//Church Tech Scheduler//EN",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${escapeIcsText_(uid)}`,
    `DTSTAMP:${dtStamp}`,
    dtStart,
    dtEnd,
    `SUMMARY:${summary}`,
    description ? `DESCRIPTION:${description}` : "",
    location ? `LOCATION:${location}` : "",
    organizerLine,
    `ATTENDEE;CN=${attendee};ROLE=REQ-PARTICIPANT;RSVP=TRUE:MAILTO:${attendee}`,
    "END:VEVENT",
    "END:VCALENDAR",
  ].filter(Boolean).join("\r\n");
}

function ensureGuestOnEvent_(event, email) {
  if (!email) return;
  const exists = event.getGuestList().some(g => String(g.getEmail() || "").trim().toLowerCase() === String(email).trim().toLowerCase());
  if (!exists) {
    try {
      event.addGuest(email);
    } catch (_) {}
  }
}

function formatIcsDate_(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyyMMdd");
}

function formatIcsDateTimeUtc_(d) {
  return Utilities.formatDate(d, "UTC", "yyyyMMdd'T'HHmmss'Z'");
}

function escapeIcsText_(s) {
  return String(s || "")
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

// =====================
// NEW: Tech Sub Emails (Confirmation + Invite with YES/NO buttons)
// =====================
function sendChangeEmailsSmart_(ss, dateKey, changes, eventId) {
  const roster = buildRosterMaps_(ss);
  const prettyDate = Utilities.formatDate(new Date(dateKey + "T00:00:00"), CONFIG.TIMEZONE, "EEE, MMM d, yyyy");

  // If it looks like a substitution (old + new), send the Tech Sub emails.
  // Otherwise, fallback to generic "Schedule Update".
  const subs = changes.filter(c => c.oldEmail && c.newEmail && c.oldEmail !== c.newEmail);

  subs.forEach(c => {
    const personAEmail = String(c.oldEmail || "").trim().toLowerCase();
    const personBEmail = String(c.newEmail || "").trim().toLowerCase();

    const personAName = roster.emailToName.get(personAEmail) || personAEmail;
    const personBName = roster.emailToName.get(personBEmail) || personBEmail;

    // Mark role RSVP as Invited when a sub is requested
    setRoleRsvpInvited_(ss, dateKey, c.role);

    // Defer Person A confirmation until Person B clicks YES.
    savePendingSubConfirmation_({
      dateKey,
      role: c.role,
      oldEmail: personAEmail,
      newEmail: personBEmail,
      createdAt: new Date().toISOString(),
    });

    // (2) Person B invite with YES/NO buttons
    const inviteHtml = buildTechSubInviteEmail_({
      personAName,
      personBName,
      personBEmail,
      role: c.role,
      prettyDate,
      dateKey,
      eventId,
    });
    GmailApp.sendEmail(TEST_EMAIL_ONLY || personBEmail, "Tech Sub Invite", "Please view in HTML.", { htmlBody: inviteHtml });
    const personBMobile = roster.emailToMobile.get(personBEmail) || "";
    if (personBMobile) {
      const sms = buildSmsTechSubInvite_({ personAName, role: c.role, prettyDate });
      sendSms_({ recipient: personBMobile, message: sms });
    }
  });

  // For non-sub changes, generic schedule update to affected emails
  const nonSubs = changes.filter(c => !(c.oldEmail && c.newEmail && c.oldEmail !== c.newEmail));
  if (nonSubs.length) sendPrettyChangeEmail_(dateKey, nonSubs);
}

function setRoleRsvpInvited_(ss, dateKey, role) {
  const schedule = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedule) return;
  ensureScheduleHeaders_(schedule);

  const row = findOrCreateScheduleRowByDate_(schedule, dateKey);
  const header = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);

  const rsvpColName =
    role === "Audio" ? "Audio RSVP" :
    role === "LiveStream" ? "LiveStream RSVP" :
    role === "PPT" ? "PPT RSVP" :
    "";

  if (!rsvpColName || H[rsvpColName] == null) return;

  schedule.getRange(row, H[rsvpColName] + 1).setValue("Invited");

  updateMonthlyRsvpForDate_(ss, dateKey, {
    Audio: role === "Audio" ? "Invited" : undefined,
    LiveStream: role === "LiveStream" ? "Invited" : undefined,
    PPT: role === "PPT" ? "Invited" : undefined,
  });
}

function buildTechSubConfirmationEmail_({ personAName, personBName, role, prettyDate }) {
  return buildPrettyEmail_({
    title: "Tech Sub Confirmation",
    subtitle: prettyDate,
    bodyHtml: `
      <div style="margin:0 0 12px 0;">Hi <b>${escapeHtml_(personAName || "there")}!</b></div>

      <div style="margin:0 0 14px 0;">
        This is to confirm that <b>${escapeHtml_(personBName || "someone")}</b> will cover for your volunteer duty as
        <b>${escapeHtml_(role)}</b> on <b>${escapeHtml_(prettyDate)}</b>.
      </div>

      <div style="margin:0 0 14px 0;color:#334155;">
        We’re looking forward to your next volunteer schedule!
      </div>

      <div style="margin:0;">Thank you for serving in our church’s tech ministry.</div>
    `,
  });
}

function buildTechSubInviteEmail_({ personAName, personBName, personBEmail, role, prettyDate, dateKey, eventId }) {
  const eid = eventId || getEventIdByDateKey_(dateKey);

  const yesLink = buildWebAppRsvpLink_({
    action: "YES",
    role,
    dateKey,
    assigneeEmail: personBEmail,
    eventId: eid,
  });

  const noLink = buildWebAppRsvpLink_({
    action: "NO",
    role,
    dateKey,
    assigneeEmail: personBEmail,
    eventId: eid,
  });

  return buildPrettyEmail_({
    title: "Tech Sub Invite",
    subtitle: prettyDate,
    bodyHtml: `
      <div style="margin:0 0 12px 0;">
        Hi <b>${escapeHtml_(personBName || "there")}!</b>
      </div>

      <div style="margin:0 0 14px 0;">
        <b>${escapeHtml_(personAName || "Someone")}</b> has tagged you to cover for his volunteer duty as
        <b>${escapeHtml_(role)}</b> on <b>${escapeHtml_(prettyDate)}</b>.
      </div>

      <div style="font-weight:800;margin:16px 0 10px 0;">Are you available?</div>

      <div style="display:flex;gap:10px;margin:0 0 14px 0;">
        <a href="${yesLink}"
          style="display:inline-block;background:#111827;color:#ffffff;text-decoration:none;
            padding:12px 16px;border-radius:12px;font-weight:800;">
          Yes
        </a>

        <a href="${noLink}"
          style="display:inline-block;background:#ffffff;color:#111827;text-decoration:none;
            padding:12px 16px;border-radius:12px;font-weight:800;border:1px solid #e5e7eb;">
          No
        </a>
      </div>

      <div style="color:#334155;margin:0 0 14px 0;">
        If you have conflicts please click <b>No</b> and notify us through Tech Group chat as soon as possible.
      </div>

      <div style="margin:0;">
        Thank you for serving in our church’s tech ministry
      </div>
    `,
  });
}

function getEventIdByDateKey_(dateKey) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const sh = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sh) return "";
  ensureScheduleHeaders_(sh);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return "";
  const H = headerIndex_(data[0]);

  for (let r = 1; r < data.length; r++) {
    const k = normalizeDateKey_(data[r][H["Date"]]);
    if (k === dateKey) return String(data[r][H["Event Id"]] || "").trim();
  }
  return "";
}

function sendPrettyChangeEmail_(dateKey, changes) {
  const items = changes.map(c => {
    const oldE = c.oldEmail ? escapeHtml_(c.oldEmail) : "-";
    const newE = c.newEmail ? escapeHtml_(c.newEmail) : "-";
    return `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;font-weight:700;">${escapeHtml_(c.role)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#475569;">${oldE}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#0f172a;">${newE}</td>
      </tr>
    `;
  }).join("");

  const html = buildPrettyEmail_({
    title: "Schedule Update",
    subtitle: dateKey,
    bodyHtml: `
      <div style="margin:0 0 10px 0;">The schedule has been updated:</div>

      <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
        style="border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;border-collapse:separate;">
        <tr style="background:#f8fafc;">
          <th align="left" style="padding:10px 12px;border-bottom:1px solid #e5e7eb;">Role</th>
          <th align="left" style="padding:10px 12px;border-bottom:1px solid #e5e7eb;">From</th>
          <th align="left" style="padding:10px 12px;border-bottom:1px solid #e5e7eb;">To</th>
        </tr>
        ${items}
      </table>

      <div style="margin:12px 0 0 0;color:#475569;">
        If you have a conflict, please reply as soon as possible.
      </div>
    `,
  });

  const emails = new Set();
  changes.forEach(c => {
    if (c.oldEmail) emails.add(String(c.oldEmail).trim().toLowerCase());
    if (c.newEmail) emails.add(String(c.newEmail).trim().toLowerCase());
  });

  const subject = `Schedule Update | ${CONFIG.TECH_TEAM_NAME}`;
  emails.forEach(e => {
    GmailApp.sendEmail(TEST_EMAIL_ONLY || e, subject, "Please view this email in HTML.", { htmlBody: html });
  });
}

// =====================
// VALIDATION HELPERS
// =====================
function getRoleColumns_() {
  return ["Audio", "LiveStream", "PPT"];
}

function dateKeyAddDays_(dateKey, days) {
  const d = new Date(dateKey + "T00:00:00");
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd");
}

function isEmailAssignedOnDate_(scheduleSheet, dateKey, email) {
  if (!email) return false;
  const data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) return false;
  const H = headerIndex_(data[0]);
  const roles = getRoleColumns_();
  const target = String(email).trim().toLowerCase();

  for (let r = 1; r < data.length; r++) {
    const k = normalizeDateKey_(data[r][H["Date"]]);
    if (k !== dateKey) continue;
    return roles.some(role => String(data[r][H[role]] || "").trim().toLowerCase() === target);
  }
  return false;
}

function countEmailAssignmentsInMonth_(scheduleSheet, email, monthKey) {
  if (!email) return 0;
  const data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  const H = headerIndex_(data[0]);
  const roles = getRoleColumns_();
  const target = String(email).trim().toLowerCase();

  let count = 0;
  for (let r = 1; r < data.length; r++) {
    const dateKey = normalizeDateKey_(data[r][H["Date"]]);
    if (!dateKey || !dateKey.startsWith(monthKey)) continue;
    roles.forEach(role => {
      if (String(data[r][H[role]] || "").trim().toLowerCase() === target) count++;
    });
  }
  return count;
}

function isConsecutiveAssignment_(email, dateKey, emailToDates) {
  if (!email) return false;
  const prevKey = dateKeyAddDays_(dateKey, -7);
  const nextKey = dateKeyAddDays_(dateKey, 7);
  const dates = emailToDates && emailToDates.get(String(email).trim().toLowerCase());
  if (!dates) return false;
  return dates.has(prevKey) || dates.has(nextKey);
}

function validateAssignmentForScheduleRow_(scheduleSheet, row, roleLabel, email, options) {
  if (!email) return null;
  const opts = options || {};

  const header = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);
  const dateKey = normalizeDateKey_(scheduleSheet.getRange(row, H["Date"] + 1).getValue());
  if (!dateKey) return "Missing Date for this row.";

  const roles = getRoleColumns_();
  const target = String(email).trim().toLowerCase();
  const currentRoleEmail = String(scheduleSheet.getRange(row, H[roleLabel] + 1).getValue() || "").trim().toLowerCase();
  if (currentRoleEmail === target) return null;

  for (const role of roles) {
    if (role === roleLabel) continue;
    const otherEmail = String(scheduleSheet.getRange(row, H[role] + 1).getValue() || "").trim().toLowerCase();
    if (otherEmail === target) {
      return "That person already has another role on this Sunday.";
    }
  }

  if (!opts.ignoreMaxPerMonth) {
    const monthKey = dateKey.slice(0, 7);
    const currentCount = countEmailAssignmentsInMonth_(scheduleSheet, target, monthKey);
    if (currentCount >= CONFIG.MAX_ASSIGNMENTS_PER_PERSON_PER_MONTH_PER_MINISTRY) {
      return `That person already has ${CONFIG.MAX_ASSIGNMENTS_PER_PERSON_PER_MONTH_PER_MINISTRY} assignment(s) this month.`;
    }
  }

  if (opts.enforceConsecutive) {
    const prevKey = dateKeyAddDays_(dateKey, -7);
    const nextKey = dateKeyAddDays_(dateKey, 7);
    if (isEmailAssignedOnDate_(scheduleSheet, prevKey, target) || isEmailAssignedOnDate_(scheduleSheet, nextKey, target)) {
      return "That person is already assigned on a consecutive Sunday.";
    }
  }

  return null;
}

function getConsecutiveWarning_(scheduleSheet, row, email) {
  if (!email) return null;
  const header = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];
  const H = headerIndex_(header);
  const dateKey = normalizeDateKey_(scheduleSheet.getRange(row, H["Date"] + 1).getValue());
  if (!dateKey) return null;

  const target = String(email).trim().toLowerCase();
  const prevKey = dateKeyAddDays_(dateKey, -7);
  const nextKey = dateKeyAddDays_(dateKey, 7);
  if (isEmailAssignedOnDate_(scheduleSheet, prevKey, target) || isEmailAssignedOnDate_(scheduleSheet, nextKey, target)) {
    return "Warning: That person is also assigned on a consecutive Sunday.";
  }
  return null;
}
