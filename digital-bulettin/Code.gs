const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  SECTIONS_SHEET: "Sections",
};

// Master list (your defaults)
const TYPE_META = [
  { type: "prelude",             title: "Prelude",                                    order: 10,  posture: "" },
  { type: "call",                title: "Call to Worship",                            order: 20,  posture: "STANDING" },
  { type: "first_hymn",          title: "First Hymn of Praise",                       order: 30,  posture: "STANDING" },
  { type: "opening_prayer",      title: "Opening Invocation / Prayer",                order: 40,  posture: "STANDING" },
  { type: "reading",             title: "Reading of the Word",                        order: 50,  posture: "SEATED" },
  { type: "decalogue",           title: "Corporate Reading of the Decalogue",         order: 60,  posture: "SEATED" },
  { type: "confession_prayer",   title: "Corporate Confession and Prayer for Pardon", order: 70,  posture: "SEATED" },
  { type: "second_hymn",         title: "Second Hymn of Praise",                      order: 80,  posture: "STANDING" },
  { type: "creed",               title: "Apostle's Creed",                            order: 90,  posture: "SEATED" },
  { type: "pastoral_prayer",     title: "Pastoral Prayer",                            order: 100, posture: "SEATED" },
  { type: "preparatory_hymn",    title: "Preparatory Hymn",                           order: 110, posture: "STANDING" },
  { type: "prayer_illumination", title: "Prayer for Illumination",                    order: 120, posture: "SEATED" },
  { type: "preaching",           title: "Preaching of the Word",                      order: 130, posture: "SEATED" },
  { type: "response_hymn",       title: "Response Hymn",                              order: 140, posture: "SEATED" },
  { type: "offertory_prayer",    title: "Offertory & Closing Prayer",                 order: 150, posture: "SEATED" },
  { type: "doxology",            title: "Doxology",                                   order: 160, posture: "STANDING" },
  { type: "benediction",         title: "Benediction",                                order: 170, posture: "STANDING" },
  { type: "postlude",            title: "Postlude / Silence for Reflection",          order: 180, posture: "SEATED" },
];

// 🔥 Define flows here (AM vs PM)
const FLOWS = {
  AM: [
    "prelude","call","first_hymn","opening_prayer","reading","decalogue",
    "confession_prayer","second_hymn","creed","pastoral_prayer",
    "preparatory_hymn","prayer_illumination","preaching","response_hymn",
    "offertory_prayer","doxology","benediction","postlude"
  ],
  PM: [
    "prelude","call","reading","opening_prayer",
    "preaching","response_hymn","benediction","postlude"
  ],
};

function doGet() {
  const tpl = HtmlService.createTemplateFromFile("AdminWizard");
  tpl.boot = getBootData_(); // <-- must be named `boot` to match the HTML
  return tpl.evaluate()
    .setTitle("Worship Order — Wizard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBootData_() {
  return {
    typeMeta: TYPE_META,
    flows: FLOWS,
    postureOptions: ["", "STANDING", "SEATED"],
  };
}

/**
 * Ensures Sections sheet exists and has required headers.
 * Returns { sh, headers, col } where col maps header->index
 */
function getSectionsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SECTIONS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SECTIONS_SHEET);
  }

  const requiredHeaders = ["service_id", "order", "type", "title", "body", "posture", "updated_at"];

  const range = sh.getDataRange();
  const values = range.getNumRows() ? range.getValues() : [];
  let headers = values.length ? values[0].map(String) : [];

  // If empty sheet or missing headers, initialize headers.
  const hasAnyHeader = headers.some(h => String(h).trim() !== "");
  if (!hasAnyHeader) {
    sh.clear();
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    headers = requiredHeaders.slice();
  } else {
    // Ensure required headers exist (append missing)
    const existing = new Set(headers.map(h => String(h).trim()));
    const missing = requiredHeaders.filter(h => !existing.has(h));
    if (missing.length) {
      sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
    }
  }

  const col = {};
  headers.forEach((h, i) => col[String(h).trim()] = i);

  return { ss, sh, headers, col };
}

/**
 * Returns the saved section for (service_id, type) or null.
 * If duplicates exist, returns the last one (most recent row) by default.
 */
function getSection(service_id, type) {
  service_id = String(service_id || "").trim();
  type = String(type || "").trim();
  if (!service_id) throw new Error("service_id is required.");
  if (!type) throw new Error("type is required.");

  const { sh, headers, col } = getSectionsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return null;

  const sidIdx = col["service_id"];
  const typeIdx = col["type"];

  // choose LAST matching row (most recent append/update)
  let foundRow = null;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][sidIdx]).trim() === service_id &&
        String(values[i][typeIdx]).trim() === type) {
      foundRow = values[i];
      break;
    }
  }
  if (!foundRow) return null;

  return {
    service_id: String(foundRow[col["service_id"]] || "").trim(),
    order: foundRow[col["order"]],
    type: String(foundRow[col["type"]] || "").trim(),
    title: String(foundRow[col["title"]] || ""),
    body: String(foundRow[col["body"]] || ""),
    posture: String(foundRow[col["posture"]] || "").trim(),
  };
}

/**
 * Upsert by unique key (service_id + type).
 * Also deduplicates any existing duplicates (keeps ONE row, deletes the rest).
 */
function upsertSection(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const service_id = String(payload.service_id || "").trim();
    const type = String(payload.type || "").trim();
    const title = String(payload.title || "").trim();
    const body = String(payload.body || "");
    const posture = String(payload.posture || "").trim().toUpperCase();

    // order can come as string/number; allow blank -> fallback to 0 check later
    const orderRaw = payload.order;
    const order = Number(orderRaw);

    if (!service_id) throw new Error("service_id is required.");
    if (!type) throw new Error("type is required.");
    if (!orderRaw && orderRaw !== 0) throw new Error("order is required.");
    if (!order || Number.isNaN(order)) throw new Error("order must be a number.");

    const { sh, headers, col } = getSectionsSheet_();
    const values = sh.getDataRange().getValues();

    const sidIdx = col["service_id"];
    const typeIdx = col["type"];

    // find ALL matching rows (may include duplicates)
    const matches = [];
    for (let i = 1; i < values.length; i++) {
      if (
        String(values[i][sidIdx]).trim() === service_id &&
        String(values[i][typeIdx]).trim() === type
      ) {
        matches.push(i + 1); // sheet row index (1-based)
      }
    }

    const row = new Array(headers.length).fill("");
    row[col["service_id"]] = service_id;
    row[col["order"]] = order;
    row[col["type"]] = type;
    row[col["title"]] = title;
    row[col["body"]] = body;
    row[col["posture"]] = posture;
    row[col["updated_at"]] = new Date();

    let mode = "created";
    let keptRowIndex = -1;

    if (matches.length) {
      // keep the LAST match as the canonical row (most recent), overwrite it
      keptRowIndex = matches[matches.length - 1];
      sh.getRange(keptRowIndex, 1, 1, row.length).setValues([row]);
      mode = "updated";

      // delete other duplicates (from bottom to top to keep indexes valid)
      const toDelete = matches.slice(0, -1).sort((a, b) => b - a);
      toDelete.forEach(r => sh.deleteRow(r));

      return { ok: true, mode, dedup_deleted: toDelete.length };
    }

    // no existing -> append
    sh.appendRow(row);
    return { ok: true, mode, dedup_deleted: 0 };

  } finally {
    lock.releaseLock();
  }
}
