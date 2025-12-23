// =====================
// Code.gs
// =====================
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
  { type: "doxology",            title: "Doxology",                                   order: 160,  posture: "STANDING" },
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
  tpl.boot = getBootData_(); // IMPORTANT: HTML expects `boot`
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

  return { sh, headers, col };
}

/**
 * Returns the saved section for (service_id, type) or null.
 * If duplicates exist, returns the last one (most recent row).
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
 * ✅ Ensures ALL sections for a service_id exist based on flow.
 * - Creates missing rows with default order/title/posture
 * - body is always blank for skeleton rows
 * - Removes duplicates for the whole service (keeps last)
 */
function ensureServiceSkeleton_(service_id, flow) {
  const types = (FLOWS[flow] || FLOWS.AM || []).slice();
  if (!types.length) throw new Error(`Flow is empty/unknown: ${flow}`);

  const metaByType = {};
  TYPE_META.forEach(m => metaByType[m.type] = m);

  const { sh, headers, col } = getSectionsSheet_();
  const values = sh.getDataRange().getValues();

  const sidIdx = col["service_id"];
  const typeIdx = col["type"];

  // Map: type -> list of row indices (1-based)
  const map = {};
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][sidIdx]).trim() !== service_id) continue;
    const t = String(values[i][typeIdx]).trim();
    if (!t) continue;
    if (!map[t]) map[t] = [];
    map[t].push(i + 1);
  }

  // Dedup per type for this service_id (keep last row)
  let dedupDeleted = 0;
  const rowsToDelete = [];
  Object.keys(map).forEach(t => {
    const rows = map[t];
    if (rows.length > 1) {
      rows.slice(0, -1).forEach(r => rowsToDelete.push(r));
    }
  });
  rowsToDelete.sort((a, b) => b - a).forEach(r => {
    sh.deleteRow(r);
    dedupDeleted++;
  });

  // Rebuild existence map after deletion
  const fresh = sh.getDataRange().getValues();
  const exists = {};
  for (let i = 1; i < fresh.length; i++) {
    if (String(fresh[i][sidIdx]).trim() !== service_id) continue;
    const t = String(fresh[i][typeIdx]).trim();
    exists[t] = true;
  }

  // Create missing rows
  let createdCount = 0;
  const rowsToAppend = [];

  types.forEach(t => {
    if (exists[t]) return;

    const meta = metaByType[t] || { title: t, order: 0, posture: "" };
    const row = new Array(headers.length).fill("");
    row[col["service_id"]] = service_id;
    row[col["order"]] = Number(meta.order || 0);
    row[col["type"]] = t;
    row[col["title"]] = String(meta.title || "");
    row[col["body"]] = ""; // skeleton body blank
    row[col["posture"]] = String(meta.posture || "");
    row[col["updated_at"]] = new Date();

    rowsToAppend.push(row);
    createdCount++;
  });

  if (rowsToAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  return { createdCount, dedupDeleted };
}

/**
 * Upsert by unique key (service_id + type).
 * ✅ On first save for a service, auto-generates the whole service skeleton (all sections in the flow) with blank bodies.
 * ✅ Deduplicates any existing duplicates (keeps one, deletes the rest).
 */
function upsertSection(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const service_id = String(payload.service_id || "").trim();
    const flow = String(payload.flow || "AM").trim().toUpperCase(); // from UI
    const type = String(payload.type || "").trim();
    const title = String(payload.title || "").trim();
    const body = String(payload.body || "");
    const posture = String(payload.posture || "").trim().toUpperCase();

    const orderRaw = payload.order;
    const order = Number(orderRaw);

    if (!service_id) throw new Error("service_id is required.");
    if (!type) throw new Error("type is required.");
    if (!orderRaw && orderRaw !== 0) throw new Error("order is required.");
    if (!order || Number.isNaN(order)) throw new Error("order must be a number.");

    // ✅ Ensure skeleton exists (all steps, blank bodies)
    const skeletonInfo = ensureServiceSkeleton_(service_id, flow);

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
    row[col["body"]] = body; // only this step gets its body
    row[col["posture"]] = posture;
    row[col["updated_at"]] = new Date();

    let mode = "created";
    let dedupDeleted = 0;

    if (matches.length) {
      // keep the LAST match as canonical row (most recent)
      const keptRowIndex = matches[matches.length - 1];
      sh.getRange(keptRowIndex, 1, 1, row.length).setValues([row]);
      mode = "updated";

      // delete other duplicates (bottom to top)
      const toDelete = matches.slice(0, -1).sort((a, b) => b - a);
      toDelete.forEach(r => sh.deleteRow(r));
      dedupDeleted = toDelete.length;
    } else {
      sh.appendRow(row);
      mode = "created";
    }

    return {
      ok: true,
      mode,
      dedup_deleted: dedupDeleted,
      skeleton_created: skeletonInfo.createdCount,
      skeleton_dedup_deleted: skeletonInfo.dedupDeleted,
    };

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const path = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : ""; 
  // use:
  // /exec?p=liturgy&service_id=2025-12-21-AM
  // /exec?p=today
  // /exec?p=admin  (your wizard)

  if (path === "admin") {
    const tpl = HtmlService.createTemplateFromFile("AdminWizard");
    tpl.boot = getBootData_();
    return tpl.evaluate()
      .setTitle("Worship Order — Admin")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (path === "liturgy") {
    const serviceId = resolveServiceId_(e);
    const tpl = HtmlService.createTemplateFromFile("PublicLiturgy");
    tpl.boot = buildPublicBoot_(serviceId);
    return tpl.evaluate()
      .setTitle("Today's Liturgy")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // default: today landing
  const serviceId = resolveServiceId_(e);
  const tpl = HtmlService.createTemplateFromFile("PublicToday");
  tpl.boot = buildPublicBoot_(serviceId);
  return tpl.evaluate()
    .setTitle("Today's Liturgy")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- PUBLIC BOOT (cached) ---
function buildPublicBoot_(service_id) {
  const data = getServiceData_(service_id);
  return {
    service_id: service_id,
    heading: data.heading || service_id,
    dateLabel: data.dateLabel || "",
    sections: data.sections || [],
  };
}

/**
 * Resolve service_id:
 * - if service_id param exists, use it
 * - else default to "today" = YYYY-MM-DD-AM (customize)
 */
function resolveServiceId_(e) {
  const q = (e && e.parameter) ? e.parameter : {};
  const sid = String(q.service_id || "").trim();
  if (sid) return sid;

  // default: today's AM (customize as needed)
  const tz = "Asia/Manila";
  const d = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  return `${d}-AM`;
}

/**
 * Returns { heading, dateLabel, sections[] } for a service_id.
 * Cached to avoid hammering Sheets with 100 users.
 */
function getServiceData_(service_id) {
  const cache = CacheService.getScriptCache();
  const key = "svc:" + service_id;
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const { sh, col } = getSectionsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return { heading: service_id, sections: [] };

  const sidIdx = col["service_id"];
  const orderIdx = col["order"];
  const typeIdx = col["type"];
  const titleIdx = col["title"];
  const bodyIdx = col["body"];
  const postureIdx = col["posture"];

  const sections = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][sidIdx]).trim() !== service_id) continue;
    sections.push({
      order: Number(values[i][orderIdx] || 0),
      type: String(values[i][typeIdx] || ""),
      title: String(values[i][titleIdx] || ""),
      body: String(values[i][bodyIdx] || ""),
      posture: String(values[i][postureIdx] || ""),
    });
  }

  sections.sort((a,b) => (a.order||0) - (b.order||0));

  const out = {
    heading: service_id.replace(/-AM$|-PM$/,"").toUpperCase(),
    dateLabel: service_id,
    sections,
  };

  cache.put(key, JSON.stringify(out), 60 * 60); // 1 hour
  return out;
}

