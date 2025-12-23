// =====================
// Code.gs (Web App Router + Public Backend + Admin Wizard)
// =====================
const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  SECTIONS_SHEET: "Sections",
  SERVICES_SHEET: "Services",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Bulletin")
    .addItem("Clear Public Cache", "clearPublicCache")
    .addItem("Generate QR Codes", "generateQrCodes")
    .addToUi();
}

function clearPublicCache() {
  const { sh, col } = getSectionsSheet_();
  const values = sh.getDataRange().getValues();
  const sidIdx = col["service_id"];

  const keys = new Set();
  for (let i = 1; i < values.length; i++) {
    const sid = String(values[i][sidIdx] || "").trim();
    if (sid) keys.add("svc:" + sid);
  }

  if (keys.size) {
    CacheService.getScriptCache().removeAll(Array.from(keys));
  }

  SpreadsheetApp.getUi().alert(
    "Cache cleared",
    `Cleared ${keys.size} service cache key(s).`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function generateQrCodes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SERVICES_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert(`Sheet not found: ${CONFIG.SERVICES_SHEET}`);
    return;
  }

  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }

  const headers = values[0].map(h => String(h || "").trim().toLowerCase());
  const sidIdx = headers.indexOf("service_id");
  const qrIdx = headers.indexOf("qr_code");
  if (sidIdx < 0 || qrIdx < 0) {
    SpreadsheetApp.getUi().alert("Missing required headers: service_id and/or qr_code.");
    return;
  }

  const startRow = 2;
  const numRows = values.length - 1;
  const qrRange = sh.getRange(startRow, qrIdx + 1, numRows, 1);
  const qrValues = qrRange.getValues();
  const formulaBase = "https://script.google.com/macros/s/AKfycbxBj2B8nCccx8wWuFmODDbjoMicggRFly_lBC8F1n0iz45ZWwSCXN4UglUvpa1NqJ22/exec?p=liturgy&service_id=";

  const out = [];
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const existing = String(qrValues[i][0] || "").trim();
    const sid = String(values[i + 1][sidIdx] || "").trim();
    if (existing || !sid) {
      out.push([existing]);
      continue;
    }

    const formula = `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" & ENCODEURL("${formulaBase}" & A${row}))`;
    out.push([formula]);
  }

  qrRange.setFormulas(out);
  SpreadsheetApp.getUi().alert("QR codes generated for empty rows.");
}

// Master list (defaults used by Admin + skeleton)
const TYPE_META = [
  { type: "prelude",             title: "Prelude Liturgical Silence For Preparation", order: 10,  posture: "SEATED" },
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
  { type: "response_hymn",       title: "Response Hymn",                              order: 140, posture: "STANDING" },
  { type: "offertory_prayer",    title: "Offertory & Closing Prayer",                 order: 150, posture: "SEATED" },
  { type: "doxology",            title: "Doxology",                                   order: 160, posture: "STANDING" },
  { type: "benediction",         title: "Benediction",                                order: 170, posture: "STANDING" },
  { type: "postlude",            title: "Postlude / Silence for Reflection",          order: 180, posture: "SEATED" },
];

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

// ---------------------
// Web App Router
// ---------------------
function doGet(e) {
  const p = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : "today";

  if (p === "admin") {
    const tpl = HtmlService.createTemplateFromFile("AdminWizard");
    tpl.boot = getBootData_();
    return tpl.evaluate()
      .setTitle("Worship Order — Admin")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (p === "liturgy") {
    const serviceId = resolveServiceId_(e);
    const tpl = HtmlService.createTemplateFromFile("PublicLiturgy");
    tpl.boot = buildPublicBoot_(serviceId);
    return tpl.evaluate()
      .setTitle("Today's Liturgy")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // default: landing
  const serviceId = resolveServiceId_(e);
  const tpl = HtmlService.createTemplateFromFile("PublicToday");
  tpl.boot = buildPublicBoot_(serviceId);
  return tpl.evaluate()
    .setTitle("Today's Liturgy")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------
// Admin Boot
// ---------------------
function getBootData_() {
  return {
    typeMeta: TYPE_META,
    flows: FLOWS,
    postureOptions: ["", "STANDING", "SEATED"],
  };
}

// ---------------------
// Public Boot (safe + helpful)
// ---------------------
function buildPublicBoot_(service_id) {
  try {
    const data = getServiceData_(service_id);
    const service = getServiceInfo_(service_id);
    return {
      ok: true,
      service_id,
      heading: service.title || data.heading || service_id,
      dateLabel: service.date || data.dateLabel || service_id,
      service,
      sections: data.sections || [],
      error: ""
    };
  } catch (err) {
    return {
      ok: false,
      service_id,
      heading: "Liturgy",
      dateLabel: service_id,
      service: {},
      sections: [],
      error: (err && err.message) ? err.message : String(err)
    };
  }
}

function resolveServiceId_(e) {
  const q = (e && e.parameter) ? e.parameter : {};
  const sid = String(q.service_id || "").trim();
  if (sid) return sid;

  const tz = "Asia/Manila";
  const d = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  return `${d}-AM`;
}

// ---------------------
// Sheet helpers
// ---------------------
function getSectionsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SECTIONS_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.SECTIONS_SHEET);

  const requiredHeaders = ["service_id", "order", "type", "title", "body", "posture", "updated_at"];

  const range = sh.getDataRange();
  const values = range.getNumRows() ? range.getValues() : [];
  let headers = values.length ? values[0].map(String) : [];

  const hasAnyHeader = headers.some(h => String(h).trim() !== "");
  if (!hasAnyHeader) {
    sh.clear();
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    headers = requiredHeaders.slice();
  } else {
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

// ---------------------
// Public Data (cached)
// ---------------------
function getServiceData_(service_id) {
  const cache = CacheService.getScriptCache();
  const key = "svc:" + service_id;
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const { sh, col } = getSectionsSheet_();
  const range = sh.getDataRange();
  const values = range.getValues();
  const richValues = range.getRichTextValues();
  if (values.length <= 1) {
    const outEmpty = { heading: formatHeading_(service_id), dateLabel: service_id, sections: [] };
    cache.put(key, JSON.stringify(outEmpty), 60 * 10);
    return outEmpty;
  }

  const sidIdx = col["service_id"];
  const orderIdx = col["order"];
  const typeIdx = col["type"];
  const titleIdx = col["title"];
  const bodyIdx = col["body"];
  const postureIdx = col["posture"];

  const sections = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][sidIdx]).trim() !== service_id) continue;
    const richCell = richValues[i] ? richValues[i][bodyIdx] : null;
    const bodyRich = richTextToRuns_(richCell);
    sections.push({
      order: Number(values[i][orderIdx] || 0),
      type: String(values[i][typeIdx] || ""),
      title: String(values[i][titleIdx] || ""),
      body: String(values[i][bodyIdx] || ""),
      body_rich: bodyRich,
      posture: String(values[i][postureIdx] || ""),
    });
  }

  sections.sort((a,b) => (a.order||0) - (b.order||0));

  const out = {
    heading: formatHeading_(service_id),
    dateLabel: service_id,
    sections,
  };

  cache.put(key, JSON.stringify(out), 60 * 60); // 1 hour
  return out;
}

function formatHeading_(service_id) {
  // "2025-12-21-AM" -> "2025-12-21"
  return String(service_id || "").replace(/-AM$|-PM$/i, "");
}

function richTextToRuns_(richText) {
  if (!richText || !richText.getRuns) return [];
  const runs = richText.getRuns();
  if (!runs || !runs.length) return [];

  const out = [];
  let hasStyle = false;
  runs.forEach((r) => {
    const text = r.getText ? r.getText() : "";
    if (!text) return;
    const style = r.getTextStyle ? r.getTextStyle() : null;
    const bold = style ? !!style.isBold() : false;
    const italic = style ? !!style.isItalic() : false;
    const link = r.getLinkUrl ? (r.getLinkUrl() || "") : "";
    if (bold || italic || link) hasStyle = true;
    out.push({ text, bold, italic, link });
  });

  if (!out.length) return [];
  if (!hasStyle && out.length === 1) return [];
  return out;
}

function getServiceInfo_(service_id) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SERVICES_SHEET);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return {};

  const headers = values[0].map(h => String(h || "").trim().toLowerCase());
  const idx = (name) => headers.indexOf(name);

  const sidIdx = idx("service_id");
  if (sidIdx < 0) return {};

  const titleIdx = idx("title");
  const dateIdx = idx("date");
  const presiderIdx = idx("presider");
  const preacherIdx = idx("preacher");

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][sidIdx] || "").trim() !== service_id) continue;
    return {
      service_id,
      title: titleIdx >= 0 ? String(values[i][titleIdx] || "").trim() : "",
      date: dateIdx >= 0 ? formatServiceDate_(values[i][dateIdx]) : "",
      presider: presiderIdx >= 0 ? String(values[i][presiderIdx] || "").trim() : "",
      preacher: preacherIdx >= 0 ? String(values[i][preacherIdx] || "").trim() : "",
    };
  }

  return {};
}

function formatServiceDate_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMMM d, yyyy");
  }
  return String(value || "").trim();
}

// ---------------------
// Admin CRUD (used by your AdminWizard)
// ---------------------
function getSection(service_id, type) {
  service_id = String(service_id || "").trim();
  type = String(type || "").trim();
  if (!service_id) throw new Error("service_id is required.");
  if (!type) throw new Error("type is required.");

  const { sh, col } = getSectionsSheet_();
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

function ensureServiceSkeleton_(service_id, flow) {
  const types = (FLOWS[flow] || FLOWS.AM || []).slice();
  if (!types.length) throw new Error(`Flow is empty/unknown: ${flow}`);

  const metaByType = {};
  TYPE_META.forEach(m => metaByType[m.type] = m);

  const { sh, headers, col } = getSectionsSheet_();
  const values = sh.getDataRange().getValues();

  const sidIdx = col["service_id"];
  const typeIdx = col["type"];

  const map = {};
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][sidIdx]).trim() !== service_id) continue;
    const t = String(values[i][typeIdx]).trim();
    if (!t) continue;
    if (!map[t]) map[t] = [];
    map[t].push(i + 1);
  }

  let dedupDeleted = 0;
  const rowsToDelete = [];
  Object.keys(map).forEach(t => {
    const rows = map[t];
    if (rows.length > 1) rows.slice(0, -1).forEach(r => rowsToDelete.push(r));
  });
  rowsToDelete.sort((a, b) => b - a).forEach(r => { sh.deleteRow(r); dedupDeleted++; });

  const fresh = sh.getDataRange().getValues();
  const exists = {};
  for (let i = 1; i < fresh.length; i++) {
    if (String(fresh[i][sidIdx]).trim() !== service_id) continue;
    const t = String(fresh[i][typeIdx]).trim();
    exists[t] = true;
  }

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

function upsertSection(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const service_id = String(payload.service_id || "").trim();
    const flow = String(payload.flow || "AM").trim().toUpperCase();
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

    const skeletonInfo = ensureServiceSkeleton_(service_id, flow);

    const { sh, headers, col } = getSectionsSheet_();
    const values = sh.getDataRange().getValues();

    const sidIdx = col["service_id"];
    const typeIdx = col["type"];

    const matches = [];
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][sidIdx]).trim() === service_id &&
          String(values[i][typeIdx]).trim() === type) {
        matches.push(i + 1);
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
    let dedupDeleted = 0;

    if (matches.length) {
      const keptRowIndex = matches[matches.length - 1];
      sh.getRange(keptRowIndex, 1, 1, row.length).setValues([row]);
      mode = "updated";

      const toDelete = matches.slice(0, -1).sort((a, b) => b - a);
      toDelete.forEach(r => sh.deleteRow(r));
      dedupDeleted = toDelete.length;
    } else {
      sh.appendRow(row);
      mode = "created";
    }

    // clear cache for this service so public sees updates immediately
    CacheService.getScriptCache().remove("svc:" + service_id);

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
