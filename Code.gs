// Pakete Pickliste – Schnittstelle zwischen Tabelle und App
const SHEET_ID = "1swNpYoqhDWqjEwJM5TB7GdG4WLgM38uJUWK34DwV7M4";
const KEY = "HIER_SCHLUESSEL";
const EDITABLE = ["versand", "status", "kiste", "abgegeben_am"];

function sheet_() { return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Berlin", "yyyy-MM-dd'T'HH:mm:ss");
  return v === null || v === undefined ? "" : String(v);
}
function rows_() {
  const v = sheet_().getDataRange().getValues();
  const h = v.shift().map(String);
  return v.filter(r => r[0] !== "").map(r => { const o = {}; h.forEach((k, i) => o[k] = cell_(r[i])); return o; });
}

function doGet(e) {
  if ((e.parameter.k || "") !== KEY) return out_({ ok: false, error: "key" });
  return out_({ ok: true, items: rows_() });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "json" }); }
  if (b.k !== KEY) return out_({ ok: false, error: "key" });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_();
    const v = sh.getDataRange().getValues();
    const h = v[0].map(String);
    const r = v.findIndex((row, i) => i > 0 && String(row[0]) === String(b.id));
    if (r < 1) return out_({ ok: false, error: "id" });
    Object.keys(b.patch || {}).forEach(k => {
      const c = h.indexOf(k);
      if (c >= 0 && EDITABLE.indexOf(k) >= 0) sh.getRange(r + 1, c + 1).setValue(String(b.patch[k] ?? ""));
    });
    const ca = h.indexOf("aktualisiert");
    if (ca >= 0) sh.getRange(r + 1, ca + 1).setValue(new Date());
    return out_({ ok: true, items: rows_() });
  } finally { lock.releaseLock(); }
}
