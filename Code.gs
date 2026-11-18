// Pakete Pickliste – Schnittstelle zwischen Tabelle und App
const SHEET_ID = "1swNpYoqhDWqjEwJM5TB7GdG4WLgM38uJUWK34DwV7M4";
const ARTIKEL_ID = "1FjA3i0qhOSlxKn5IAbktO8onHtr17yncjuf60aOj6fw"; // Vinted Artikelliste, Tab "Artikel"
const KEY = "HIER_SCHLUESSEL";        // Pickliste (Mutter)
const OWNER_KEY = "HIER_SCHLUESSEL_JULY";  // Dashboard (July)
const EDITABLE = ["versand", "status", "kiste", "abgegeben_am"];
const MONEY = ["preis", "einstellpreis"];

function sheet_() { return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Berlin", "yyyy-MM-dd'T'HH:mm:ss");
  if (v && typeof v === "object" && typeof v.getTime === "function") return Utilities.formatDate(new Date(v.getTime()), "Europe/Berlin", "yyyy-MM-dd'T'HH:mm:ss");
  return v === null || v === undefined ? "" : String(v);
}
function rows_(owner) {
  const v = sheet_().getDataRange().getValues();
  const h = v.shift().map(String);
  return v.filter(r => r[0] !== "").map(r => {
    const o = {}; h.forEach((k, i) => { if (owner || MONEY.indexOf(k) < 0) o[k] = cell_(r[i]); }); return o;
  });
}
function stock_() {
  const sh = SpreadsheetApp.openById(ARTIKEL_ID).getSheetByName("Artikel");
  const v = sh.getDataRange().getValues();
  const h = v.shift().map(String);
  const col = n => h.indexOf(n);
  const want = { artikel: "Artikel", status: "Status", preis: "Einstellpreis (€)", eingestellt: "Eingestellt am", favoriten: "Favoriten", ordner: "Fotos-Ordner" };
  return v.filter(r => r[col("Artikel")] !== "").map(r => {
    const o = {}; Object.keys(want).forEach(k => { const c = col(want[k]); o[k] = c >= 0 ? cell_(r[c]) : ""; }); return o;
  });
}
function favs_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("favoriten");
  if (!sh) return [];
  const v = sh.getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== "").map(r => ({ zeit: cell_(r[0]).replace(" ", "T"), artikel: cell_(r[1]) }));
}
function role_(k) { return k === OWNER_KEY ? "owner" : (k === KEY ? "helper" : ""); }

function doGet(e) {
  const role = role_(e.parameter.k || "");
  if (!role) return out_({ ok: false, error: "key" });
  const res = { ok: true, items: rows_(role === "owner") };
  if (role === "owner" && e.parameter.dash) { res.stock = stock_(); res.favs = favs_(); }
  return out_(res);
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "json" }); }
  const role = role_(b.k);
  if (!role) return out_({ ok: false, error: "key" });
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
    return out_({ ok: true, items: rows_(role === "owner") });
  } finally { lock.releaseLock(); }
}
