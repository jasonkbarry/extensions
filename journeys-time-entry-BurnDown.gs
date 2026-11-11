/**
 * Journeys Time Entry → Burn Down Record
 * Paste this into a NEW Apps Script project (script.google.com, signed in with the
 * Google account that can edit the Burn Down Record sheet), then:
 *   Deploy → New deployment → Type: Web app
 *     Execute as: Me
 *     Who has access: Only myself
 *   Copy the Web app URL into the extension (⚙ → Burn Down sheet).
 *
 * Writes one row per billable entry to the first row whose Customer Name (column B) is empty.
 * Columns written: B Customer, C Account #, D Date, E Hours, G CSP Level, H Allotted, I Bucket.
 * Columns A (Burn Down ID), J (Total), K–M are left alone (pre-filled / formulas).
 */
const SHEET_ID = 'PASTE_SHEET_ID_HERE';   // the long ID in the sheet's URL: docs.google.com/spreadsheets/d/<SHEET_ID>/edit
const DEFAULT_TAB = '';                    // optional; the extension sends your tab name
const TOKEN = '';   // optional: set a secret here AND in the extension settings

function doGet(e) {
  return handle_({ action: 'ping', tab: (e && e.parameter && e.parameter.tab) || DEFAULT_TAB, token: e && e.parameter && e.parameter.token });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Bad request body' }); }
  return handle_(req);
}

function handle_(req) {
  if (TOKEN && req.token !== TOKEN) return out_({ ok: false, error: 'Wrong token' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(req.tab || DEFAULT_TAB);
    if (!sh) return out_({ ok: false, error: 'Tab not found: ' + (req.tab || DEFAULT_TAB) });
    if (req.action === 'ping') return out_({ ok: true, tab: sh.getName(), nextRow: nextRow_(sh) });

    const written = [];
    for (const r of (req.rows || [])) {
      const n = nextRow_(sh);
      const p = String(r.date).split('-').map(Number);           // 'YYYY-MM-DD'
      sh.getRange(n, 2, 1, 4).setValues([[r.customer, Number(r.account) || r.account, new Date(p[0], p[1] - 1, p[2]), Number(r.hours)]]);
      sh.getRange(n, 7, 1, 3).setValues([[r.csp, Number(r.allotted) || 0, Number(r.bucket) || 0]]);
      written.push({ ref: r.ref, row: n, id: String(sh.getRange(n, 1).getValue()) });
    }
    SpreadsheetApp.flush();
    return out_({ ok: true, written });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function nextRow_(sh) {
  const last = sh.getMaxRows();
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '') return i + 2;
  }
  sh.insertRowAfter(last);
  return last + 1;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
