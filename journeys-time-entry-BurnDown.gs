/**
 * Journeys Time Entry → Burn Down Record
 *
 * TEAM SETUP (recommended): one person deploys this once for the whole team.
 *   1. script.google.com → New project (signed in with an account that can EDIT the Burn Down Record).
 *      Delete the starter code, paste this file, save.
 *   2. Deploy → New deployment → Web app
 *        Execute as: Me
 *        Who has access: Anyone within <your Google Workspace domain>   (only signed-in coworkers can use it)
 *   3. In the extension: ⚙ → Burn Down sheet → paste the sheet link and this Web app URL → Save & test
 *      → "Copy team setup code" and share the code with the team. Each person pastes it in ⚙ and
 *      picks their tab; nobody else deploys anything.
 *   UPDATING: paste the new code, save, Deploy → Manage deployments → ✏ → Version: New version → Deploy.
 *   Everyone gets the update at once (the URL stays the same).
 *
 * PERSONAL SETUP: the same steps, with "Who has access: Only myself", and use your own URL.
 *
 * Writes one row per billable entry to the first row whose Customer Name (column B) is empty.
 * Columns written: B Customer, C Account #, D Date, E Hours, G CSP Level, H Allotted, I Bucket.
 * Columns A (Burn Down ID), J (Total), K–M are left alone (pre-filled / formulas).
 * It can also list your rows for one customer and week (used by the extension's Sync view), and set a
 * row's Hours to 0 when its Journeys entry is deleted (the row is kept).
 */
const SHEET_ID = '';   // optional fallback; normally the extension sends the sheet ID (⚙ → Sheet link)
const DEFAULT_TAB = '';                    // optional; the extension sends your tab name
const TOKEN = '';           // optional: set a secret here AND in the extension settings
const SCRIPT_VERSION = 7;  // 2 'list' · 3 'zero' · 4 date check · 5 'sethours' · 6 'customers' · 7 'tabs' + caller email (team setup)

function doGet(e) {
  const q = (e && e.parameter) || {};
  return handle_({ action: 'ping', sheetId: q.sheetId, tab: q.tab || DEFAULT_TAB, token: q.token });
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
    let id = String(req.sheetId || SHEET_ID || '').trim();
    const fromUrl = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);   // a full sheet address works too
    if (fromUrl) id = fromUrl[1];
    if (!id) return out_({ ok: false, error: 'No sheet set. Paste the Burn Down sheet link in the extension settings (⚙ → Sheet link).' });
    let book;
    try { book = SpreadsheetApp.openById(id); }
    catch (err) { return out_({ ok: false, error: 'Can\'t open that sheet. Check the Sheet link, and that this Google account can edit it.' }); }
    if (req.action === 'tabs') {
      // Tab names for the extension's tab picker, plus who is calling (works when the web app is shared within the domain)
      return out_({ ok: true, version: SCRIPT_VERSION, tabs: book.getSheets().map(t => t.getName()), you: caller_() });
    }
    const sh = book.getSheetByName(req.tab || DEFAULT_TAB);
    if (!sh) return out_({ ok: false, version: SCRIPT_VERSION, error: 'Tab not found: ' + (req.tab || DEFAULT_TAB || '(none)') });
    if (req.action === 'ping') return out_({ ok: true, version: SCRIPT_VERSION, tab: sh.getName(), nextRow: nextRow_(sh), you: caller_() });

    if (req.action === 'zero') {
      // For each {customer, date, hours}: find the matching row and set Hours (E) to 0 — the row itself is kept
      const last = sh.getLastRow();
      const tz = book.getSpreadsheetTimeZone();
      const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
      const zeroed = [];
      for (const r of (req.rows || [])) {
        const want = norm_(r.customer);
        let hit = -1;
        const byRow = Number(r.row) - 2;   // exact row when the extension knows it (still checked below)
        if (byRow >= 0 && byRow < vals.length && norm_(vals[byRow][1]) === want && toYmd_(vals[byRow][3], tz) === r.date) hit = byRow;
        for (let i = 0; hit < 0 && i < vals.length; i++) {
          const v = vals[i];
          if (norm_(v[1]) === want && toYmd_(v[3], tz) === r.date && Math.abs((Number(v[4]) || 0) - Number(r.hours)) < 1e-6) { hit = i; break; }
        }
        if (hit < 0) { zeroed.push({ ref: r.ref, row: null }); continue; }
        sh.getRange(hit + 2, 5).setValue(0);
        vals[hit][4] = 0;
        zeroed.push({ ref: r.ref, row: hit + 2, id: String(vals[hit][0]) });
      }
      SpreadsheetApp.flush();
      return out_({ ok: true, version: SCRIPT_VERSION, zeroed });
    }

    if (req.action === 'customers') {
      // One entry per customer name in this tab, taken from its most recent row
      const last = sh.getLastRow();
      const byName = {};
      if (last >= 2) {
        sh.getRange(2, 1, last - 1, 9).getValues().forEach(v => {
          const name = String(v[1] == null ? '' : v[1]).trim();
          if (!name) return;
          const k = norm_(name);
          const prev = byName[k] || { spellings: {} };
          prev.spellings[name] = (prev.spellings[name] || 0) + 1;
          byName[k] = { spellings: prev.spellings, account: String(v[2] == null ? '' : v[2]).trim(), csp: String(v[6] || '').trim(),
            allotted: v[7] === '' ? '' : Number(v[7]) || 0, bucket: v[8] === '' ? '' : Number(v[8]) || 0 };
        });
      }
      // name = the spelling used most often; account/CSP/hours = the most recent row
      const customers = Object.keys(byName).map(k => {
        const e = byName[k];
        const customer = Object.keys(e.spellings).sort((a, b) => e.spellings[b] - e.spellings[a])[0];
        return { customer, account: e.account, csp: e.csp, allotted: e.allotted, bucket: e.bucket };
      }).sort((a, b) => a.customer.localeCompare(b.customer));
      return out_({ ok: true, version: SCRIPT_VERSION, customers });
    }

    if (req.action === 'sethours') {
      // For each {row, customer, date, newHours}: check the row still belongs to that customer/date, then set Hours (E)
      const last = sh.getLastRow();
      const tz = book.getSpreadsheetTimeZone();
      const updated = [];
      for (const r of (req.rows || [])) {
        const n = Number(r.row);
        const h = Number(r.newHours);
        if (!(n >= 2 && n <= last) || !(h >= 0 && h <= 24)) { updated.push({ ref: r.ref, row: null, error: 'Bad row or hours' }); continue; }
        const v = sh.getRange(n, 1, 1, 5).getValues()[0];
        if (norm_(v[1]) !== norm_(r.customer) || toYmd_(v[3], tz) !== r.date) { updated.push({ ref: r.ref, row: null, error: 'Row changed in the sheet; sync again' }); continue; }
        sh.getRange(n, 5).setValue(h);
        updated.push({ ref: r.ref, row: n, id: String(v[0]), hours: h });
      }
      SpreadsheetApp.flush();
      return out_({ ok: true, version: SCRIPT_VERSION, updated });
    }

    if (req.action === 'list') {
      // Rows for one customer (column B, case-insensitive) with dates from..to (YYYY-MM-DD, inclusive)
      const last = sh.getLastRow();
      const rows = [];
      if (last >= 2) {
        const tz = book.getSpreadsheetTimeZone();
        const want = norm_(req.customer);
        sh.getRange(2, 1, last - 1, 5).getValues().forEach((v, i) => {
          if (!norm_(v[1]) || norm_(v[1]) !== want) return;
          const d = toYmd_(v[3], tz);
          if (!d || d < req.from || d > req.to) return;
          rows.push({ row: i + 2, id: String(v[0]), customer: String(v[1]), account: String(v[2]), date: d, hours: Number(v[4]) || 0 });
        });
      }
      return out_({ ok: true, version: SCRIPT_VERSION, rows });
    }


    const written = [];
    for (const r of (req.rows || [])) {
      const m = String(r.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);   // 'YYYY-MM-DD'
      if (!m) { written.push({ ref: r.ref, row: null, error: 'No valid date (' + r.date + ')' }); continue; }   // skip, never write a dateless row
      const n = nextRow_(sh);
      sh.getRange(n, 2, 1, 4).setValues([[r.customer, Number(r.account) || r.account, new Date(+m[1], +m[2] - 1, +m[3], 12), Number(r.hours)]]);
      sh.getRange(n, 4).setNumberFormat('M/d/yyyy');
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

function caller_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function norm_(x) { return String(x == null ? '' : x).trim().toLowerCase(); }

function toYmd_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);   // M/D/YYYY text
  if (m) return `${m[3]}-${('0' + m[1]).slice(-2)}-${('0' + m[2]).slice(-2)}`;
  const iso = String(v || '').trim().match(/^\d{4}-\d{2}-\d{2}/);
  return iso ? iso[0] : '';
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
