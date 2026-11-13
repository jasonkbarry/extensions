/**
 * Process Ideas: backend for https://jasonkbarry.github.io/extensions/ideas.html
 *
 * Stores ideas in a Google Sheet ("Process Ideas") and screenshots in a Drive folder, both owned by
 * whoever deploys this script. The sheet and folder are created automatically the first time.
 *
 * SETUP (one time)
 *   1. Go to script.google.com and click New project. Delete the starter code, paste this file, and save.
 *   2. Run ▸ choose "setup" ▸ Run. Approve the permissions (Sheets and Drive).
 *      Execution log prints the links to the new sheet and screenshot folder.
 *   3. Deploy ▸ New deployment ▸ Web app
 *        Execute as: Me
 *        Who has access: Anyone      (the Ideas page is public and people don't sign in)
 *      Copy the Web app URL (ends in /exec) and put it in ideas.html (IDEAS_API).
 *
 * EVALUATING IDEAS: open the sheet. Set the Status column (dropdown) and type notes in Evaluation.
 * The Ideas page shows both. Delete a row to remove an idea.
 *
 * CHANGING THIS SCRIPT LATER: paste the new code, save, then Deploy ▸ Manage deployments ▸ ✏ ▸
 * Version: New version ▸ Deploy (keeps the same URL).
 */
const SHEET_NAME = 'Ideas';
const STATUSES = ['New', 'Evaluating', 'Planned', 'Building', 'Done', 'Not now'];
const HEAD = ['ID', 'Submitted', 'Name', 'Date', 'Process frustration', 'Business case', 'Ideas to try', 'Screenshots', 'Status', 'Evaluation', 'Votes'];
const COL = Object.fromEntries(HEAD.map((h, i) => [h, i]));       // 0-based
const MAX_IMAGES = 6, MAX_IMAGE_BYTES = 6 * 1024 * 1024, MAX_TEXT = 5000;
const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];

function setup() {
  const s = ensure_();
  Logger.log('Ideas sheet: ' + s.book.getUrl());
  Logger.log('Screenshots folder: ' + s.folder.getUrl());
}

function ensure_() {
  const props = PropertiesService.getScriptProperties();
  let book = null, folder = null;
  const sid = props.getProperty('SHEET_ID'), fid = props.getProperty('FOLDER_ID');
  if (sid) { try { book = SpreadsheetApp.openById(sid); } catch (e) {} }
  if (!book) {
    book = SpreadsheetApp.create('Process Ideas');
    props.setProperty('SHEET_ID', book.getId());
    const sh = book.getSheets()[0];
    sh.setName(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
    sh.getRange(2, COL.Status + 1, sh.getMaxRows() - 1, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
    sh.setColumnWidth(COL['Process frustration'] + 1, 320);
    sh.setColumnWidth(COL['Business case'] + 1, 320);
    sh.setColumnWidth(COL['Ideas to try'] + 1, 260);
    sh.setColumnWidth(COL.Evaluation + 1, 320);
  }
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) {} }
  if (!folder) {
    folder = DriveApp.createFolder('Process Ideas - Screenshots');
    props.setProperty('FOLDER_ID', folder.getId());
  }
  return { book, folder, sh: book.getSheetByName(SHEET_NAME) || book.getSheets()[0] };
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const clip_ = (v, n) => String(v == null ? '' : v).trim().slice(0, n || MAX_TEXT);

function rows_(s) {
  const last = s.sh.getLastRow();
  return last < 2 ? [] : s.sh.getRange(2, 1, last - 1, HEAD.length).getValues();
}

function list_(s) {
  return rows_(s).filter(r => r[COL.ID]).map(r => ({
    id: String(r[COL.ID]),
    submitted: r[COL.Submitted] instanceof Date ? r[COL.Submitted].toISOString() : String(r[COL.Submitted]),
    name: String(r[COL.Name]),
    date: r[COL.Date] instanceof Date ? Utilities.formatDate(r[COL.Date], Session.getScriptTimeZone(), 'M/d/yyyy') : String(r[COL.Date]),
    frustration: String(r[COL['Process frustration']]),
    business: String(r[COL['Business case']]),
    ideas: String(r[COL['Ideas to try']]),
    screenshots: String(r[COL.Screenshots] || '').split(',').map(x => x.trim()).filter(Boolean),
    status: String(r[COL.Status] || 'New'),
    evaluation: String(r[COL.Evaluation] || ''),
    votes: Number(r[COL.Votes]) || 0
  }));
}

function doGet(e) {
  const q = (e && e.parameter) || {};
  const s = ensure_();
  if (q.img) {
    // Only serve files that belong to an idea, never arbitrary Drive files
    const allowed = rows_(s).some(r => String(r[COL.Screenshots] || '').split(',').map(x => x.trim()).indexOf(q.img) >= 0);
    if (!allowed) return json_({ ok: false, error: 'Not found' });
    const blob = DriveApp.getFileById(q.img).getBlob();
    return json_({ ok: true, type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) });
  }
  return json_({ ok: true, ideas: list_(s) });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  const s = ensure_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.action === 'vote') return vote_(s, req.id);
    return submit_(s, req);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function submit_(s, req) {
  if (req.website) return json_({ ok: true, id: 'thanks' });          // honeypot field filled = bot
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('rate') || 0);
  if (n >= 30) return json_({ ok: false, error: 'Lots of ideas just came in. Please try again in a few minutes.' });
  cache.put('rate', String(n + 1), 600);

  const name = clip_(req.name, 100), frustration = clip_(req.frustration), business = clip_(req.business);
  if (!name || !frustration || !business) return json_({ ok: false, error: 'Name, process frustration and business case are required.' });

  const ids = [];
  for (const img of (req.images || []).slice(0, MAX_IMAGES)) {
    if (IMAGE_TYPES.indexOf(img.type) < 0) continue;
    const bytes = Utilities.base64Decode(String(img.data || ''));
    if (!bytes.length || bytes.length > MAX_IMAGE_BYTES) continue;
    const ext = img.type.split('/')[1].replace('jpeg', 'jpg');
    const f = s.folder.createFile(Utilities.newBlob(bytes, img.type, 'idea-' + Date.now() + '-' + ids.length + '.' + ext));
    ids.push(f.getId());
  }

  const id = 'IDEA-' + ('000' + (s.sh.getLastRow())).slice(-4);
  s.sh.appendRow([id, new Date(), name, clip_(req.date, 20), frustration, business, clip_(req.ideas), ids.join(','), 'New', '', 0]);
  return json_({ ok: true, id });
}

function vote_(s, id) {
  const rows = rows_(s);
  const i = rows.findIndex(r => String(r[COL.ID]) === String(id));
  if (i < 0) return json_({ ok: false, error: 'Idea not found' });
  const cell = s.sh.getRange(i + 2, COL.Votes + 1);
  const v = (Number(cell.getValue()) || 0) + 1;
  cell.setValue(v);
  return json_({ ok: true, votes: v });
}
