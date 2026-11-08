/**
 * Frog Wednesday practice RSVPs: a tiny backend for the site's Practice tab.
 *
 * One-time setup (about 3 minutes):
 *   1. Create a new Google Sheet, e.g. "Frog Practice RSVPs".
 *   2. Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: Me      Who has access: Anyone
 *      Click Deploy and approve the permissions it asks for.
 *   4. Copy the Web app URL (ends in /exec) and put it in PRACTICE_API in site/app.js.
 *
 * Each RSVP is one row on the "RSVPs" tab: date | name | status | note | updated.
 * Answering again with the same name for the same date replaces that row, so people can change their answer.
 * After editing this script later: Deploy → Manage deployments → edit → Version: New version (the URL stays the same).
 */
const TAB = 'RSVPs';
const STATUSES = ['in', 'maybe', 'out'];

function tab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.appendRow(['date', 'name', 'status', 'note', 'updated']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(s, max) {
  return String(s || '').replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

function list_(sh, date) {
  const rows = sh.getDataRange().getDisplayValues().slice(1);
  return rows
    .filter((r) => r[0] === date)
    .map((r) => ({ name: r[1], status: r[2], note: r[3], updated: r[4] }));
}

// Everything goes through GET so the static site can call it without CORS trouble.
//   ?action=list&date=2026-11-11
//   ?action=set&date=2026-11-11&name=Sam&status=in&note=late
function doGet(e) {
  const p = (e && e.parameter) || {};
  const date = clean_(p.date, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'bad date' });
  const sh = tab_();

  if (p.action === 'set') {
    const name = clean_(p.name, 40);
    const status = clean_(p.status, 10).toLowerCase();
    const note = clean_(p.note, 80);
    if (!name) return json_({ ok: false, error: 'name required' });
    if (STATUSES.indexOf(status) < 0) return json_({ ok: false, error: 'bad status' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const values = sh.getDataRange().getDisplayValues();
      const key = name.toLowerCase();
      const now = new Date().toISOString();
      // Plain text so Sheets doesn't turn the date or names into something else.
      const row = ["'" + date, "'" + name, status, "'" + note, now];
      let found = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === date && values[i][1].toLowerCase() === key) { found = i + 1; break; }
      }
      if (found > 0) sh.getRange(found, 1, 1, 5).setValues([row]);
      else sh.appendRow(row);
    } finally {
      lock.releaseLock();
    }
  }
  return json_({ ok: true, date: date, rsvps: list_(sh, date) });
}
