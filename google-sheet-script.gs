// Shared lists for the Noah's Green Square pages: bookings, and the expenses / other income ledger.
// Paste this into Extensions > Apps Script of a Google Sheet, then deploy it as a web app
// (Execute as: Me, Who has access: Anyone). The web app link is the key to the lists: keep it private.
//
// Each list is one tab of the sheet. A tab's first row names its columns. A field the tab has no
// column for yet gets a new column by itself, so the pages can gain fields without this script changing.

const LISTS = {
  bookings: {
    tab: 'Bookings',
    fields: ['id', 'name', 'place', 'phone', 'event', 'date', 'from', 'to', 'rent', 'discount', 'deposit',
      'advance', 'balpaid', 'baldate', 'secrefunded', 'booked', 'cancelled', 'cancelreason', 'keepadvance', 'updated'],
  },
  ledger: {
    tab: 'Ledger',
    fields: ['id', 'type', 'date', 'category', 'amount', 'eventId', 'note', 'updated'],
  },
};
const MAX_COLUMNS = 40;

function sheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LISTS[key].tab);
  if (!sh) sh = ss.insertSheet(LISTS[key].tab);
  if (sh.getLastRow() === 0) {
    sh.appendRow(LISTS[key].fields);
    sh.setFrozenRows(1);
  }
  return sh;
}

function header_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function ids_(sh) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
}

function list_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = header_(sh);
  return sh.getRange(2, 1, last - 1, head.length).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const b = {};
      head.forEach(function (f, i) { if (f) b[f] = r[i]; });
      return b;
    });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// The answer names the list it holds, and repeats it as `bookings` for pages written before the ledger
function answer_(key, sh) {
  const rows = list_(sh);
  const out = { ok: true, list: key, rows: rows };
  if (key === 'bookings') out.bookings = rows;
  return json_(out);
}

function doGet(e) {
  const key = (e && e.parameter && e.parameter.list) || 'bookings';
  if (!LISTS[key]) return json_({ ok: false, error: 'Unknown list' });
  return answer_(key, sheet_(key));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const key = req.list || 'bookings';
    if (!LISTS[key]) return json_({ ok: false, error: 'Unknown list' });
    const sh = sheet_(key);
    const ids = ids_(sh);
    const b = req.row || req.booking;
    if (req.action === 'save' && b && b.id) {
      delete b.synced;
      b.updated = new Date().toISOString();
      const head = header_(sh);
      Object.keys(b).forEach(function (k) {
        if (head.indexOf(k) < 0 && /^[a-z][A-Za-z0-9]{0,20}$/.test(k) && head.length < MAX_COLUMNS) {
          sh.getRange(1, head.length + 1).setValue(k);
          head.push(k);
        }
      });
      const row = head.map(function (f) { return b[f] == null ? '' : String(b[f]); });
      const i = ids.indexOf(String(b.id));
      const r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
      // Plain text, so dates, times and phone numbers stay exactly as typed
      sh.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([row]);
    } else if (req.action === 'delete' && req.id) {
      const i = ids.indexOf(String(req.id));
      if (i >= 0) sh.deleteRow(i + 2);
    } else {
      return json_({ ok: false, error: 'Unknown request' });
    }
    return answer_(key, sh);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
