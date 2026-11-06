/**
 * Jenna & Mathew's Bali wedding: RSVP receiver.
 *
 * Paste this into Extensions > Apps Script in your RSVP Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * Each RSVP becomes one row on the "RSVPs" tab, and each guest becomes one row on
 * the "Guests" tab (handy for catering and dietary counts). If someone sends their
 * RSVP again with the same email (or the same name if they left email blank),
 * their old rows are replaced, so the sheet always shows their latest answer.
 */

const RSVP_TAB = 'RSVPs';
const GUEST_TAB = 'Guests';

const RSVP_HEADERS = ['Updated', 'Key', 'Name', 'Email', 'Mobile', 'Attending',
  'Adults', 'Teens', 'Children', 'Guest names', 'Arriving', 'Leaving',
  'Accommodation', 'Welcome dinner', 'Song request', 'Message'];

const GUEST_HEADERS = ['Updated', 'Key', 'RSVP from', 'Guest name', 'Age group',
  'Dietary requirements', 'Allergies / other needs'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return json({ ok: true }); // spam trap: real guests never fill this in
    if (!d.name || !d.attending) return json({ ok: false, error: 'missing fields' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rsvps = tab(ss, RSVP_TAB, RSVP_HEADERS);
    const guests = tab(ss, GUEST_TAB, GUEST_HEADERS);

    const key = String(d.email || d.name).trim().toLowerCase();
    removeKey(rsvps, key);
    removeKey(guests, key);

    const now = new Date();
    const list = Array.isArray(d.guests) ? d.guests : [];
    const count = age => list.filter(g => g.age === age).length;

    rsvps.appendRow([now, key, d.name, d.email, d.phone,
      d.attending === 'yes' ? 'Yes' : 'No',
      count('adult'), count('teen'), count('child'),
      list.map(g => g.name).join(', '),
      d.arrive, d.depart, d.stay, d.welcome, d.song, d.message].map(clean));

    list.forEach(g => guests.appendRow([now, key, d.name, g.name,
      { adult: 'Adult', teen: 'Teen (12–17)', child: 'Child (under 12)' }[g.age] || g.age,
      (g.diet || []).join(', '), g.notes].map(clean)));

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Visiting the web app URL in a browser shows this, so you can check it's live.
function doGet() {
  return ContentService.createTextOutput('RSVP receiver is running.');
}

function tab(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function removeKey(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const keys = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]).toLowerCase() === key) sh.deleteRow(i + 2);
  }
}

// Stops a guest's typing being treated as a spreadsheet formula.
function clean(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
