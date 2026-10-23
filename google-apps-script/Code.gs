/**
 * Varsha & Chellam wedding site: backend (Google Apps Script web app)
 *
 * One Google Sheet (three tabs) plus one Drive folder, all filled automatically:
 *   - FIRST tab        : the RSVP form (the tab your RSVPs already go to; it is not renamed)
 *   - "YouTube Live"   : guests asking for the YouTube live link (name, phone, email)
 *   - "Upload Photos"  : guests asking for the photo upload link (name, phone, email)
 *   - Drive folder    : photos uploaded from the private /share-photos page
 * The admin page (/admin) reads the RSVP tab with the admin password, and can clear it.
 *
 * Setup (about 3 minutes). Your project already exists, so this only updates it:
 *   1. Go to script.google.com > open your wedding project. Replace ALL the code with this file.
 *   2. Project Settings (gear icon) > Script properties: ADMIN_KEY must still be set to your password.
 *   3. Pick the function "setup" in the toolbar and press Run. Approve the permissions if asked
 *      (Drive access is needed again for the photo folder). The log should say "Setup OK".
 *   4. Deploy > Manage deployments > pencil (edit) > Version: New version > Deploy.
 *      The web app URL stays the same, so nothing changes in config.js.
 *
 * The sheet stays private. The script runs as you, and the public can only add rows.
 * Nobody can read or delete anything without ADMIN_KEY.
 */

const SHEET_ID = '1nsryr4OJdMiHY2SsBhxc5V-ci9dTbxoEPM13In_YJRk';
const DRIVE_FOLDER_ID = '1goi613i4MlBu9TyKvfgKymkrNwAsyYCi';
const MAX_PHOTO_BYTES = 12 * 1024 * 1024;
const PHOTO_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/heic': 'heic', 'image/heif': 'heif', 'image/gif': 'gif' };

const TABS = {
  rsvp:     { name: '',              headers: ['Timestamp (IST)', 'Name', 'Phone', 'Email', 'Attending', 'Events', 'Guests', 'Message', 'Language'], text: [1, 2, 3, 4, 5, 6, 8, 9] },
  livereq:  { name: 'YouTube Live',  headers: ['Timestamp (IST)', 'Name', 'Phone', 'Email', 'Language'], text: [1, 2, 3, 4, 5] },
  photoreq: { name: 'Upload Photos', headers: ['Timestamp (IST)', 'Name', 'Phone', 'Email', 'Language'], text: [1, 2, 3, 4, 5] }
};
const EVENTS = ['Reception (31 Oct)', 'Muhurtham (1 Nov)'];

function doGet() {
  return out_({ ok: true, service: 'wedding-backend' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.website) return out_({ ok: true });            // honeypot field: real guests leave it empty
    switch (body.action) {
      case 'rsvp':     return out_(saveRsvp_(body));
      case 'livereq':  return out_(saveRequest_('livereq', body));
      case 'photoreq': return out_(saveRequest_('photoreq', body));
      case 'photo':    return out_(savePhoto_(body));
      case 'list':     return out_(listRsvps_(body));
      case 'clear':    return out_(clearRsvps_(body));
      default:         return out_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    console.error(err);
    return out_({ ok: false, error: 'Server error' });
  }
}

/* ---------- RSVP ---------- */
function saveRsvp_(b) {
  const name = clean_(b.name, 100);
  const attending = b.attending === 'Yes' ? 'Yes' : (b.attending === 'No' ? 'No' : '');
  if (!name || !attending) return { ok: false, error: 'Name and attendance are required' };

  const guests = attending === 'Yes' ? Math.min(Math.max(parseInt(b.guests, 10) || 1, 1), 20) : 0;
  const events = attending === 'Yes' && Array.isArray(b.events)
    ? b.events.filter(function (x) { return EVENTS.indexOf(x) > -1; }).join(', ') : '';
  appendRow_('rsvp', [
    now_(), name, clean_(b.phone, 30), clean_(b.email, 120), attending, events, guests,
    clean_(b.message, 600), b.lang === 'ta' ? 'Tamil' : 'English'
  ]);
  return { ok: true };
}

/* ---------- Live-link and photo-link requests: name and phone required, email may be blank ---------- */
function saveRequest_(kind, b) {
  const name = clean_(b.name, 100);
  const phone = clean_(b.phone, 30);
  if (!name) return { ok: false, error: 'Name is required' };
  if (phone.replace(/\D/g, '').length < 7) return { ok: false, error: 'A valid phone number is required' };
  appendRow_(kind, [now_(), name, phone, clean_(b.email, 120), b.lang === 'ta' ? 'Tamil' : 'English']);
  return { ok: true };
}

/* ---------- Photos from the /share-photos page go into the Drive folder ---------- */
function savePhoto_(b) {
  const mime = String(b.mime || '').toLowerCase();
  const ext = PHOTO_TYPES[mime];
  if (!ext || !b.data) return { ok: false, error: 'Unsupported file' };
  const who = clean_(b.uploader, 60);
  if (!who) return { ok: false, error: 'Name is required' };
  const bytes = Utilities.base64Decode(b.data);
  if (bytes.length > MAX_PHOTO_BYTES) return { ok: false, error: 'File too large' };

  const slug = who.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'guest';
  const stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
  const fileName = stamp + '_' + slug + '_' + Math.floor(Math.random() * 9000 + 1000) + '.' + ext;
  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(Utilities.newBlob(bytes, mime, fileName));
  file.setDescription('Uploaded by: ' + who);
  return { ok: true };
}

/* ---------- sheet helpers ---------- */
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
}

function tabSheet_(kind) {
  const t = TABS[kind];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = t.name ? (ss.getSheetByName(t.name) || ss.insertSheet(t.name)) : ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold').setBackground('#F6E3BC');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, t.headers.length, 160);
    if (kind === 'rsvp') sheet.setColumnWidth(8, 320);
  }
  return sheet;
}

function appendRow_(kind, row) {
  const t = TABS[kind];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = tabSheet_(kind);
    const r = sheet.getLastRow() + 1;
    const range = sheet.getRange(r, 1, 1, t.headers.length);
    range.setNumberFormats([t.headers.map(function (_, i) { return t.text.indexOf(i + 1) > -1 ? '@' : '0'; })]);   // text format keeps +91 and stops formulas
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Admin ---------- */
function authError_(b) {
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('admin_fails') || '0', 10);
  if (fails >= 10) return { ok: false, error: 'Too many attempts. Try again in 10 minutes.' };

  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key) return { ok: false, error: 'ADMIN_KEY is not set in Script properties.' };
  if (!b.key || String(b.key) !== key) {
    cache.put('admin_fails', String(fails + 1), 600);
    return { ok: false, error: 'Wrong password' };
  }
  cache.remove('admin_fails');
  return null;
}

function listRsvps_(b) {
  const bad = authError_(b);
  if (bad) return bad;
  const t = TABS.rsvp, sheet = tabSheet_('rsvp');
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, t.headers.length).getDisplayValues() : [];
  return {
    ok: true,
    headers: t.headers,
    rows: rows,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit',
    folderUrl: 'https://drive.google.com/drive/folders/' + DRIVE_FOLDER_ID
  };
}

/* delete every RSVP row (the header row stays) */
function clearRsvps_(b) {
  const bad = authError_(b);
  if (bad) return bad;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = tabSheet_('rsvp');
    const n = sheet.getLastRow() - 1;
    if (n > 0) sheet.deleteRows(2, n);
    return { ok: true, deleted: Math.max(n, 0) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor: authorises access, creates the two new tabs, checks the photo folder. */
function setup() {
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) {
    throw new Error('Add the ADMIN_KEY script property first (Project Settings > Script properties).');
  }
  const names = Object.keys(TABS).map(function (k) { return tabSheet_(k).getName(); });
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  console.log('Setup OK. Tabs ready: ' + names.join(', ') + '. Photo folder: ' + folder.getName());
}
