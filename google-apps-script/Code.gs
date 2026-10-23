/**
 * Varsha & Chellam wedding site: backend (Google Apps Script web app)
 *
 * What it does
 *   - "rsvp"  : appends one row per RSVP to your Google Sheet
 *   - "photo" : saves each guest photo into your Google Drive folder
 *   - "list"  : returns the RSVP rows to the admin page, only with the admin password
 *
 * One-time setup (about 5 minutes)
 *   1. Open your Google Sheet. Copy its ID from the URL:
 *        https://docs.google.com/spreadsheets/d/<THIS_PART>/edit
 *   2. Open your Drive folder for photos. Copy its ID from the URL:
 *        https://drive.google.com/drive/folders/<THIS_PART>
 *   3. Go to script.google.com > New project. Paste this whole file, and fill in the two IDs below.
 *   4. Project Settings (gear icon) > Script properties > Add property
 *        ADMIN_KEY = a password only you know     (this protects the RSVP list)
 *   5. Pick the function "setup" in the toolbar and press Run. Approve the permissions
 *      (Google shows an "unverified app" screen for your own script; Advanced > Go to project).
 *      The log should say "Setup OK".
 *   6. Deploy > New deployment > type "Web app"
 *        Execute as: Me          Who has access: Anyone
 *      Press Deploy and copy the Web app URL (ends in /exec).
 *   7. Paste that URL into SCRIPT_URL in config.js on the website. Done.
 *
 * If you change this file later: Deploy > Manage deployments > edit (pencil) > Version: New version.
 *
 * The sheet and folder stay private. The script runs as you, and the public can only
 * add RSVPs and photos. They cannot read or delete anything. The RSVP list is readable
 * only with ADMIN_KEY, through admin.html.
 */

const SHEET_ID = '1nsryr4OJdMiHY2SsBhxc5V-ci9dTbxoEPM13In_YJRk';
const DRIVE_FOLDER_ID = '1goi613i4MlBu9TyKvfgKymkrNwAsyYCi';
const RSVP_TAB = '';   // leave empty to use the FIRST tab of the sheet, or put a tab name here

const HEADERS = ['Timestamp (IST)', 'Name', 'Phone', 'Email', 'Attending', 'Events', 'Guests', 'Message', 'Language'];
const TEXT_COLS = [1, 2, 3, 4, 5, 6, 8, 9];       // 1-based columns stored as plain text (keeps +91 and stops formulas)
const MAX_PHOTO_BYTES = 12 * 1024 * 1024;
const PHOTO_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/heic': 'heic', 'image/heif': 'heif', 'image/gif': 'gif' };
const EVENTS = ['Reception (31 Oct)', 'Muhurtham (1 Nov)'];

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'settings') {   // public: the site reads the live link from here
    return out_({ ok: true, youtubeUrl: PropertiesService.getScriptProperties().getProperty('YOUTUBE_URL') || '' });
  }
  return out_({ ok: true, service: 'wedding-backend' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.website) return out_({ ok: true });            // honeypot field: real guests leave it empty
    switch (body.action) {
      case 'rsvp':  return out_(saveRsvp_(body));
      case 'photo': return out_(savePhoto_(body));
      case 'list':  return out_(listRsvps_(body));
      case 'clear': return out_(clearRsvps_(body));
      case 'setYoutube': return out_(setYoutube_(body));
      default:      return out_({ ok: false, error: 'Unknown action' });
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
  const row = [
    Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'),
    name, clean_(b.phone, 30), clean_(b.email, 120), attending, events, guests,
    clean_(b.message, 600), b.lang === 'ta' ? 'Tamil' : 'English'
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = rsvpSheet_();
    const r = sheet.getLastRow() + 1;
    const range = sheet.getRange(r, 1, 1, HEADERS.length);
    range.setNumberFormats([HEADERS.map(function (_, i) { return TEXT_COLS.indexOf(i + 1) > -1 ? '@' : '0'; })]);
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function rsvpSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = RSVP_TAB ? (ss.getSheetByName(RSVP_TAB) || ss.insertSheet(RSVP_TAB)) : ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#F6E3BC');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HEADERS.length, 150);
    sheet.setColumnWidth(8, 320);
  }
  return sheet;
}

/* ---------- Photos ---------- */
function savePhoto_(b) {
  const mime = String(b.mime || '').toLowerCase();
  const ext = PHOTO_TYPES[mime];
  if (!ext || !b.data) return { ok: false, error: 'Unsupported file' };
  const bytes = Utilities.base64Decode(b.data);
  if (bytes.length > MAX_PHOTO_BYTES) return { ok: false, error: 'File too large' };

  const who = clean_(b.uploader, 60);
  if (!who) return { ok: false, error: 'Name is required' };
  const slug = who.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'guest';
  const stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
  const fileName = stamp + '_' + slug + '_' + Math.floor(Math.random() * 9000 + 1000) + '.' + ext;

  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(Utilities.newBlob(bytes, mime, fileName));
  file.setDescription('Uploaded by: ' + who);
  return { ok: true };
}

/* ---------- Admin: list RSVPs ---------- */
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

  const sheet = rsvpSheet_();
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getDisplayValues() : [];
  return {
    ok: true,
    headers: HEADERS,
    rows: rows,
    youtubeUrl: PropertiesService.getScriptProperties().getProperty('YOUTUBE_URL') || '',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit',
    folderUrl: 'https://drive.google.com/drive/folders/' + DRIVE_FOLDER_ID
  };
}

/* ---------- Admin: delete every RSVP row (the header row stays) ---------- */
function clearRsvps_(b) {
  const bad = authError_(b);
  if (bad) return bad;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = rsvpSheet_();
    const n = sheet.getLastRow() - 1;
    if (n > 0) sheet.deleteRows(2, n);
    return { ok: true, deleted: Math.max(n, 0) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Admin: set the YouTube live link shown on the site ---------- */
function setYoutube_(b) {
  const bad = authError_(b);
  if (bad) return bad;
  const url = clean_(b.url, 300);
  if (url && !/^https:\/\/(www\.|m\.)?(youtube\.com|youtube-nocookie\.com|youtu\.be)\//i.test(url)) {
    return { ok: false, error: 'Please paste a youtube.com or youtu.be link.' };
  }
  const props = PropertiesService.getScriptProperties();
  if (url) props.setProperty('YOUTUBE_URL', url); else props.deleteProperty('YOUTUBE_URL');
  return { ok: true, youtubeUrl: url };
}

/* ---------- helpers ---------- */
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor to authorise access and check both IDs. */
function setup() {
  const sheet = rsvpSheet_();
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) {
    throw new Error('Add the ADMIN_KEY script property first (step 4).');
  }
  console.log('Setup OK. Sheet tab "' + sheet.getName() + '", Drive folder "' + folder.getName() + '".');
}
