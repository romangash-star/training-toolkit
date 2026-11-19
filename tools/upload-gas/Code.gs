/**
 * Receives presentations from romanai.net/data/upload/ (day 3 of "קבלת החלטות מבוססת נתונים"),
 * saves each one to the owner's Drive in a sub-folder per local authority, and emails the owner a link.
 * Runs as the owner; anyone can post (participants have no Google sign-in).
 */
var FOLDER_NAME = 'מצגות משתתפים - קבלת החלטות מבוססת נתונים';
var MAX_BYTES = 35 * 1024 * 1024;
var ALLOWED = ['pptx', 'ppt', 'pdf', 'odp', 'key', 'potx', 'ppsx'];

/** Run once from the editor: approves the permissions, creates the folder, sends a test email. */
function setup() {
  var folder = getFolder_();
  MailApp.sendEmail({
    to: ownerEmail_(),
    subject: 'העלאת מצגות מוכנה ✓',
    htmlBody: '<div dir="rtl">ההעלאה מהאתר מחוברת. מצגות שיועלו יישמרו בתיקייה: ' +
      '<a href="' + folder.getUrl() + '">' + FOLDER_NAME + '</a></div>'
  });
  Logger.log('ok: ' + folder.getUrl());
}

function doGet() {
  return json_({ ok: true });
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (b.website) return json_({ ok: true }); // honeypot: bots fill hidden fields
    if (b.ping) return json_({ ok: true, pong: true }); // connection check from /check/ and the upload page
    if (b.kind === 'dataai') return dataai_(b); // text-only form from /data-ai/my-file/
    var name = clean_(b.name, 80), org = clean_(b.org, 80), dept = clean_(b.dept, 80), note = clean_(b.note, 500);
    var fileName = clean_(b.fileName, 150);
    var ext = (fileName.split('.').pop() || '').toLowerCase();
    if (!name || !b.data) return json_({ ok: false, error: 'missing' });
    if (ALLOWED.indexOf(ext) < 0) return json_({ ok: false, error: 'type' });
    var bytes = Utilities.base64Decode(b.data);
    if (bytes.length > MAX_BYTES) return json_({ ok: false, error: 'size' });

    var stamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd HH-mm');
    var saveAs = [stamp, org, name].filter(String).join(' - ') + '.' + ext;
    var blob = Utilities.newBlob(bytes, b.mimeType || 'application/octet-stream', saveAs);
    var orgFolder = getOrgFolder_(org || 'ללא רשות');
    var file = orgFolder.createFile(blob);
    file.setDescription('הועלה מהאתר. שם: ' + name + ' | רשות: ' + org + ' | מחלקה: ' + dept + (note ? ' | הערה: ' + note : ''));

    var mb = (bytes.length / 1048576).toFixed(1);
    MailApp.sendEmail({
      to: ownerEmail_(),
      subject: 'מצגת חדשה: ' + name + (org ? ' (' + org + ')' : ''),
      htmlBody: '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;line-height:1.7">' +
        '<b>הועלתה מצגת חדשה מהאתר</b><br>' +
        'שם: ' + esc_(name) + '<br>' +
        'רשות: ' + esc_(org || '—') + '<br>' +
        'מחלקה: ' + esc_(dept || '—') + '<br>' +
        (note ? 'הערה: ' + esc_(note) + '<br>' : '') +
        'קובץ: ' + esc_(fileName) + ' (' + mb + 'MB)<br><br>' +
        '<a href="' + file.getUrl() + '">פתיחת המצגת ב-Drive</a> · ' +
        '<a href="' + orgFolder.getUrl() + '">התיקייה של ' + esc_(org || 'ללא רשות') + '</a> · ' +
        '<a href="' + getFolder_().getUrl() + '">כל הרשויות</a></div>'
    });
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

/** "הקובץ שלי בשבוע הבא" + feedback, from /data-ai/my-file/. Text only, no file: one email to the owner per answer. */
function dataai_(b) {
  var f = function (k, n) { return clean_(b[k], n || 300); };
  if (!f('name', 80)) return json_({ ok: false, error: 'missing' });
  var rows = [['שם', f('name', 80)], ['רשות', f('org', 80)], ['מחלקה', f('dept', 80)],
    ['קובץ התרגול', f('practice', 60)], ['רמה', f('level', 20)], ['כלי', f('tool', 60)],
    ['הקובץ מהעבודה', f('file', 600)], ['הסבב שאנסה עליו', f('round', 60)],
    ['עד כמה היום היה שימושי (1-5)', f('rate', 2)], ['מה היה הכי שימושי', f('good', 800)], ['מה היה חסר או קשה', f('bad', 800)]];
  MailApp.sendEmail({
    to: ownerEmail_(),
    subject: 'נתונים ו-AI, הקובץ שלי: ' + f('name', 80) + (f('org', 80) ? ' (' + f('org', 80) + ')' : ''),
    htmlBody: '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;line-height:1.7">' +
      rows.map(function (r) { return '<b>' + r[0] + ':</b> ' + esc_(r[1] || '—'); }).join('<br>') + '</div>'
  });
  return json_({ ok: true, kind: 'dataai' });
}

function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  }
  var folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}
/** One folder per authority inside the main folder. A lock keeps two uploads from the same
    authority, arriving together, from creating the folder twice. */
function getOrgFolder_(org) {
  var main = getFolder_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var it = main.getFoldersByName(org);
    return it.hasNext() ? it.next() : main.createFolder(org);
  } finally {
    lock.releaseLock();
  }
}
function ownerEmail_() { return Session.getEffectiveUser().getEmail(); }
function clean_(s, n) { return String(s || '').replace(/[\\\/:*?"<>|\r\n]+/g, ' ').trim().slice(0, n); }
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
