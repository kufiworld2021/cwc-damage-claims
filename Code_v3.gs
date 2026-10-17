/**
 * عالم الكوفي — خادم بلاغات الضرر (Google Apps Script)
 * ------------------------------------------------------------
 * 1) افتح Google Sheet جديد ← Extensions ← Apps Script ← امسح الموجود والصق هذا الملف
 * 2) اختر الدالة setup من القائمة فوق واضغط Run مرة وحدة ووافق على الصلاحيات
 * 3) Deploy ← New deployment ← Web app
 *      Execute as: Me   ·   Who has access: Anyone
 * 4) انسخ رابط الـ Web app (ينتهي بـ /exec) وأرسله — يُلصق في API_URL داخل صفحة البلاغ
 * ------------------------------------------------------------
 */

var CFG = {
  SHEET: 'البلاغات',
  ALERT_TO: 'Ammar1111161@gmail.com',            // إيميلات إضافية للتنبيه (افصل بفاصلة) — صاحب الشيت يوصله دائمًا
  FOLDER_NAME: 'بلاغات عالم الكوفي — الصور',
  START_NO: 1001,
  PREFIX: 'CWC-',
  TZ: 'Asia/Riyadh',
  MIN_PH: 2,
  MAX_PH: 8
};

var COLS = ['رقم البلاغ', 'تاريخ البلاغ', 'رقم الطلب', 'الجوال', 'الاسم', 'تاريخ الاستلام',
  'نوع الضرر', 'الوصف', 'عدد الصور', 'مجلد الصور', 'الحالة', 'رسالة للعميل', 'آخر تحديث', 'اللغة', 'نُقل للمركز',
  'تنبيه', 'نص التطبيق (انسخه)'];
// أرقام الأعمدة
var C = { ID: 1, AT: 2, ORDER: 3, PHONE: 4, NAME: 5, DELIV: 6, TYPES: 7, DESC: 8, NPH: 9, FOLDER: 10,
  STATUS: 11, NOTE: 12, UPD: 13, LANG: 14, MOVED: 15, FLAG: 16, APP: 17 };
var SHARE_PHOTOS = true; // مجلد الصور يفتح لمن عنده الرابط — عشان الموظف يشوفه من مركز البلاغات

// الحالة كما تكتبونها في الشيت ← الكود اللي تفهمه صفحة التتبع
var STATUS = {
  'جديد': 'new',
  'قيد المراجعة': 'review',
  'نحتاج معلومات': 'info',
  'تم اتخاذ القرار': 'decided',
  'تم الشحن / التحويل': 'shipped',
  'مغلق': 'closed',
  'غير مقبول': 'rejected'
};

/* ================= الإعداد (مرة وحدة) ================= */
function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  var sh = ss.getSheetByName(CFG.SHEET) || ss.insertSheet(CFG.SHEET, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS])
    .setFontWeight('bold').setBackground('#53230E').setFontColor('#F6E7D2').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  sh.setColumnWidths(1, COLS.length, 120);
  sh.setColumnWidth(C.TYPES, 200); sh.setColumnWidth(C.DESC, 260); sh.setColumnWidth(C.NOTE, 260);

  var p = PropertiesService.getScriptProperties();
  p.setProperty('SHEET_ID', ss.getId());
  if (!p.getProperty('FOLDER')) p.setProperty('FOLDER', DriveApp.createFolder(CFG.FOLDER_NAME).getId());
  if (!p.getProperty('NEXT')) p.setProperty('NEXT', String(CFG.START_NO));
  MailApp.getRemainingDailyQuota(); // يطلب صلاحية الإيميل
  Logger.log('تم الإعداد ✓ — الآن: Deploy ← New deployment ← Web app');
}

/* ================= استقبال بلاغ جديد ================= */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents || '{}');
    if (d.a !== 'new') return out_({ ok: false, err: 'bad_action' });

    var order = digits_(d.order).slice(0, 20);
    var phone = digits_(d.phone).slice(0, 15);
    var photos = (d.photos || []).slice(0, CFG.MAX_PH);
    if (order.length < 4 || phone.length < 9 || !d.deliv) return out_({ ok: false, err: 'fields' });
    if (photos.length < CFG.MIN_PH) return out_({ ok: false, err: 'photos' });

    var p = PropertiesService.getScriptProperties();
    lock.waitLock(20000);
    var no = Number(p.getProperty('NEXT') || CFG.START_NO);
    p.setProperty('NEXT', String(no + 1));
    lock.releaseLock();

    var id = CFG.PREFIX + no;
    var now = new Date();
    var name = clean_(d.name, 60);
    var desc = clean_(d.desc, 600);
    var types = (d.typesAr || []).map(function (s) { return clean_(s, 40); }).join('، ');
    var deliv = parseDate_(d.deliv);

    // الصور ← مجلد خاص بالبلاغ
    var folder = DriveApp.getFolderById(p.getProperty('FOLDER')).createFolder(id + ' · طلب ' + order);
    var blobs = photos.map(function (b64, i) {
      return Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', id + '-' + (i + 1) + '.jpg');
    });
    blobs.forEach(function (b) { folder.createFile(b); });
    if (SHARE_PHOTOS) { try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {} }
    var url = folder.getUrl();

    var sh = sheet_();

    // تنبيه: نفس الطلب له بلاغ سابق؟
    var flag = '', last = sh.getLastRow();
    if (last > 1) {
      var dup = sh.getRange(2, C.ORDER, last - 1, 1).getValues()
        .filter(function (r) { return digits_(r[0]) === order; }).length;
      if (dup) flag = 'للطلب ' + dup + ' بلاغ سابق';
    }

    // نص جاهز للصق في مركز البلاغات (خانة «بلاغ وارد من عميل»)
    var app = 'KUFI' + JSON.stringify({ id: id, order: order, name: name, phone: phone,
      types: (d.types || []).map(function (s) { return clean_(s, 10); }),
      deliv: String(d.deliv || '').slice(0, 10), desc: desc, n: blobs.length, drive: url, flag: flag });

    // صف جديد في الشيت
    sh.appendRow([id, now, "'" + order, "'" + phone, name, deliv, types, desc, blobs.length,
      url, 'جديد', '', now, d.lang === 'en' ? 'EN' : 'AR', '', flag, app]);
    formatRow_(sh, sh.getLastRow());

    // إيميل التنبيه + نص التطبيق
    sendAlert_({ id: id, at: now, order: order, phone: phone, name: name, deliv: deliv, types: types,
      desc: desc, n: blobs.length, folder: url, sheet: sh.getParent().getUrl(), flag: flag, app: app }, blobs);

    return out_({ ok: true, id: id });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ================= التتبع ================= */
function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.a !== 'track') return out_({ ok: true, service: 'cwc-claims' });
  var id = String(q.id || '').toUpperCase().replace(/\s/g, '');
  var p4 = digits_(q.p);
  if (!id || p4.length !== 4) return out_({ ok: false });

  var sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return out_({ ok: false });
  var v = sh.getRange(2, 1, last - 1, C.LANG).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    var r = v[i];
    if (String(r[C.ID - 1]).toUpperCase() !== id) continue;
    if (digits_(r[C.PHONE - 1]).slice(-4) !== p4) return out_({ ok: false });
    return out_({
      ok: true,
      id: r[C.ID - 1],
      order: String(r[C.ORDER - 1]),
      code: STATUS[String(r[C.STATUS - 1]).trim()] || 'new',
      note: String(r[C.NOTE - 1] || ''),
      at: iso_(r[C.AT - 1]),
      upd: iso_(r[C.UPD - 1])
    });
  }
  return out_({ ok: false });
}

/* ================= تحديث «آخر تحديث» تلقائيًا ================= */
function onEdit(e) {
  var r = e.range, sh = r.getSheet();
  if (sh.getName() !== CFG.SHEET || r.getRow() < 2) return;
  if (r.getLastColumn() < C.STATUS || r.getColumn() > C.NOTE) return; // فقط الحالة أو الرسالة
  sh.getRange(r.getRow(), C.UPD, r.getNumRows(), 1).setValue(new Date());
}

/* ================= مساعدات ================= */
function sendAlert_(o, blobs) {
  var txt = appText_(o);
  var to = [Session.getEffectiveUser().getEmail()].concat(String(CFG.ALERT_TO || '').split(','))
    .map(function (x) { return x.trim(); }).filter(String).join(',');
  var html = '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:14px;color:#1E120B">'
    + '<div style="background:#53230E;color:#F6E7D2;padding:12px 16px;border-radius:10px;font-size:16px;font-weight:bold">'
    + 'بلاغ ضرر جديد — ' + esc_(o.id) + '</div>'
    + (o.flag ? '<p style="background:#FDECEA;color:#B42318;border-radius:10px;padding:10px 14px;font-weight:bold">⚠️ ' + esc_(o.flag) + '</p>' : '')
    + '<p style="margin:14px 0 6px;color:#7D6A5C">نص التطبيق — انسخه كاملًا والصقه في مركز البلاغات ← تسجيل بلاغ ← «بلاغ وارد من عميل»:</p>'
    + '<div dir="ltr" style="background:#1E120B;color:#F6E7D2;border-radius:10px;padding:12px 14px;font-family:Consolas,monospace;font-size:12.5px;word-break:break-all;user-select:all">'
    + esc_(o.app) + '</div>'
    + '<p style="margin:16px 0 6px;color:#7D6A5C">تفاصيل البلاغ:</p>'
    + '<pre style="background:#FBF4EA;border:1px solid #E7D6C0;border-radius:10px;padding:12px 14px;white-space:pre-wrap;font-family:Tahoma,Arial,sans-serif;font-size:14px;line-height:1.8">'
    + esc_(txt) + '</pre>'
    + '<p><a href="' + o.folder + '" style="color:#53230E;font-weight:bold">فتح مجلد الصور</a> · '
    + '<a href="' + o.sheet + '" style="color:#53230E;font-weight:bold">فتح الشيت</a></p>'
    + '<p style="color:#7D6A5C;font-size:12px">الصور مرفقة بهذا الإيميل (' + o.n + ').</p></div>';
  MailApp.sendEmail({
    to: to,
    subject: 'بلاغ ضرر جديد ' + o.id + ' · طلب ' + o.order,
    body: 'نص التطبيق (انسخه والصقه في مركز البلاغات):\n' + o.app + '\n\n' + (o.flag ? '⚠️ ' + o.flag + '\n\n' : '')
      + txt + '\n\nالشيت: ' + o.sheet,
    htmlBody: html,
    attachments: blobs,
    name: 'بلاغات عالم الكوفي'
  });
}

function appText_(o) {
  return [
    'بلاغ ضرر — ' + o.id,
    'تاريخ البلاغ: ' + fmt_(o.at, 'dd/MM/yyyy HH:mm'),
    'رقم الطلب: ' + o.order,
    'العميل: ' + (o.name || '—'),
    'الجوال: ' + o.phone,
    'تاريخ الاستلام: ' + fmt_(o.deliv, 'dd/MM/yyyy'),
    'نوع الضرر: ' + (o.types || '—'),
    'الوصف: ' + (o.desc || '—'),
    'عدد الصور: ' + o.n,
    'مجلد الصور: ' + o.folder
  ].join('\n');
}

function formatRow_(sh, r) {
  sh.getRange(r, C.AT).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(r, C.DELIV).setNumberFormat('dd/mm/yyyy');
  sh.getRange(r, C.UPD).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(r, C.STATUS).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(STATUS), true).build());
  sh.getRange(r, C.MOVED).insertCheckboxes();
}

function sheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  return ss.getSheetByName(CFG.SHEET);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[\u0000-\u001F]+/g, ' ').trim().slice(0, n).replace(/^[=+\-@]/, "'$&"); }
function parseDate_(s) { var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ''; }
function fmt_(d, f) { return d instanceof Date ? Utilities.formatDate(d, CFG.TZ, f) : ''; }
function iso_(d) { return d instanceof Date ? d.toISOString() : ''; }
function esc_(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
