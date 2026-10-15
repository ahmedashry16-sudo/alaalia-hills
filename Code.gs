/**
 * العالية هيلز — خادم المزامنة مع تسجيل الدخول والصلاحيات
 * ------------------------------------------------------------
 * - الدخول برمز تحقق (6 أرقام) يُرسل إلى الإيميل، والجلسة صالحة 30 يوماً.
 * - أي إيميل جديد يُسجَّل كـ "طلب دخول" ويصل للمدير إشعار بالإيميل.
 * - المدير يوافق من داخل الموقع (أو من تبويب Users في الشيت) بصلاحية:
 *      view  = مستعرض فقط      edit = مستعرض + تعديل الحالات
 *      admin = تعديل + إدارة المستخدمين     rejected = مرفوض
 * - الحالات تبقى في تبويب «القطع» كما هي، مع عمود جديد «عُدّل بواسطة».
 * - قراءة وكتابة الحالات لا تتم إلا بجلسة صالحة، والتعديل لصلاحية edit/admin فقط.
 *
 * التركيب: الصق هذا الملف بدل الكود القديم، شغّل الدالة setup مرة واحدة،
 * ثم Deploy → Manage deployments → ✏️ → Version: New version → Deploy
 * (بهذا يبقى رابط الـ Web App نفسه ولا تحتاج لتغييره في الموقع).
 */

// ======================= الإعدادات =======================
const ADMIN_EMAIL = 'ahmedashry16@gmail.com';   // المدير الذي يوافق على الطلبات
const SITE_URL    = 'https://ahmedashry16-sudo.github.io/alaalia-hills/%D8%A7%D9%84%D8%B9%D8%A7%D9%84%D9%8A%D8%A9_%D9%87%D9%8A%D9%84%D8%B2_Desktop.html';
const APP_NAME    = 'العالية هيلز';
const SESSION_DAYS = 30;
const OTP_MINUTES  = 10;
const OTP_MAX_TRIES = 5;

const USERS_SHEET  = 'Users';
const STATUS_SHEET = 'القطع';          // تبويب الحالات الحالي (كما هو)
const PREFIX = 'v2_';
const AR = { available: 'متاحة', sold: 'مباعة', reserved: 'محجوزة' };
const EN = { 'متاحة': 'available', 'مباعة': 'sold', 'محجوزة': 'reserved' };
const VALID_STATUS = ['available', 'sold', 'reserved'];
const ROLES = ['pending', 'view', 'edit', 'admin', 'rejected'];

// ======================= نقاط الدخول =======================
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return out_(route_(p));
}
function doGet(e) {
  // يدعم GET أيضاً (للتجربة من المتصفح). بدون جلسة لن يُرجع أي بيانات.
  return out_(route_((e && e.parameter) || {}));
}

function route_(p) {
  try {
    switch (p.action) {
      case 'start':   return start_(p);
      case 'verify':  return verify_(p);
      case 'logout':  return logout_(p);
      case 'me':      { const u = auth_(p, 'view'); return { success: true, email: u.email, name: u.name, role: u.role }; }
      case 'get':     return getData_(p);
      case 'set':     return setOne_(p);
      case 'bulk':    return setBulk_(p);
      case 'users':   return listUsers_(p);
      case 'setRole': return setRole_(p);
      default:        return { success: false, error: 'unknown_action' };
    }
  } catch (err) {
    const msg = String(err && err.message || err);
    if (msg.indexOf('AUTH:') === 0) return { success: false, auth: msg.slice(5) };
    return { success: false, error: msg };
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ======================= الإعداد لأول مرة =======================
/** شغّلها مرة واحدة من المحرر (ستطلب منك صلاحيات الشيت وإرسال الإيميل). */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let us = ss.getSheetByName(USERS_SHEET);
  if (!us) {
    us = ss.insertSheet(USERS_SHEET);
    us.appendRow(['email', 'name', 'role', 'requested_at', 'approved_at', 'approved_by', 'last_login']);
    us.setFrozenRows(1);
    us.getRange('C2:C').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).build());
  }
  if (!findUserRow_(ADMIN_EMAIL)) {
    us.appendRow([ADMIN_EMAIL, 'المدير', 'admin', new Date(), new Date(), 'setup', '']);
  }
  const st = sheet_(STATUS_SHEET);
  if (String(st.getRange(1, 7).getValue()) === '') st.getRange(1, 7).setValue('عُدّل بواسطة')
    .setFontWeight('bold').setBackground('#1f4e3d').setFontColor('#ffffff');
  cleanupSessions_();
  // يرسل إيميل تجريبي للتأكد من صلاحية الإرسال
  MailApp.sendEmail(ADMIN_EMAIL, APP_NAME + ' — تم تفعيل تسجيل الدخول',
    'تم تجهيز نظام الدخول بنجاح.\n\n' + SITE_URL);
  Logger.log('Setup done');
}

// ======================= تسجيل الدخول =======================
function normEmail_(e) { return String(e || '').trim().toLowerCase(); }
function validEmail_(e) { return /^[^\s@'"<>]+@[^\s@'"<>]+\.[^\s@]{2,}$/.test(e); }

/** الخطوة 1: إدخال الإيميل. */
function start_(p) {
  const email = normEmail_(p.email);
  if (!validEmail_(email)) return { success: false, error: 'invalid_email' };
  const u = getUser_(email);

  if (!u) {
    const name = String(p.name || '').trim().slice(0, 80);
    if (!name) return { success: true, state: 'need_name' };
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      if (!findUserRow_(email)) {
        sheet_(USERS_SHEET).appendRow([email, name, 'pending', new Date(), '', '', '']);
      }
    } finally { lock.releaseLock(); }
    notifyAdmin_(email, name);
    return { success: true, state: 'pending' };
  }
  if (u.role === 'pending')  return { success: true, state: 'pending' };
  if (u.role === 'rejected') return { success: true, state: 'rejected' };

  // مستخدم معتمد → إرسال رمز
  const cache = CacheService.getScriptCache();
  if (cache.get('otpRL_' + email)) return { success: true, state: 'code_sent', throttled: true };
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('otp_' + email, JSON.stringify({ h: hash_(code + email), n: 0 }), OTP_MINUTES * 60);
  cache.put('otpRL_' + email, '1', 60);
  MailApp.sendEmail({
    to: email,
    subject: 'رمز الدخول إلى ' + APP_NAME + ': ' + code,
    htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px">' +
      'رمز الدخول إلى <b>' + APP_NAME + '</b>:' +
      '<div style="font-size:30px;letter-spacing:6px;font-weight:bold;margin:14px 0;color:#0d4a52">' + code + '</div>' +
      'صالح لمدة ' + OTP_MINUTES + ' دقائق. إذا لم تطلب هذا الرمز تجاهل الرسالة.</div>'
  });
  return { success: true, state: 'code_sent' };
}

/** الخطوة 2: إدخال الرمز → جلسة. */
function verify_(p) {
  const email = normEmail_(p.email);
  const code = String(p.code || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('otp_' + email);
  if (!raw) return { success: false, error: 'code_expired' };
  const o = JSON.parse(raw);
  if (o.n >= OTP_MAX_TRIES) { cache.remove('otp_' + email); return { success: false, error: 'too_many_tries' }; }
  if (hash_(code + email) !== o.h) {
    o.n++; cache.put('otp_' + email, JSON.stringify(o), OTP_MINUTES * 60);
    return { success: false, error: 'wrong_code', left: OTP_MAX_TRIES - o.n };
  }
  cache.remove('otp_' + email);
  const u = getUser_(email);
  if (!u || ['view', 'edit', 'admin'].indexOf(u.role) < 0) return { success: false, error: 'not_allowed' };

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('S_' + hash_(token),
    JSON.stringify({ e: email, x: Date.now() + SESSION_DAYS * 864e5 }));
  updateUserCell_(email, 7, new Date());           // last_login
  return { success: true, token: token, email: email, name: u.name, role: u.role };
}

function logout_(p) {
  if (p.token) PropertiesService.getScriptProperties().deleteProperty('S_' + hash_(String(p.token)));
  return { success: true };
}

/** يتحقق من الجلسة ويقرأ الصلاحية الحالية من الشيت (أي تغيير يطبّق فوراً). */
function auth_(p, need) {
  const token = String(p.token || '');
  if (!token) throw new Error('AUTH:login');
  const props = PropertiesService.getScriptProperties();
  const key = 'S_' + hash_(token);
  const raw = props.getProperty(key);
  if (!raw) throw new Error('AUTH:login');
  const s = JSON.parse(raw);
  if (s.x < Date.now()) { props.deleteProperty(key); throw new Error('AUTH:login'); }
  const u = getUser_(s.e);
  if (!u || ['view', 'edit', 'admin'].indexOf(u.role) < 0) { props.deleteProperty(key); throw new Error('AUTH:revoked'); }
  const rank = { view: 1, edit: 2, admin: 3 };
  if (rank[u.role] < rank[need]) throw new Error('AUTH:forbidden');
  return u;
}

function cleanupSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => {
    if (k.indexOf('S_') !== 0) return;
    try { if (JSON.parse(all[k]).x < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

// ======================= البيانات =======================
function toEn_(v) {
  v = String(v || '').trim();
  if (EN[v]) return EN[v];
  if (AR[v.toLowerCase()]) return v.toLowerCase();
  return 'available';
}
/** يقرأ تبويب «القطع»: البلوك | رقم القطعة | الاستخدام | المساحة | الحالة | آخر تحديث | عُدّل بواسطة */
function readStatuses_() {
  const sh = sheet_(STATUS_SHEET);
  const n = sh.getLastRow();
  const map = {}, rowOf = {};
  if (n > 1) sh.getRange(2, 1, n - 1, 5).getValues().forEach((r, i) => {
    if (r[0] === '' || r[1] === '') return;
    const k = PREFIX + String(r[0]).trim() + '_' + String(r[1]).trim();
    map[k] = toEn_(r[4]); rowOf[k] = i + 2;
  });
  return { sh: sh, map: map, rowOf: rowOf };
}

function getData_(p) {
  const u = auth_(p, 'view');
  return { success: true, data: readStatuses_().map, role: u.role, email: u.email, name: u.name };
}

function writeStatuses_(updates, who) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = readStatuses_(), now = new Date(), append = [];
    Object.keys(updates).forEach(k => {
      const v = String(updates[k]);
      const m = String(k).match(/^v2_(\w+)_(\w+)$/);
      if (!m || VALID_STATUS.indexOf(v) < 0) return;
      if (d.rowOf[k]) {
        d.sh.getRange(d.rowOf[k], 5, 1, 2).setValues([[AR[v], now]]);
        d.sh.getRange(d.rowOf[k], 7).setValue(who);
      } else append.push([m[1], m[2], '', '', AR[v], now, who]);
    });
    if (append.length) d.sh.getRange(d.sh.getLastRow() + 1, 1, append.length, 7).setValues(append);
  } finally { lock.releaseLock(); }
}

function setOne_(p) {
  const u = auth_(p, 'edit');
  const o = {}; o[String(p.key)] = String(p.status);
  writeStatuses_(o, u.email);
  return { success: true };
}

function setBulk_(p) {
  const u = auth_(p, 'edit');
  let upd = p.updates;
  if (typeof upd === 'string') { try { upd = JSON.parse(upd); } catch (e) { upd = {}; } }
  writeStatuses_(upd || {}, u.email);
  return { success: true };
}

// ======================= إدارة المستخدمين =======================
function listUsers_(p) {
  auth_(p, 'admin');
  const sh = sheet_(USERS_SHEET), n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, 7).getValues() : [];
  return {
    success: true,
    users: rows.filter(r => r[0]).map(r => ({
      email: normEmail_(r[0]), name: String(r[1] || ''), role: String(r[2] || 'pending'),
      requested: r[3] ? new Date(r[3]).toISOString() : '', lastLogin: r[6] ? new Date(r[6]).toISOString() : ''
    }))
  };
}

function setRole_(p) {
  const admin = auth_(p, 'admin');
  const email = normEmail_(p.email), role = String(p.role || '');
  if (ROLES.indexOf(role) < 0 && role !== 'remove') return { success: false, error: 'bad_role' };
  if (email === normEmail_(ADMIN_EMAIL)) return { success: false, error: 'cannot_change_owner' };
  const row = findUserRow_(email);
  if (!row) return { success: false, error: 'no_user' };
  const sh = sheet_(USERS_SHEET);
  const prev = String(sh.getRange(row, 3).getValue());
  if (role === 'remove') { sh.deleteRow(row); return { success: true }; }
  sh.getRange(row, 3).setValue(role);
  if (role === 'view' || role === 'edit' || role === 'admin') {
    sh.getRange(row, 5, 1, 2).setValues([[new Date(), admin.email]]);
    if (prev === 'pending') {
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'تمت الموافقة على دخولك إلى ' + APP_NAME,
          htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px">تمت الموافقة على طلبك بصلاحية <b>' +
            (role === 'view' ? 'مستعرض' : role === 'edit' ? 'تعديل' : 'مدير') + '</b>.<br><br>' +
            '<a href="' + SITE_URL + '">افتح الخريطة</a> وأدخل إيميلك لتصلك رسالة برمز الدخول.</div>'
        });
      } catch (e) {}
    }
  }
  return { success: true };
}

function notifyAdmin_(email, name) {
  const cache = CacheService.getScriptCache();
  if (cache.get('notif_' + email)) return;
  cache.put('notif_' + email, '1', 6 * 3600);
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'طلب دخول جديد — ' + APP_NAME + ': ' + name,
      htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px">' +
        'طلب دخول جديد:<br><b>' + escHtml_(name) + '</b> — ' + escHtml_(email) + '<br><br>' +
        'للموافقة: <a href="' + SITE_URL + '">افتح الموقع</a> ← زر «المستخدمين» واختر مستعرض أو تعديل.<br>' +
        'أو اكتب الصلاحية (view / edit / rejected) في تبويب Users في الشيت.</div>'
    });
  } catch (e) {}
}

// ======================= أدوات =======================
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('شغّل الدالة setup أولاً — تبويب ' + name + ' غير موجود');
  return sh;
}
function findUserRow_(email) {
  const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;
  const e = normEmail_(email);
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (normEmail_(col[i][0]) === e) return i + 2;
  return 0;
}
function getUser_(email) {
  const row = findUserRow_(email);
  if (!row) return null;
  const r = sheet_(USERS_SHEET).getRange(row, 1, 1, 3).getValues()[0];
  let role = String(r[2] || 'pending').trim().toLowerCase();
  if (ROLES.indexOf(role) < 0) role = 'pending';
  if (normEmail_(email) === normEmail_(ADMIN_EMAIL)) role = 'admin';
  return { email: normEmail_(r[0]), name: String(r[1] || ''), role: role, row: row };
}
function updateUserCell_(email, col, val) {
  const row = findUserRow_(email); if (row) sheet_(USERS_SHEET).getRange(row, col).setValue(val);
}
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function escHtml_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
