/** @OnlyCurrentDoc */
var ADMIN_PHONE = '79990000000';
var ADMIN_PASSWORD = 'ЗАМЕНИТЕ';
var ADMIN_NAME = 'Администратор';
var EDIT_PAST_MONTHS = true;
var SESSION_SEC = 21600;

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); }
  return s;
}
var USERS_H = ['phone', 'name', 'role', 'salt', 'hash', 'active'];
var ROWS_H = ['phone', 'month', 'project', 'work', 'amount', 'updatedAt'];
var PAY_H = ['id', 'phone', 'month', 'amount', 'date', 'note'];
function paysSheet() { return sh('Выплаты', PAY_H); }
function users() { return sh('Сотрудники', USERS_H); }
function rows() { return sh('Записи', ROWS_H); }
function ms(v) { return v instanceof Date ? Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM') : String(v); }

function normPhone(p) {
  var d = String(p || '').replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '8') d = '7' + d.slice(1);
  if (d.length === 10) d = '7' + d;
  return d;
}
function hash(salt, pass) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ((x & 255) + 256).toString(16).slice(1); }).join('');
}
function findUser(phone) {
  var s = users(), n = s.getLastRow();
  if (n < 2) return null;
  var v = s.getRange(2, 1, n - 1, 6).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === phone) return { row: i + 2, phone: String(v[i][0]), name: v[i][1], role: v[i][2], salt: v[i][3], hash: v[i][4], active: v[i][5] !== false && v[i][5] !== 'FALSE' };
  return null;
}
function putUser(phone, name, role, pass) {
  CacheService.getScriptCache().remove('u_' + phone);
  var salt = Utilities.getUuid(), h = hash(salt, pass), u = findUser(phone), s = users();
  if (u) s.getRange(u.row, 1, 1, 6).setValues([[phone, name || u.name, role || u.role, salt, h, true]]);
  else { s.appendRow([phone, name, role || 'user', salt, h, true]); s.getRange(s.getLastRow(), 1).setNumberFormat('@').setValue(phone); }
}

function setupAdmin() {
  paysSheet();
  if (ADMIN_PASSWORD.indexOf('ЗАМЕНИТЕ') === 0 || ADMIN_PASSWORD.length < 6) throw new Error('Впишите пароль админа (от 6 знаков)');
  users(); rows();
  putUser(normPhone(ADMIN_PHONE), ADMIN_NAME, 'admin', ADMIN_PASSWORD);
}

function doGet() { return out({ ok: true, app: 'kabinet', adminSet: users().getLastRow() > 1 }); }

function doPost(e) {
  var b; try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ error: 'bad' }); }
  var W = { adminAddPay: 1, adminDelPay: 1, saveMonth: 1, changePassword: 1, adminSaveUser: 1 }, lock = null;
  if (W[b.action]) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
  try {
    if (b.action === 'login') return out(login(b));
    var me = auth(b.token);
    if (!me) return out({ error: 'auth' });
    switch (b.action) {
      case 'myMonths': return out({ data: myMonths(me.phone), pays: myPays(me.phone) });
      case 'saveMonth': return out(saveMonth(me, b));
      case 'changePassword': return out(changePassword(me, b));
      case 'adminMonth': return admin(me) || out(adminMonth(b.month));
      case 'adminUsers': return admin(me) || out({ users: listUsers() });
      case 'adminAddPay': return admin(me) || out(adminAddPay(b));
      case 'adminDelPay': return admin(me) || out(adminDelPay(b));
      case 'adminSaveUser': return admin(me) || out(adminSaveUser(b));
      case 'adminEmployeeMonths': return admin(me) || out({ data: myMonths(normPhone(b.phone)), pays: myPays(normPhone(b.phone)) });
    }
    return out({ error: 'bad action' });
  } finally { if (lock) lock.releaseLock(); }
}
function admin(me) { return me.role === 'admin' ? null : out({ error: 'forbidden' }); }

function login(b) {
  var phone = normPhone(b.phone), c = CacheService.getScriptCache(), fk = 'f_' + phone;
  if (Number(c.get(fk) || 0) >= 5) return { error: 'locked' };
  var u = findUser(phone);
  if (!u || !u.active || hash(u.salt, String(b.password || '')) !== u.hash) {
    c.put(fk, String(Number(c.get(fk) || 0) + 1), 900);
    return { error: 'wrong' };
  }
  c.remove(fk);
  var t = Utilities.getUuid() + Utilities.getUuid();
  c.put('t_' + t, phone, SESSION_SEC);
  return { token: t, name: u.name, role: u.role, phone: phone, editPast: EDIT_PAST_MONTHS };
}
function auth(t) {
  if (!t) return null;
  var c = CacheService.getScriptCache(), phone = c.get('t_' + t);
  if (!phone) return null;
  var j = c.get('u_' + phone);
  if (j) return JSON.parse(j);
  var u = findUser(phone);
  if (!u || !u.active) return null;
  var m = { phone: u.phone, name: u.name, role: u.role };
  c.put('u_' + phone, JSON.stringify(m), SESSION_SEC);
  return m;
}

function myMonths(phone) {
  var s = rows(), n = s.getLastRow(), res = {};
  if (n < 2) return res;
  s.getRange(2, 1, n - 1, 5).getValues().forEach(function (r) {
    if (String(r[0]) !== phone) return;
    (res[ms(r[1])] = res[ms(r[1])] || []).push({ p: r[2], w: r[3], a: Number(r[4]) || 0 });
  });
  return res;
}
function allPays() {
  var s = paysSheet(), n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, 6).getValues().map(function (r) { return { id: String(r[0]), phone: String(r[1]), month: ms(r[2]), a: Number(r[3]) || 0, d: String(r[4]), n: String(r[5]) }; });
}
function myPays(phone) {
  var res = {};
  allPays().forEach(function (p) { if (p.phone === phone) (res[p.month] = res[p.month] || []).push({ a: p.a, d: p.d, n: p.n }); });
  return res;
}
function adminAddPay(b) {
  var phone = normPhone(b.phone), month = String(b.month || ''), a = Number(String(b.amount).replace(',', '.'));
  if (!findUser(phone) || !/^\d{4}-\d{2}$/.test(month) || !(a > 0)) return { error: 'bad' };
  var s = paysSheet(), r = s.getLastRow() + 1;
  s.getRange(r, 1, 1, 6).setNumberFormat('@').setValues([[Utilities.getUuid(), phone, month, a, String(b.date || '').slice(0, 10), String(b.note || '').slice(0, 200)]]);
  return { ok: true };
}
function adminDelPay(b) {
  var s = paysSheet(), n = s.getLastRow();
  if (n < 2) return { ok: true };
  var v = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(b.id)) { s.deleteRow(i + 2); break; }
  return { ok: true };
}
function curMonth() { return Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM'); }
function saveMonth(me, b) {
  var phone = me.phone, month = String(b.month || '');
  if (me.role === 'admin' && b.phone) phone = normPhone(b.phone);
  if (!/^\d{4}-\d{2}$/.test(month)) return { error: 'bad month' };
  if (me.role !== 'admin' && !EDIT_PAST_MONTHS && month < curMonth()) return { error: 'past locked' };
  var s = rows(), n = s.getLastRow();
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 2).getValues();
    for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === phone && ms(v[i][1]) === month) s.deleteRow(i + 2);
  }
  var now = Date.now(), add = [];
  (b.rows || []).slice(0, 200).forEach(function (r) {
    var a = Number(String(r.a).replace(',', '.')) || 0;
    if (!r.p && !r.w && !a) return;
    add.push([phone, month, String(r.p || '').slice(0, 200), String(r.w || '').slice(0, 200), a, now]);
  });
  if (add.length) {
    var start = s.getLastRow() + 1;
    s.getRange(start, 1, add.length, 2).setNumberFormat('@');
    s.getRange(start, 1, add.length, 6).setValues(add);
  }
  return { saved: add.length };
}
function changePassword(me, b) {
  if (!b.newPassword || String(b.newPassword).length < 6) return { error: 'short' };
  var cu = findUser(me.phone);
  if (hash(cu.salt, String(b.oldPassword || '')) !== cu.hash) return { error: 'wrong' };
  putUser(me.phone, me.name, me.role, String(b.newPassword));
  return { ok: true };
}

function listUsers() {
  var s = users(), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 6).getValues().map(function (r) { return { phone: String(r[0]), name: r[1], role: r[2], active: r[5] !== false && r[5] !== 'FALSE' }; });
}
function adminMonth(month) {
  var us = listUsers(), by = {};
  us.forEach(function (u) { by[u.phone] = { phone: u.phone, name: u.name, active: u.active, rows: [], total: 0 }; });
  var s = rows(), n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, 5).getValues().forEach(function (r) {
    if (ms(r[1]) !== month) return;
    var p = String(r[0]); if (!by[p]) return;
    var a = Number(r[4]) || 0;
    by[p].rows.push({ p: r[2], w: r[3], a: a }); by[p].total += a;
  });
  Object.keys(by).forEach(function (k) { by[k].pays = []; by[k].paid = 0; });
  allPays().forEach(function (p) {
    if (p.month !== month || !by[p.phone]) return;
    by[p.phone].pays.push({ id: p.id, a: p.a, d: p.d, n: p.n }); by[p.phone].paid += p.a;
  });
  return { month: month, employees: Object.keys(by).map(function (k) { return by[k]; }) };
}
function adminSaveUser(b) {
  var phone = normPhone(b.phone);
  if (phone.length !== 11) return { error: 'phone' };
  var u = findUser(phone);
  if (b.active === false || b.active === true) {
    if (!u) return { error: 'nouser' };
    users().getRange(u.row, 6).setValue(b.active);
    CacheService.getScriptCache().remove('u_' + phone);
    return { ok: true };
  }
  if (!u && (!b.password || String(b.password).length < 6)) return { error: 'short' };
  if (u && !b.password) { users().getRange(u.row, 2).setValue(b.name || u.name); return { ok: true }; }
  if (String(b.password).length < 6) return { error: 'short' };
  putUser(phone, b.name || (u && u.name) || '', (u && u.role) || 'user', String(b.password));
  return { ok: true };
}
