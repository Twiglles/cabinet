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
var USERS_H = ['phone', 'name', 'role', 'salt', 'hash', 'active', 'must'];
var ROWS_H = ['phone', 'month', 'project', 'work', 'amount', 'updatedAt'];
var PAY_H = ['id', 'phone', 'month', 'amount', 'date', 'note'];
function paysSheet() { return sh('Выплаты', PAY_H); }
var SESS_H = ['hash', 'phone', 'exp'];
var REMEMBER_DAYS = 30;
var ROLES = ['user', 'manager', 'head', 'admin'];
var SALES_H = ['phone', 'month', 'json', 'total', 'updatedAt'];
var TIERS = {
  A: [[25000, 5000, 0], [50000, 7500, 3], [75000, 10000, 6], [100000, 13000, 8], [130000, 15000, 9], [160000, 17000, 10], [200000, 25000, 12], [230000, 30000, 13], [270000, 35000, 14], [300000, 40000, 15], [400000, 45000, 16], [500000, 50000, 18]],
  KM: [[900000, 50000, 0], [1000000, 50000, 0.5], [1200000, 50000, 1], [1300000, 50000, 2]]
};
var OKLAD_KM = 50000, KPI_SHARE = 0.4, KPI_W = [0.5, 0.3, 0.2];
var KM_FROM = '2026-08';
function gridFor(role, month) { var g = role === 'head' ? 'KM' : role === 'manager' ? 'A' : null; if (g === 'KM' && month && month < KM_FROM) g = 'A'; return g; }
function snum(v) { return Number(String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.')) || 0; }
function tierFor(rev, list) { for (var i = 0; i < list.length; i++) if (rev <= list[i][0]) return i; return list.length - 1; }
function calcSales(s, grid) {
  var rev = 0; (s.cats || []).forEach(function (c) { rev += snum(c.rev); });
  var adj = snum(s.adj), i = tierFor(rev, TIERS.A), t = TIERS.A[i];
  var base = grid === 'KM' ? OKLAD_KM : t[1], sum = base + rev * t[2] / 100, extra = 0;
  if (grid === 'KM') {
    var co = snum(s.company), cp = TIERS.KM[tierFor(co, TIERS.KM)][2], kp = 0;
    KPI_W.forEach(function (w, k) { kp += w * Math.min(100, Math.max(0, snum((s.kpi || [])[k]))) / 100; });
    extra = co * cp / 100 + OKLAD_KM * KPI_SHARE * kp;
  }
  return { rev: rev, total: Math.round((sum + extra + adj) * 100) / 100 };
}
function salesSheet() { return sh('Продажи', SALES_H); }
function salesMap(phone) {
  var s = salesSheet(), n = s.getLastRow(), res = {};
  if (n < 2) return res;
  s.getRange(2, 1, n - 1, 3).getValues().forEach(function (r) {
    if (String(r[0]) !== phone) return;
    try { res[ms(r[1])] = JSON.parse(r[2]); } catch (e) {}
  });
  return res;
}
function myData(phone, role) {
  var g = gridFor(role), res = { data: myMonths(phone), pays: myPays(phone), grid: g };
  if (g) res.sales = salesMap(phone);
  return res;
}
function saveSales(me, b) {
  var phone = me.phone, role = me.role, month = String(b.month || '');
  if ((me.role === 'admin' || me.role === 'head') && b.phone) { phone = normPhone(b.phone); var tu = findUser(phone); role = tu && tu.role; if (me.role === 'head' && role !== 'manager') return { error: 'forbidden' }; }
  if (!/^\d{4}-\d{2}$/.test(month)) return { error: 'bad month' };
  var g = gridFor(role, month); if (!g) return { error: 'nosales' };
  var x = b.sales || {}, cl = { adj: String(x.adj || '').slice(0, 20), company: String(x.company || '').slice(0, 20), kpi: [0, 1, 2].map(function (k) { return String((x.kpi || [])[k] || '').slice(0, 6); }), cats: [] };
  for (var i = 0; i < 7; i++) { var c = (x.cats || [])[i] || {}; cl.cats.push({ w: [0, 1, 2, 3].map(function (k) { return String((c.w || [])[k] || '').slice(0, 8); }), rev: String(c.rev || '').slice(0, 15) }); }
  var tot = calcSales(cl, g).total, s = salesSheet(), n = s.getLastRow(), row = 0;
  if (n > 1) { var v = s.getRange(2, 1, n - 1, 2).getValues(); for (var j = 0; j < v.length; j++) if (String(v[j][0]) === phone && ms(v[j][1]) === month) { row = j + 2; break; } }
  if (!row) { row = n + 1; s.getRange(row, 1, 1, 3).setNumberFormat('@'); }
  s.getRange(row, 1, 1, 5).setValues([[phone, month, JSON.stringify(cl), tot, Date.now()]]);
  return { saved: 1, total: tot };
}
function sessSheet() { return sh('Сессии', SESS_H); }
function tokHash(t) { return hash('sess', t); }
function delSessions(phone, h, keep) {
  var s = sessSheet(), n = s.getLastRow();
  if (n < 2) return;
  var v = s.getRange(2, 1, n - 1, 3).getValues(), now = Date.now();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) !== keep && (String(v[i][1]) === phone && !h || String(v[i][0]) === h || Number(v[i][2]) < now)) s.deleteRow(i + 2);
}
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
  var v = s.getRange(2, 1, n - 1, 7).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === phone) return { row: i + 2, phone: String(v[i][0]), name: v[i][1], role: v[i][2], salt: v[i][3], hash: v[i][4], active: v[i][5] !== false && v[i][5] !== 'FALSE', must: v[i][6] === true || v[i][6] === 'TRUE' };
  return null;
}
function putUser(phone, name, role, pass, must, keep) {
  CacheService.getScriptCache().remove('u_' + phone);
  delSessions(phone, '', keep);
  var salt = Utilities.getUuid(), h = hash(salt, pass), u = findUser(phone), s = users();
  s.getRange(1, 7).setValue('must');
  s.getRange(2, 3, Math.max(1, s.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).setAllowInvalid(false).build());
  if (u) s.getRange(u.row, 1, 1, 7).setValues([[phone, name || u.name, role || u.role, salt, h, true, !!must]]);
  else { s.appendRow([phone, name, role || 'user', salt, h, true, !!must]); s.getRange(s.getLastRow(), 1).setNumberFormat('@').setValue(phone); }
}

function setupAdmin() {
  paysSheet();
  if (ADMIN_PASSWORD.indexOf('ЗАМЕНИТЕ') === 0 || ADMIN_PASSWORD.length < 6) throw new Error('Впишите пароль админа (от 6 знаков)');
  users(); rows();
  putUser(normPhone(ADMIN_PHONE), ADMIN_NAME, 'admin', ADMIN_PASSWORD, false);
}

function doGet() { return out({ ok: true, app: 'kabinet', adminSet: users().getLastRow() > 1 }); }

function doPost(e) {
  var b; try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ error: 'bad' }); }
  var W = { logout: 1, adminAddPay: 1, adminDelPay: 1, saveMonth: 1, saveSales: 1, changePassword: 1, adminSaveUser: 1 }, lock = null;
  if (W[b.action]) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
  try {
    if (b.action === 'login') return out(login(b));
    var me = auth(b.token);
    if (!me) return out({ error: 'auth' });
    if (me.must && b.action !== 'changePassword' && b.action !== 'logout') return out({ error: 'must' });
    switch (b.action) {
      case 'logout': return out(logout(b));
      case 'myMonths': return out(myData(me.phone, me.role));
      case 'saveMonth': return out(saveMonth(me, b));
      case 'saveSales': return out(saveSales(me, b));
      case 'changePassword': return out(changePassword(me, b));
      case 'headMonth': return head(me) || out(headMonth(b.month));
      case 'headEmployee': return head(me) || out(headEmployee(normPhone(b.phone)));
      case 'adminMonth': return admin(me) || out(adminMonth(b.month));
      case 'adminUsers': return admin(me) || out({ users: listUsers() });
      case 'adminAddPay': return admin(me) || out(adminAddPay(b));
      case 'adminDelPay': return admin(me) || out(adminDelPay(b));
      case 'adminSaveUser': return admin(me) || out(adminSaveUser(b));
      case 'adminEmployeeMonths': return admin(me) || out(myData(normPhone(b.phone), (findUser(normPhone(b.phone)) || {}).role));
    }
    return out({ error: 'bad action' });
  } finally { if (lock) lock.releaseLock(); }
}
function head(me) { return me.role === 'head' ? null : out({ error: 'forbidden' }); }
function headMonth(month) {
  var by = {}, res = [];
  listUsers().forEach(function (u) { if (u.role === 'manager' && u.active) { by[u.phone] = { phone: u.phone, name: u.name, rev: 0, total: 0 }; res.push(by[u.phone]); } });
  var s = salesSheet(), n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, 4).getValues().forEach(function (r) {
    var p = String(r[0]); if (ms(r[1]) !== month || !by[p]) return;
    by[p].total = Number(r[3]) || 0;
  });
  if (n > 1) s.getRange(2, 1, n - 1, 3).getValues().forEach(function (r) {
    var p = String(r[0]); if (ms(r[1]) !== month || !by[p]) return;
    try { by[p].rev = calcSales(JSON.parse(r[2]), 'A').rev; } catch (e) {}
  });
  return { month: month, employees: res };
}
function headEmployee(phone) {
  var u = findUser(phone);
  if (!u || u.role !== 'manager') return { error: 'forbidden' };
  return { data: {}, pays: {}, sales: salesMap(phone), grid: 'A' };
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
  if (b.remember) {
    var ss = sessSheet(), r = ss.getLastRow() + 1;
    ss.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[tokHash(t), phone, String(Date.now() + REMEMBER_DAYS * 86400000)]]);
  }
  var res = { token: t, name: u.name, role: u.role, phone: phone, editPast: EDIT_PAST_MONTHS, must: u.must };
  if (!u.must) { var md = myData(phone, u.role); res.data = md.data; res.pays = md.pays; res.sales = md.sales; res.grid = md.grid; }
  if (u.role === 'admin' && !u.must) { res.sumMonth = curMonth(); res.sum = adminMonth(res.sumMonth); }
  return res;
}
function auth(t) {
  if (!t) return null;
  var c = CacheService.getScriptCache(), phone = c.get('t_' + t);
  if (!phone) {
    var ss = sessSheet(), n = ss.getLastRow(), h = tokHash(t);
    if (n < 2) return null;
    var v = ss.getRange(2, 1, n - 1, 3).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][0]) === h && Number(v[i][2]) > Date.now()) { phone = String(v[i][1]); break; }
    if (!phone) return null;
    c.put('t_' + t, phone, SESSION_SEC);
  }
  var j = c.get('u_' + phone);
  if (j) return JSON.parse(j);
  var u = findUser(phone);
  if (!u || !u.active) return null;
  var m = { phone: u.phone, name: u.name, role: u.role, must: u.must };
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
function logout(b) {
  CacheService.getScriptCache().remove('t_' + b.token);
  delSessions('', tokHash(b.token));
  return { ok: true };
}
function changePassword(me, b) {
  if (!b.newPassword || String(b.newPassword).length < 6) return { error: 'short' };
  var cu = findUser(me.phone);
  if (!cu.must && hash(cu.salt, String(b.oldPassword || '')) !== cu.hash) return { error: 'wrong' };
  putUser(me.phone, me.name, me.role, String(b.newPassword), false, tokHash(b.token));
  return { ok: true };
}

function listUsers() {
  var s = users(), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 6).getValues().map(function (r) { return { phone: String(r[0]), name: r[1], role: r[2], active: r[5] !== false && r[5] !== 'FALSE' }; });
}
function adminMonth(month) {
  var us = listUsers(), by = {};
  us.forEach(function (u) { by[u.phone] = { phone: u.phone, name: u.name, active: u.active, role: u.role, rows: [], total: 0 }; });
  var s = rows(), n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, 5).getValues().forEach(function (r) {
    if (ms(r[1]) !== month) return;
    var p = String(r[0]); if (!by[p]) return;
    var a = Number(r[4]) || 0;
    by[p].rows.push({ p: r[2], w: r[3], a: a }); by[p].total += a;
  });
  var ss = salesSheet(), sn = ss.getLastRow();
  if (sn > 1) ss.getRange(2, 1, sn - 1, 4).getValues().forEach(function (r) {
    var p = String(r[0]); if (ms(r[1]) !== month || !by[p] || !gridFor(by[p].role)) return;
    var t = Number(r[3]) || 0; by[p].sales = t; by[p].total += t;
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
  var role = ROLES.indexOf(b.role) >= 0 ? b.role : (u && u.role) || 'user';
  if (u && u.role === 'admin' && role !== 'admin') return { error: 'admin' };
  if (!u && (!b.password || String(b.password).length < 6)) return { error: 'short' };
  if (u && !b.password) {
    users().getRange(u.row, 2, 1, 2).setValues([[b.name || u.name, role]]);
    CacheService.getScriptCache().remove('u_' + phone);
    return { ok: true };
  }
  if (String(b.password).length < 6) return { error: 'short' };
  putUser(phone, b.name || (u && u.name) || '', role, String(b.password), true);
  return { ok: true };
}

function importZp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), src = null;
  ss.getSheets().forEach(function (x) { if (!src && /База|zp_import/.test(x.getName())) src = x; });
  if (!src) throw new Error('Нет листа с данными zp: скопируйте лист «База» из таблицы zp в эту таблицу');
  var v = src.getDataRange().getValues(), items = {}, names = {};
  v.forEach(function (r) { if (!r[0]) return; try { items[String(r[0])] = JSON.parse(r[2]); } catch (e) {} });
  (items['@mgrs'] || []).forEach(function (m) { names[m.id] = String(m.name).toLowerCase().replace(/ё/g, 'е'); });
  var us = listUsers(), map = {}, log = [];
  Object.keys(names).forEach(function (id) {
    var f = us.filter(function (u) { return gridFor(u.role) && String(u.name).toLowerCase().replace(/ё/g, 'е').indexOf(names[id]) === 0; });
    if (f.length === 1) map[id] = f[0]; else log.push('Не найден сотрудник (с ролью менеджера/руководителя): ' + names[id]);
  });
  var s = salesSheet(), n = s.getLastRow(), have = {};
  if (n > 1) s.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { have[String(r[0]) + '|' + ms(r[1])] = 1; });
  var add = [], now = Date.now(), cnt = {};
  Object.keys(items).forEach(function (k) {
    var q = k.split('|'); if (q.length !== 2 || !map[q[0]] || !/^\d{4}-\d{2}$/.test(q[1])) return;
    var d = items[k], u = map[q[0]], key = u.phone + '|' + q[1];
    if (have[key] || !d || !d.cats) return;
    var cl = { adj: String(d.adj || ''), company: String(d.company || ''), kpi: [0, 1, 2].map(function (i) { return String((d.kpi || [])[i] || ''); }), cats: d.cats.slice(0, 7).map(function (c) { return { w: [0, 1, 2, 3].map(function (i) { return String((c.w || [])[i] || ''); }), rev: String(c.rev || '') }; }) };
    var used = cl.cats.some(function (c) { return snum(c.rev) || c.w.some(snum); }) || snum(cl.adj) || snum(cl.company);
    if (!used) return;
    add.push([u.phone, q[1], JSON.stringify(cl), calcSales(cl, gridFor(u.role, q[1])).total, now]);
    cnt[u.name] = (cnt[u.name] || 0) + 1;
  });
  if (add.length) {
    var st = s.getLastRow() + 1;
    s.getRange(st, 1, add.length, 2).setNumberFormat('@');
    s.getRange(st, 1, add.length, 5).setValues(add);
  }
  Object.keys(cnt).forEach(function (k) { log.push(k + ': перенесено месяцев ' + cnt[k]); });
  log.push('Всего строк: ' + add.length);
  Logger.log(log.join('\n'));
  return log.join('\n');
}
