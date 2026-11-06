/**
 * تقاريري — الخادم (Google Apps Script مربوط بالشيت)
 * الأوراق تُنشأ تلقائياً عند أول طلب، أو شغّل الدالة setup() يدوياً مرة واحدة.
 * النشر: Deploy > New deployment > Web app
 *        Execute as: Me   |   Who has access: Anyone
 */

const SCHEMA = {
  Reports:   ['id','type','title','label','entity','place','date','created_at','updated_at','pinned','archived','details','notes'],
  Decisions: ['id','report_id','text','owner','due','status','done_at','created_at','updated_at'],
  People:    ['name','entity','created_at'],
  Settings:  ['key','value']
};
const TOKEN_DAYS = 120;   // مدة بقاء الجهاز متصلاً
const MAX_FAILS  = 5;     // محاولات خاطئة قبل القفل
const LOCK_MIN   = 10;    // مدة القفل بالدقائق

/* ========== نقاط الدخول ========== */

function setup() {
  Object.keys(SCHEMA).forEach(sheet_);
  return 'تم تجهيز الأوراق';
}

function doGet() {
  return out_({ ok: true, app: 'taqarir', time: new Date().toISOString() });
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = route_(req);
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return out_(res);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  if (req.action === 'login') return login_(String(req.pin || ''));
  if (!validToken_(req.token)) return { ok: false, error: 'AUTH' };
  switch (req.action) {
    case 'ping':      return { ok: true };
    case 'getAll':    return { ok: true, data: getAll_() };
    case 'batch':     return batch_(req.ops || []);
    case 'changePin': return changePin_(String(req.oldPin || ''), String(req.newPin || ''));
    case 'logout':    revokeToken_(req.token); return { ok: true };
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

/* ========== الأوراق ========== */

// معرّف الشيت «تقاريري — قاعدة البيانات»
const SHEET_ID = '1NsOYdDQAGbi1LPENpPNFy91nydz_DIqkt9tCo_-18x8';
let SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(SHEET_ID)); }

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const h = SCHEMA[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#EEF2F7');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@'); // نص خام: لا تحويل للتواريخ
  }
  return sh;
}

function cell_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? '1' : '';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function findRow_(sh, col, val) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  const s = String(val);
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === s) return i + 2;
  return 0;
}

function upsert_(name, obj, keyField) {
  keyField = keyField || 'id';
  const sh = sheet_(name), h = SCHEMA[name];
  const row = h.map(function (k) { return cell_(obj[k]); });
  const found = findRow_(sh, h.indexOf(keyField) + 1, obj[keyField]);
  const target = found || sh.getLastRow() + 1;
  const rng = sh.getRange(target, 1, 1, h.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
}

function deleteWhere_(name, pred) {
  const sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return;
  const v = sh.getRange(2, 1, last - 1, SCHEMA[name].length).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (pred(v[i])) sh.deleteRow(i + 2);
}

function readAll_(name) {
  const sh = sheet_(name), h = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, last - 1, h.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const o = {};
      h.forEach(function (k, i) {
        let x = r[i];
        if (x instanceof Date) x = Utilities.formatDate(x, tz, k === 'due' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm");
        o[k] = x === null ? '' : String(x);
      });
      return o;
    });
}

function getSetting_(key) {
  const sh = sheet_('Settings'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === key) return String(v[i][1]);
  return null;
}

/* ========== قراءة كل البيانات ========== */

function getAll_() {
  const truthy = function (x) { return x === '1' || x === 'TRUE' || x === 'true'; };
  const reports = readAll_('Reports').map(function (r) {
    let d = {};
    try { d = JSON.parse(r.details || '{}'); } catch (e) {}
    r.details = d;
    r.pinned = truthy(r.pinned);
    r.archived = truthy(r.archived);
    return r;
  });
  let labels = null;
  try { labels = JSON.parse(getSetting_('labels') || 'null'); } catch (e) {}
  return { reports: reports, decisions: readAll_('Decisions'), labels: labels };
}

/* ========== الكتابة (دفعة عمليات) ========== */

function batch_(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let done = 0;
  try {
    for (let i = 0; i < ops.length; i++) { applyOp_(ops[i]); done++; }
    return { ok: true, done: done };
  } catch (err) {
    return { ok: false, done: done, error: String((err && err.message) || err) };
  } finally {
    lock.releaseLock();
  }
}

function applyOp_(op) {
  switch (op.op) {
    case 'saveReport':   return saveReport_(op.report, op.decisions || []);
    case 'deleteReport': return deleteReport_(op.id);
    case 'saveDecision': return upsert_('Decisions', op.decision);
    case 'saveLabels':   return upsert_('Settings', { key: 'labels', value: JSON.stringify(op.labels || []) }, 'key');
  }
  throw new Error('UNKNOWN_OP ' + op.op);
}

function saveReport_(r, decs) {
  if (!r || !r.id) throw new Error('NO_ID');
  const copy = Object.assign({}, r, { details: JSON.stringify(r.details || {}) });
  upsert_('Reports', copy);
  const keep = {};
  decs.forEach(function (d) { keep[String(d.id)] = 1; });
  deleteWhere_('Decisions', function (row) { return String(row[1]) === String(r.id) && !keep[String(row[0])]; });
  decs.forEach(function (d) { upsert_('Decisions', Object.assign({}, d, { report_id: r.id })); });
  savePeople_(r, decs);
}

function deleteReport_(id) {
  deleteWhere_('Reports', function (row) { return String(row[0]) === String(id); });
  deleteWhere_('Decisions', function (row) { return String(row[1]) === String(id); });
}

function savePeople_(r, decs) {
  const d = r.details || {};
  const names = [].concat(d.attendees || [], d.absent || [], decs.map(function (x) { return x.owner; }))
    .map(function (n) { return String(n || '').trim(); }).filter(String);
  if (!names.length) return;
  const entity = String(r.entity || '').trim();
  const known = {};
  readAll_('People').forEach(function (p) { known[p.name + '|' + p.entity] = 1; });
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  const rows = [];
  names.forEach(function (n) {
    const k = n + '|' + entity;
    if (!known[k]) { known[k] = 1; rows.push([n, entity, now]); }
  });
  if (!rows.length) return;
  const sh = sheet_('People');
  const rng = sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3);
  rng.setNumberFormat('@');
  rng.setValues(rows);
}

/* ========== الرمز السري والجلسات ========== */

function props_() { return PropertiesService.getScriptProperties(); }

function hash_(pin, salt) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pin, Utilities.Charset.UTF_8));
}

function login_(pin) {
  const p = props_();
  const lockUntil = Number(p.getProperty('lock_until') || 0);
  if (Date.now() < lockUntil) return { ok: false, error: 'LOCKED', minutes: Math.ceil((lockUntil - Date.now()) / 60000) };
  if (!/^\d{4,8}$/.test(pin)) return { ok: false, error: 'PIN_FORMAT' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let hash = p.getProperty('pin_hash'), created = false;
    if (!hash) { // أول دخول: الرمز المكتوب يصبح هو الرمز
      const salt = Utilities.getUuid();
      p.setProperty('pin_salt', salt);
      p.setProperty('pin_hash', hash_(pin, salt));
      created = true;
    } else if (hash_(pin, p.getProperty('pin_salt')) !== hash) {
      const f = Number(p.getProperty('fails') || 0) + 1;
      if (f >= MAX_FAILS) {
        p.setProperty('lock_until', String(Date.now() + LOCK_MIN * 60000));
        p.setProperty('fails', '0');
        return { ok: false, error: 'LOCKED', minutes: LOCK_MIN };
      }
      p.setProperty('fails', String(f));
      return { ok: false, error: 'BAD_PIN', left: MAX_FAILS - f };
    }
    p.setProperty('fails', '0');
    return { ok: true, token: issueToken_(), created: created };
  } finally {
    lock.releaseLock();
  }
}

function changePin_(oldPin, newPin) {
  const p = props_();
  if (hash_(oldPin, p.getProperty('pin_salt')) !== p.getProperty('pin_hash')) return { ok: false, error: 'BAD_PIN' };
  if (!/^\d{4,8}$/.test(newPin)) return { ok: false, error: 'PIN_FORMAT' };
  const salt = Utilities.getUuid();
  p.setProperty('pin_salt', salt);
  p.setProperty('pin_hash', hash_(newPin, salt));
  p.setProperty('tokens', '{}'); // خروج كل الأجهزة الأخرى
  return { ok: true, token: issueToken_() };
}

function tokens_() {
  try { return JSON.parse(props_().getProperty('tokens') || '{}'); } catch (e) { return {}; }
}

function issueToken_() {
  const t = tokens_(), now = Date.now();
  Object.keys(t).forEach(function (k) { if (t[k] < now) delete t[k]; });
  const tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  t[tok] = now + TOKEN_DAYS * 864e5;
  props_().setProperty('tokens', JSON.stringify(t));
  return tok;
}

function validToken_(tok) {
  if (!tok) return false;
  const t = tokens_();
  return !!t[tok] && t[tok] > Date.now();
}

function revokeToken_(tok) {
  const t = tokens_();
  delete t[tok];
  props_().setProperty('tokens', JSON.stringify(t));
}

/** لإعادة تعيين الرمز السري إن نسيته: شغّلها يدوياً من المحرر، ثم ادخل برمز جديد */
function resetPin() {
  ['pin_hash', 'pin_salt', 'fails', 'lock_until'].forEach(function (k) { props_().deleteProperty(k); });
  props_().setProperty('tokens', '{}');
  return 'تم: أول دخول قادم سيحدد الرمز الجديد';
}
