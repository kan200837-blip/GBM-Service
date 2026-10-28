/**
 * รุ่น 7 — สร้างเป็นไฟล์ใหม่ใน Apps Script ชื่อ Report.gs (ใช้คู่กับ Code.gs)
 *
 *  1) เข้าใช้งานด้วย "อีเมล + รหัสเข้าใช้" (ผู้บริหาร / ผู้บันทึก / ผู้ดูแลระบบ)
 *     จำเป็นเมื่อเปิดเว็บแบบไม่ต้องลงชื่อเข้าใช้ หรือใช้ Gmail ส่วนตัว ซึ่งระบบอ่านอีเมลผู้เปิดไม่ได้
 *     ผู้ดูแลระบบตั้งรหัสให้แต่ละคนที่หน้า ผู้ดูแลระบบ > ผู้ใช้และสิทธิ์ · รหัสเก็บเป็นแฮชในชีต User_Codes
 *     ครั้งแรก: รัน setMyAdminCode() จากหน้าแก้ไขสคริปต์ เพื่อรับรหัสของเจ้าของระบบ (ดูใน Execution log)
 *  2) รายการเหตุการณ์ที่ รปภ. บันทึกขณะตรวจ (getGuardEvents) เปิดดูได้ว่าเป็นเรื่องอะไร วันที่ เวลา จุดไหน ใครพบ
 *  3) รายงานสรุปรายสัปดาห์ / รายเดือน / รายไตรมาส ส่งอีเมลผู้บริหารอัตโนมัติ
 *     ครั้งแรก: รัน authorizeReports() หนึ่งครั้งเพื่ออนุญาตสิทธิ์ส่งอีเมลและตั้งเวลา
 *     สัปดาห์: ทุกวันจันทร์ 07:00 (สรุป 7 วันก่อนหน้า) · เดือน: วันที่ 1 เวลา 08:00 (สรุปเดือนก่อน)
 *     ไตรมาส: วันที่ 1 ต.ค. / ม.ค. / เม.ย. / ก.ค. เวลา 08:00 (สรุปไตรมาสที่เพิ่งจบ)
 */

var CALL_USER_ = null;
const TOKEN_DAYS = 30;
const CODE_SHEET = 'User_Codes';
const CALL_DENY_ = ['doGet', 'setup', 'connectExistingDb', 'upgradeDb', 'clearCache', 'authorizeDrive', 'authorizeReports', 'setMyAdminCode',
  'callAs', 'loginWithCode', 'logoutToken', 'rptWeeklyJob', 'rptMonthlyJob'];

/** ฟังก์ชันที่รันได้เฉพาะเจ้าของสคริปต์จากหน้าแก้ไข (กันการเรียกผ่าน google.script.run จากคนทั่วไป) */
function ownerOnly_() {
  const a = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const o = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!a || a !== o) throw new Error('ฟังก์ชันนี้รันได้จากหน้าแก้ไขสคริปต์โดยเจ้าของระบบเท่านั้น');
}

/* ---------------- เข้าใช้งานด้วยรหัส ---------------- */

/** หน้าเว็บเรียกทุกฟังก์ชันผ่านตัวนี้เมื่อเข้าสู่ระบบด้วยรหัส: ตั้งผู้ใช้ของการเรียกนี้ แล้วเรียกฟังก์ชันปลายทาง */
function callAs(token, fn, args) {
  fn = String(fn || '');
  if (CALL_DENY_.indexOf(fn) >= 0 || /_$/.test(fn) || typeof globalThis[fn] !== 'function') throw new Error('ไม่อนุญาตให้เรียก ' + fn);
  CALL_USER_ = tokenUser_(token);
  try { return globalThis[fn].apply(null, Array.isArray(args) ? args : []); }
  finally { CALL_USER_ = null; }
}

function tokenUser_(token) {
  token = String(token || '');
  const expired = 'SESSION_EXPIRED: หมดเวลาเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่';
  if (!/^[A-Za-z0-9-]{40,100}$/.test(token)) throw new Error(expired);
  const props = PropertiesService.getScriptProperties(), raw = props.getProperty('TK_' + token);
  if (!raw) throw new Error(expired);
  let t = null;
  try { t = JSON.parse(raw); } catch (e) { t = null; }
  if (!t || !(t.exp > Date.now())) { props.deleteProperty('TK_' + token); throw new Error(expired); }
  return { email: String(t.email).toLowerCase() };
}

/** ลบ token: email = ลบของผู้ใช้นั้นทั้งหมด · ทุกครั้งลบ token ที่หมดอายุ */
function purgeTokens_(email) {
  const p = PropertiesService.getScriptProperties(), all = p.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('TK_') !== 0) return;
    let t = null;
    try { t = JSON.parse(all[k]); } catch (e) { t = null; }
    if (!t || !(t.exp > now) || (email && String(t.email) === email)) p.deleteProperty(k);
  });
}

function codeSheet_() {
  const ss = db_();
  let sh = ss.getSheetByName(CODE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CODE_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['email', 'code_hash', 'set_at']]).setFontWeight('bold').setBackground('#e6efec');
    sh.setFrozenRows(1);
  }
  return sh;
}
function hashUserCode_(email, code) { return hashPin_('user|' + email + '|' + code); }
function codeHashOf_(email) {
  const v = codeSheet_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === email) return String(v[i][1] || '');
  return '';
}
function writeCode_(email, hash) {
  const sh = codeSheet_(), v = sh.getDataRange().getValues(), now = new Date().toISOString();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === email) { sh.getRange(i + 1, 2, 1, 2).setValues([[hash, hash ? now : '']]); return; }
  }
  if (hash) sh.appendRow([email, hash, now]);
}
function randomCode_(n) {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const b = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  let s = '';
  for (let i = 0; i < n; i++) s += a.charAt(parseInt(b.substr(i * 2, 2), 16) % a.length);
  return s;
}

/** เข้าสู่ระบบ: อีเมล + รหัสเข้าใช้ → token (อายุ 30 วัน) · ผิดเกิน 8 ครั้ง/15 นาที ต่ออีเมล พักการใช้ */
function loginWithCode(email, code) {
  email = String(email || '').trim().toLowerCase();
  code = String(code || '').trim();
  if (!email || !code) throw new Error('กรุณาใส่อีเมลและรหัสเข้าใช้');
  const cache = CacheService.getScriptCache(), fk = 'lgf_' + Utilities.base64EncodeWebSafe(email).slice(0, 120);
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 8) throw new Error('ใส่รหัสผิดหลายครั้ง ลองใหม่ใน 15 นาที');
  const u = read_('Users').filter(function (x) { return String(x.email || '').trim().toLowerCase() === email && isActive_(x.active); })[0];
  const h = codeHashOf_(email);
  if (!u || !h || h !== hashUserCode_(email, code)) {
    cache.put(fk, String(fails + 1), 900);
    throw new Error('อีเมลหรือรหัสเข้าใช้ไม่ถูกต้อง (หรือบัญชีถูกปิดใช้งาน)');
  }
  cache.remove(fk);
  purgeTokens_('');
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('TK_' + token, JSON.stringify({ email: email, exp: Date.now() + TOKEN_DAYS * 864e5 }));
  audit_({ email: email, role: u.role }, [{ action: 'เข้าสู่ระบบด้วยรหัส', sheet: 'Users', key: email, field: '', old: '', new: u.role }]);
  return { token: token, email: email, name: String(u.name || email), role: String(u.role), days: TOKEN_DAYS };
}

function logoutToken(token) {
  token = String(token || '');
  if (/^[A-Za-z0-9-]{40,100}$/.test(token)) PropertiesService.getScriptProperties().deleteProperty('TK_' + token);
  return true;
}

/** ผู้ดูแลระบบตั้ง/ล้างรหัสเข้าใช้ของผู้ใช้ (code ว่าง = ล้าง) · ตั้งใหม่แล้วการเข้าสู่ระบบเดิมของคนนั้นหมดอายุทันที */
function setUserCode(email, code) {
  const u = requireAdmin_();
  email = String(email || '').trim().toLowerCase();
  code = String(code || '').trim();
  if (code && !/^[A-Za-z0-9]{6,32}$/.test(code)) throw new Error('รหัสเข้าใช้ต้องเป็นตัวอักษรอังกฤษหรือตัวเลข 6–32 ตัว');
  if (!read_('Users').some(function (x) { return String(x.email || '').trim().toLowerCase() === email; })) throw new Error('ไม่พบผู้ใช้ ' + email + ' (บันทึกรายชื่อผู้ใช้ก่อน)');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    writeCode_(email, code ? hashUserCode_(email, code) : '');
    purgeTokens_(email);
    audit_(u, [{ action: code ? 'ตั้งรหัสเข้าใช้' : 'ล้างรหัสเข้าใช้', sheet: CODE_SHEET, key: email, field: 'code', old: '', new: code ? '(ตั้งใหม่)' : '' }]);
  } finally { lock.releaseLock(); }
  return { email: email, has_code: !!code };
}

/** อีเมลที่ตั้งรหัสเข้าใช้แล้ว (หน้าผู้ดูแลระบบ) */
function getUserCodeStatus() {
  requireAdmin_();
  const out = {};
  codeSheet_().getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && r[1]) out[String(r[0]).trim().toLowerCase()] = String(r[2] || ''); });
  return out;
}

/** รันครั้งแรกจากหน้าแก้ไขสคริปต์: ตั้งรหัสเข้าใช้ของเจ้าของระบบ แล้วดูรหัสใน Execution log */
function setMyAdminCode() {
  ownerOnly_();
  const email = String(Session.getEffectiveUser().getEmail()).toLowerCase();
  const sh = db_().getSheetByName('Users'), h = headOf_(sh);
  if (!read_('Users').some(function (x) { return String(x.email || '').trim().toLowerCase() === email; })) {
    const o = { email: email, name: 'เจ้าของระบบ', role: 'admin', units: '*', active: 1 };
    sh.appendRow(h.map(function (c) { return o[c] === undefined ? '' : o[c]; }));
  }
  const code = randomCode_(10);
  writeCode_(email, hashUserCode_(email, code));
  purgeTokens_(email);
  Logger.log('อีเมล: ' + email + '   รหัสเข้าใช้: ' + code + '   (จดไว้ ระบบไม่แสดงรหัสนี้อีก)');
  return code;
}

/* ---------------- เหตุการณ์ที่ รปภ. บันทึกขณะตรวจ ---------------- */

function guardEventRows_(keep) {
  const reg = gLoad_(), pn = {}, rn = {};
  reg.points.forEach(function (p) { pn[p.point_id] = p; });
  reg.routes.forEach(function (r) { rn[r.route_id] = r.name; });
  const out = [];
  gReadRounds_().forEach(function (r) {
    if (!r.ev || !keep(r)) return;
    let sc = [];
    try { sc = JSON.parse(r.scans); } catch (e) { sc = []; }
    const f = fiscalOf_(r.duty_date);
    sc.forEach(function (x) {
      if (!x[2]) return;
      const p = pn[x[1]] || {};
      out.push({ date: r.date, time: String(x[0]).slice(0, 5), duty_date: r.duty_date, fy: f.fy, m: f.month, shift: r.shift, name: r.name,
        point: p.name || String(x[1]), building: p.building || '', route: rn[r.route_id] || r.route_id, event: String(x[2]) });
    });
  });
  out.sort(function (a, b) { const x = a.date + a.time, y = b.date + b.time; return x < y ? 1 : (x > y ? -1 : 0); });
  return out;
}

/**
 * รายการเหตุการณ์ q = {ym:'YYYY-MM'} หรือ {fy, month?} · กรองเพิ่ม {name, point}
 * คืน {total, rows:[{date,time,shift,name,point,building,route,event,fy,m}]}
 */
function getGuardEvents(q) {
  q = q || {};
  currentUser_();
  return cached_(function () {
    const ym = /^\d{4}-\d{2}$/.test(String(q.ym || '')) ? String(q.ym) : '';
    const fy = Number(q.fy) || 0, m = Number(q.month) || 0, name = gStr_(q.name), point = gStr_(q.point);
    if (!ym && !fy) throw new Error('ระบุเดือนหรือปีงบประมาณ');
    const rows = guardEventRows_(function (r) {
      if (ym && r.duty_date.slice(0, 7) !== ym) return false;
      if (fy) { const f = fiscalOf_(r.duty_date); if (f.fy !== fy || (m && f.month !== m)) return false; }
      return !name || r.name === name;
    }).filter(function (x) { return !point || x.point === point; });
    return { total: rows.length, rows: rows.slice(0, 3000) };
  });
}

/* ---------------- รายงานสรุปส่งอีเมลผู้บริหาร ---------------- */

const TH_MON_ = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const RPT_COL_ = { ok: ['#1d7f4b', '#e3f3ea', 'ผ่านเป้า'], warn: ['#a96a00', '#fbefd5', 'ใกล้เป้า'], bad: ['#bf3a30', '#fae3e0', 'ไม่ผ่านเป้า'],
  verify: ['#6a55b0', '#ebe6f7', 'ต้องยืนยัน'], nodata: ['#8a9a97', '#eceff0', 'ไม่มีข้อมูล'], info: ['#3d6c9e', '#e1ecf7', 'ติดตาม'] };
const RPT_KINDS_ = { week: 'สัปดาห์', month: 'เดือน', quarter: 'ไตรมาส' };

function n_(x) { return x === null || x === undefined || x === '' || isNaN(Number(x)) ? null : Number(x); }
function thd_(ds) { const p = String(ds).split('-'); return Number(p[2]) + ' ' + TH_MON_[Number(p[1]) - 1] + ' ' + pad2_((Number(p[0]) + 543) % 100); }
function calOf_(fy, m) { return { y: fy - 543 - (m <= 3 ? 1 : 0), m: ((m + 8) % 12) + 1 }; }
function fmon_(fy, m) { const c = calOf_(fy, m); return TH_MON_[c.m - 1] + ' ' + pad2_((c.y + 543) % 100); }
function calFirst_(fy, m) { const c = calOf_(fy, m); return c.y + '-' + pad2_(c.m) + '-01'; }
function calLast_(fy, m) { const c = calOf_(fy, m); return c.y + '-' + pad2_(c.m) + '-' + pad2_(new Date(Date.UTC(c.y, c.m, 0)).getUTCDate()); }
function upto_(m) { const a = []; for (let i = 1; i <= m; i++) a.push(i); return a; }
function esc_(s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function settingVal_(key) { let v = null; read_('Settings').forEach(function (r) { if (r.key === key) v = r.value; }); return v; }
function setSetting_(key, val) {
  const sh = db_().getSheetByName('Settings'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([key, val]);
}

function rptPeriod_(kind) {
  const today = todayStr_(), cur = fiscalOf_(today);
  if (kind === 'week') {
    const to = gDayAdd_(today, -1), from = gDayAdd_(today, -7), f = fiscalOf_(to);
    return { kind: kind, from: from, to: to, fy: f.fy, months: [f.month], ytd: upto_(f.month), level: 'month',
      label: 'ประจำสัปดาห์ ' + thd_(from) + ' – ' + thd_(to), plabel: fmon_(f.fy, f.month) + ' (ถึง ' + thd_(to) + ')' };
  }
  if (kind === 'month') {
    const to = gDayAdd_(today.slice(0, 8) + '01', -1), f = fiscalOf_(to);
    return { kind: kind, from: to.slice(0, 8) + '01', to: to, fy: f.fy, months: [f.month], ytd: upto_(f.month), level: 'month',
      label: 'ประจำเดือน ' + fmon_(f.fy, f.month), plabel: fmon_(f.fy, f.month) };
  }
  let q = Math.ceil(cur.month / 3) - 1, fy = cur.fy;
  if (q < 1) { q = 4; fy -= 1; }
  const ms = [q * 3 - 2, q * 3 - 1, q * 3];
  return { kind: 'quarter', from: calFirst_(fy, ms[0]), to: calLast_(fy, ms[2]), fy: fy, months: ms, ytd: upto_(ms[2]), level: 'quarter',
    label: 'ประจำไตรมาส ' + q + ' ปีงบประมาณ ' + fy + ' (' + fmon_(fy, ms[0]) + ' – ' + fmon_(fy, ms[2]) + ')', plabel: 'ไตรมาส ' + q };
}

/** ตัวชี้วัดพร้อมค่ารายเดือนของปีงบ (กรอกเอง + รายไตรมาส + คำนวณจากแผน/ผลตรวจ) และเป้าหมายของปีนั้น — ตรรกะเดียวกับหน้าเว็บ */
function rptKpis_(fy) {
  const master = read_('KPI_Master').filter(function (k) { return isActive_(k.active); }).sort(function (a, b) { return Number(a.order) - Number(b.order); });
  const tg = {}, cells = {}, byId = {};
  read_('KPI_Targets').forEach(function (t) { if (Number(t.fy) === fy && t.dir) tg[t.kpi_id] = t; });
  master.forEach(function (k) { byId[k.kpi_id] = k; });
  function put(id, m, a, b) { (cells[id] = cells[id] || {})[m] = { num: a, den: b }; }
  read_('KPI_Data').forEach(function (r) {
    if (Number(r.fy) !== fy) return;
    const a = n_(r.num), b = n_(r.den);
    if (a !== null || b !== null) put(String(r.kpi_id), Number(r.month), a, b);
  });
  read_('KPI_Quarter').forEach(function (q) {
    const k = byId[q.kpi_id], v = n_(q.value), qn = Number(q.quarter);
    if (!k || v === null || Number(q.fy) !== fy || k.source === 'pm' || !(qn >= 1 && qn <= 4)) return;
    [qn * 3 - 2, qn * 3 - 1, qn * 3].forEach(function (m, i) {
      if (cells[k.kpi_id] && cells[k.kpi_id][m]) return;
      if (k.kind === 'ratio') put(k.kpi_id, m, v, 100); else put(k.kpi_id, m, i === 2 ? v : 0, null);
    });
  });
  const pmK = master.filter(function (k) { return k.source === 'pm'; });
  if (pmK.length) {
    const items = read_('PM_Items'), plan = {};
    read_('PM_Plan').forEach(function (r) { if (Number(r.fy) === fy) plan[norm_(r.code) + '|' + Number(r.month)] = r; });
    const logs = readLogs_(0).filter(function (l) { return l.fy === fy; });
    pmK.forEach(function (k) {
      delete cells[k.kpi_id];
      const its = items.filter(function (i) { return i.kpi_id === k.kpi_id; }).map(function (i) { return norm_(i.code); });
      const sum = k.pm_calc === 'sum';
      for (let m = 1; m <= 12; m++) {
        if (k.kind === 'count') {
          const ml = logs.filter(function (l) { return l.month === m && its.indexOf(l.code) >= 0; });
          if (ml.length) put(k.kpi_id, m, ml.filter(function (l) { return l.result === 'พบปัญหา'; }).length, null);
          continue;
        }
        let num = 0, den = 0, any = false;
        its.forEach(function (c) {
          const r = plan[c + '|' + m]; if (!r) return;
          const p = n_(r.plan), a = n_(r.actual); if (!(p > 0)) return;
          den += sum ? p : 1;
          if (a !== null) { any = true; if (sum) num += Math.min(a, p); else if (a >= p) num++; }
        });
        if (any) put(k.kpi_id, m, num, den);
      }
    });
  }
  return master.map(function (k0) {
    const k = Object.assign({}, k0), t = tg[k.kpi_id];
    if (t) { k.dir = t.dir; k.thr = n_(t.thr) || 0; k.target_text = t.target_text || ''; }
    k.target_text = k.target_text || targetText_(k.dir, k.thr, k.kind);
    k.cells = cells[k.kpi_id] || {};
    return k;
  });
}
function rptStatus_(k, v, level) {
  if (v === null || v === undefined) return 'nodata';
  const d = k.dir, t = Number(k.thr);
  if (d === 'ge') return v >= t - 1e-9 ? 'ok' : (v >= t * 0.9 ? 'warn' : 'bad');
  if (d === 'gt') return v > t ? 'ok' : (v >= t * 0.9 ? 'warn' : 'bad');
  if (d === 'lt') return v < t ? 'ok' : (v < t * 1.5 ? 'warn' : 'bad');
  if (d === 'zero') return v === 0 ? 'ok' : 'bad';
  if (d === 'le_year') { if (level === 'year') return v <= t ? 'ok' : 'bad'; return v === 0 ? 'ok' : 'warn'; }
  return 'info';
}
function rptAgg_(k, months, level) {
  let num = 0, den = 0, n = 0, bad = 0;
  months.forEach(function (m) {
    const c = k.cells[m]; if (!c) return;
    if (k.kind === 'ratio') {
      if (c.num === null || c.den === null || !(c.den > 0)) return;
      if (Number(k.max100) === 1 && c.num > c.den) { bad++; return; }
      num += c.num; den += c.den; n++;
    } else { if (c.num === null) return; num += c.num; n++; }
  });
  if (!n) return { value: null, status: bad ? 'verify' : 'nodata', n: 0 };
  const v = k.kind === 'ratio' ? num / den * 100 : num;
  return { value: v, status: rptStatus_(k, v, level), n: n };
}
function rptVal_(k, v) {
  if (v === null || v === undefined) return '–';
  return k.kind === 'ratio' ? (Math.round(v * 10) / 10).toFixed(1) + '%' : String(Math.round(v * 10) / 10);
}

function badge_(st, txt) { const c = RPT_COL_[st] || RPT_COL_.nodata; return '<span style="display:inline-block;padding:1px 9px;border-radius:10px;font-size:12px;font-weight:600;color:' + c[0] + ';background:' + c[1] + '">' + esc_(txt || c[2]) + '</span>'; }
function sec_(t) { return '<h3 style="font-size:16px;color:#0d6e68;border-left:5px solid #0d6e68;padding-left:8px;margin:22px 0 6px">' + esc_(t) + '</h3>'; }
function para_(t) { return '<p style="margin:4px 0 8px;font-size:13.5px;color:#5d716e">' + t + '</p>'; }
function tbl_(heads, rows) {
  const th = 'padding:6px 8px;background:#eef4f2;text-align:left;font-size:12.5px;border-bottom:1px solid #cfdcd8';
  const td = 'padding:6px 8px;border-bottom:1px solid #e3ebe8;font-size:13px;vertical-align:top';
  return '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:4px 0 10px"><tr>' +
    heads.map(function (h) { return '<th style="' + th + '">' + esc_(h) + '</th>'; }).join('') + '</tr>' +
    rows.map(function (r) { return '<tr>' + r.map(function (c) { return '<td style="' + td + '">' + c + '</td>'; }).join('') + '</tr>'; }).join('') + '</table>';
}
function tile_(v, lbl, col) {
  return '<td style="width:25%;padding:10px 12px;border:1px solid #dbe4e1;border-left:5px solid ' + col + ';background:#fff">' +
    '<div style="font-size:26px;font-weight:700;color:' + col + '">' + esc_(v) + '</div><div style="font-size:12.5px;color:#5d716e">' + esc_(lbl) + '</div></td>';
}

/** สร้างอีเมลรายงาน kind = week | month | quarter → {subject, html} */
function buildReport_(kind) {
  if (!RPT_KINDS_[kind]) throw new Error('ชนิดรายงานไม่ถูกต้อง');
  const P = rptPeriod_(kind), ks = rptKpis_(P.fy);
  const org = String(settingVal_('org') || 'ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม');
  const inR = function (d) { return d >= P.from && d <= P.to; };
  const rows = ks.map(function (k) { return { k: k, p: rptAgg_(k, P.months, P.level), y: rptAgg_(k, P.ytd, 'year') }; });
  const cnt = { ok: 0, warn: 0, bad: 0, verify: 0, nodata: 0, info: 0 };
  rows.forEach(function (r) { cnt[r.y.status]++; });
  const judged = cnt.ok + cnt.warn + cnt.bad;
  const watch = rows.filter(function (r) { return ['bad', 'warn', 'verify'].indexOf(r.y.status) >= 0 || ['bad', 'verify'].indexOf(r.p.status) >= 0; });

  const itemName = {};
  read_('PM_Items').forEach(function (i) { itemName[norm_(i.code)] = i.name; });
  const logs = readLogs_(kind === 'quarter' ? 130 : 45).filter(function (l) { return inR(l.date); });
  const probs = logs.filter(function (l) { return l.result === 'พบปัญหา'; });
  const rounds = gReadRounds_().filter(function (r) { return inR(r.duty_date); });
  const evs = guardEventRows_(function (r) { return inR(r.duty_date); });
  const cName = {};
  read_('Contracts').forEach(function (c) { cName[String(c.contract_id)] = c.company; });
  const visits = readVisits_().filter(function (v) { return !v.voided && inR(v.done_date); });
  const late = buildContracts_({ role: 'admin', email: '', units: ['*'] }, { lite: true }).filter(function (c) { return isActive_(c.active) && c.sum.late_rounds > 0; });

  let pm = null;
  if (kind !== 'week') {
    const plan = read_('PM_Plan').filter(function (r) { return Number(r.fy) === P.fy && P.months.indexOf(Number(r.month)) >= 0 && n_(r.plan) > 0; });
    const notDone = plan.filter(function (r) { return !(n_(r.actual) !== null && n_(r.actual) >= n_(r.plan)); });
    pm = { total: plan.length, done: plan.length - notDone.length, notDone: notDone, pct: plan.length ? Math.round((plan.length - notDone.length) / plan.length * 1000) / 10 : null };
  }
  const missing = [];
  if (kind !== 'week') {
    rows.forEach(function (r) {
      const k = r.k;
      if (k.source === 'pm' || k.source === 'guard') return;
      const ms = P.months.filter(function (m) { const c = k.cells[m]; return !(c && (k.kind === 'ratio' ? c.num !== null && c.den !== null : c.num !== null)); });
      if (ms.length) missing.push([esc_(k.kpi_id + ' ' + k.name), esc_(k.unit || ''), esc_(ms.map(function (m) { return fmon_(P.fy, m); }).join(', '))]);
    });
  }

  const scoreTxt = judged ? Math.round(cnt.ok / judged * 100) + '%' : '–';
  const comp = rounds.filter(function (r) { return r.complete; }).length;
  let h = '<div style="font-family:Tahoma,\'Noto Sans Thai\',Arial,sans-serif;color:#17302d;max-width:780px;margin:0 auto;line-height:1.5">';
  h += '<div style="border-bottom:3px solid #0d6e68;padding-bottom:8px;margin-bottom:12px"><div style="font-size:20px;font-weight:700;color:#0d6e68">รายงานสรุปผลการดำเนินงาน ' + esc_(org) + '</div>' +
    '<div style="font-size:14px;color:#5d716e">' + esc_(P.label) + ' · ปีงบประมาณ ' + P.fy + ' · ออกรายงาน ' + thd_(todayStr_()) + '</div></div>';
  h += '<table cellspacing="6" cellpadding="0" style="width:100%;border-collapse:separate"><tr>' +
    tile_(scoreTxt, 'ตัวชี้วัดผ่านเป้า สะสมปีงบ (' + cnt.ok + '/' + judged + ')', judged && cnt.ok / judged >= 0.8 ? '#1d7f4b' : '#a96a00') +
    tile_(String(cnt.bad), 'ตัวชี้วัดไม่ผ่านเป้า', cnt.bad ? '#bf3a30' : '#1d7f4b') +
    tile_(String(probs.length), 'พบปัญหาจากการตรวจบำรุงรักษา', probs.length ? '#a96a00' : '#1d7f4b') +
    tile_(String(evs.length), 'เหตุการณ์ที่ รปภ. บันทึก', evs.length ? '#a96a00' : '#1d7f4b') + '</tr></table>';
  const badNames = rows.filter(function (r) { return r.y.status === 'bad'; }).map(function (r) { return r.k.name.replace(/\s*\(.*?\)\s*/g, ' ').trim(); });
  h += '<p style="font-size:14.5px;margin:10px 0">' + esc_(judged ? 'สะสมปีงบผ่านเป้า ' + cnt.ok + ' จาก ' + judged + ' ตัวชี้วัด' + (badNames.length ? ' · ไม่ผ่านเป้า: ' + badNames.slice(0, 4).join(', ') + (badNames.length > 4 ? ' และอีก ' + (badNames.length - 4) + ' ตัว' : '') : ' · ไม่มีตัวชี้วัดที่ไม่ผ่านเป้า') : 'ยังไม่มีผลตัวชี้วัดเพียงพอสำหรับสรุป') +
    esc_(' · ช่วงนี้บันทึกบำรุงรักษา ' + logs.length + ' ครั้ง รปภ. ตรวจตรา ' + rounds.length + ' รอบ' + (late.length ? ' · สัญญาเกินกำหนด ' + late.length + ' ฉบับ' : '')) + '</p>';

  if (kind === 'week') {
    h += sec_('ตัวชี้วัดที่ต้องติดตาม (' + watch.length + ')');
    h += watch.length ? tbl_(['ตัวชี้วัด', 'เป้า', 'เดือนนี้ ' + P.plabel, 'สะสมปีงบ', 'สถานะ'], watch.map(function (r) {
      return [esc_(r.k.kpi_id + ' ' + r.k.name), esc_(r.k.target_text), esc_(rptVal_(r.k, r.p.value)), esc_(rptVal_(r.k, r.y.value)), badge_(r.y.status)];
    })) : para_('ทุกตัวชี้วัดผ่านเป้าหรืออยู่ในเกณฑ์');
  } else {
    h += sec_('ผลตัวชี้วัด ' + P.plabel + ' และสะสมปีงบ');
    let g = null;
    const kr = [];
    rows.forEach(function (r) {
      if (r.k.group !== g) { g = r.k.group; kr.push(['<b style="color:#0d6e68">' + esc_(g || '') + '</b>', '', '', '', '']); }
      kr.push([esc_(r.k.kpi_id + ' ' + r.k.name), esc_(r.k.target_text), esc_(rptVal_(r.k, r.p.value)) + ' ' + badge_(r.p.status, ' '), esc_(rptVal_(r.k, r.y.value)), badge_(r.y.status)]);
    });
    h += tbl_(['ตัวชี้วัด', 'เป้า', P.plabel, 'สะสมปีงบ', 'สถานะสะสม'], kr);
  }

  if (pm) {
    h += sec_('การบำรุงรักษาตามแผน ' + P.plabel);
    h += para_('ครบตามแผน ' + pm.done + ' จาก ' + pm.total + ' รายการ-เดือน' + (pm.pct === null ? '' : ' (' + pm.pct + '%)'));
    if (pm.notDone.length) h += tbl_(['รหัส', 'รายการ', 'เดือน', 'ผลจริง / แผน'], pm.notDone.slice(0, 15).map(function (r) {
      const c = norm_(r.code);
      return [esc_(c), esc_(itemName[c] || ''), esc_(fmon_(P.fy, Number(r.month))), esc_((n_(r.actual) === null ? 'ไม่บันทึก' : n_(r.actual)) + ' / ' + n_(r.plan))];
    })) + (pm.notDone.length > 15 ? para_('… และอีก ' + (pm.notDone.length - 15) + ' รายการ') : '');
  }
  h += sec_('ปัญหาที่พบจากการตรวจบำรุงรักษา (' + probs.length + ')');
  h += probs.length ? tbl_(['วันที่', 'รายการ / อุปกรณ์', 'รายละเอียด', 'ผู้บันทึก'], probs.slice(0, 15).map(function (l) {
    return [esc_(thd_(l.date)), esc_((l.eq_id || l.code) + ' ' + (itemName[l.code] || '')), esc_(l.note), esc_(l.by)];
  })) + (probs.length > 15 ? para_('… และอีก ' + (probs.length - 15) + ' รายการ') : '') : para_('ไม่พบปัญหาในช่วงนี้');

  h += sec_('งานรักษาความปลอดภัย');
  h += para_('ตรวจตรา ' + rounds.length + ' รอบ · สแกนครบทุกจุด ' + comp + ' รอบ' + (rounds.length ? ' (' + Math.round(comp / rounds.length * 1000) / 10 + '%)' : '') + ' · เหตุการณ์ที่บันทึก ' + evs.length + ' ครั้ง');
  if (evs.length) h += tbl_(['วันที่ เวลา', 'เรื่องที่พบ', 'จุด / อาคาร', 'ผู้พบ (เวร)'], evs.slice(0, 20).map(function (e) {
    return [esc_(thd_(e.date) + ' ' + e.time), '<b>' + esc_(e.event) + '</b>', esc_(e.point + (e.building ? ' · ' + e.building : '')), esc_(e.name + ' (' + e.shift + ')')];
  })) + (evs.length > 20 ? para_('… และอีก ' + (evs.length - 20) + ' เหตุการณ์ (ดูทั้งหมดในระบบ แท็บตรวจตรา รปภ.)') : '');

  h += sec_('สัญญาบำรุงรักษา');
  h += visits.length ? tbl_(['วันที่', 'บริษัท', 'รอบ', 'ผล', 'หมายเหตุ'], visits.slice(0, 15).map(function (v) {
    return [esc_(thd_(v.done_date)), esc_(cName[v.contract_id] || v.contract_id), esc_(v.round ? 'รอบ ' + v.round : 'นอกรอบ'), badge_(v.result === 'พบปัญหา' ? 'bad' : 'ok', v.result || 'ปกติ'), esc_(v.note)];
  })) : para_('ช่วงนี้ไม่มีการเข้าบำรุงรักษาตามสัญญา');
  if (late.length) h += para_('<b style="color:#bf3a30">สัญญาที่เกินกำหนดรอบ:</b> ' + late.map(function (c) { return esc_(c.company + ' (' + c.sum.late_rounds + ' รอบ)'); }).join(', '));

  if (missing.length) { h += sec_('ข้อมูลตัวชี้วัดที่ยังไม่ส่ง (' + missing.length + ')'); h += tbl_(['ตัวชี้วัด', 'หน่วย', 'เดือนที่ขาด'], missing.slice(0, 20)); }

  const url = appUrl_();
  if (url) h += '<p style="margin:22px 0 6px"><a href="' + esc_(url) + '" style="display:inline-block;background:#0b3d91;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600">เปิดหน้าสรุปผู้บริหาร</a></p>';
  h += '<p style="font-size:11.5px;color:#8a9a97;margin-top:18px">อีเมลนี้ส่งอัตโนมัติจากระบบรายงาน KPI · ข้อมูล ณ วันที่ออกรายงาน · ต้องการเปลี่ยนผู้รับหรือยกเลิก ติดต่อผู้ดูแลระบบ</p></div>';
  return { subject: '[รายงาน KPI] ' + org + ' ' + P.label, html: h };
}

function reportRecipients_() {
  const list = [];
  function add(e) { e = String(e || '').trim().toLowerCase(); if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e) && list.indexOf(e) < 0) list.push(e); }
  String(settingVal_('rpt_to') || '').split(/[,;\s]+/).forEach(add);
  const ex = settingVal_('rpt_exec');
  if (ex === null || isTrue_(ex)) read_('Users').forEach(function (u) { if (u.role === 'exec' && isActive_(u.active)) add(u.email); });
  return list;
}

function sendReport_(kind, who, to) {
  const r = buildReport_(kind);
  to = to || reportRecipients_();
  if (!to.length) throw new Error('ยังไม่มีผู้รับรายงาน (เพิ่มผู้ใช้บทบาทผู้บริหาร หรือใส่อีเมลผู้รับในหน้าผู้ดูแลระบบ)');
  MailApp.sendEmail({ to: to.join(','), subject: r.subject, htmlBody: r.html, name: 'ระบบรายงาน KPI' });
  setSetting_('rpt_last_' + kind, new Date().toISOString() + '|' + to.length);
  audit_({ email: who || 'ระบบ (ตั้งเวลา)', role: 'admin' }, [{ action: 'ส่งรายงานอีเมล', sheet: 'Settings', key: kind, field: 'to', old: '', new: to.length + ' ผู้รับ' }]);
  return { sent: to.length, to: to, subject: r.subject };
}

/** ดูตัวอย่างรายงาน (ผู้ดูแลระบบ และผู้บริหาร) */
function previewReport(kind) {
  const u = currentUser_();
  if (u.role !== 'admin' && u.role !== 'exec') throw new Error('เฉพาะผู้ดูแลระบบและผู้บริหาร');
  return buildReport_(String(kind));
}

/** ส่งรายงานทันที · test = ส่งถึงตัวเองเท่านั้น */
function sendReportNow(kind, test) {
  const u = requireAdmin_();
  if (test && !u.email) throw new Error('ระบบไม่ทราบอีเมลของคุณ ให้เข้าสู่ระบบด้วยอีเมลและรหัสเข้าใช้ก่อน');
  return sendReport_(String(kind), u.email || u.name, test ? [u.email] : null);
}

function getReportSettings() {
  requireAdmin_();
  const trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).filter(function (h) { return h === 'rptWeeklyJob' || h === 'rptMonthlyJob'; });
  let quota = null;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = null; }
  const ex = settingVal_('rpt_exec');
  return {
    to: String(settingVal_('rpt_to') || ''), exec: ex === null ? true : isTrue_(ex),
    weekly: settingOn_('rpt_weekly'), monthly: settingOn_('rpt_monthly'), quarterly: settingOn_('rpt_quarterly'),
    recipients: reportRecipients_(), triggers: trig, quota: quota,
    last: { week: String(settingVal_('rpt_last_week') || ''), month: String(settingVal_('rpt_last_month') || ''), quarter: String(settingVal_('rpt_last_quarter') || '') }
  };
}

function saveReportSettings(o) {
  const u = requireAdmin_();
  o = o || {};
  const to = String(o.to || '').trim();
  const bad = to.split(/[,;\s]+/).filter(String).filter(function (e) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
  if (bad.length) throw new Error('อีเมลไม่ถูกต้อง: ' + bad.join(', '));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setSetting_('rpt_to', to);
    setSetting_('rpt_exec', o.exec ? '1' : '0');
    setSetting_('rpt_weekly', o.weekly ? '1' : '0');
    setSetting_('rpt_monthly', o.monthly ? '1' : '0');
    setSetting_('rpt_quarterly', o.quarterly ? '1' : '0');
  } finally { lock.releaseLock(); }
  ensureReportTriggers_(!!(o.weekly || o.monthly || o.quarterly));
  audit_(u, [{ action: 'ตั้งค่ารายงานอีเมล', sheet: 'Settings', key: 'rpt', field: '', old: '',
    new: [o.weekly ? 'สัปดาห์' : '', o.monthly ? 'เดือน' : '', o.quarterly ? 'ไตรมาส' : ''].filter(String).join(', ') || 'ปิดทั้งหมด' }]);
  return getReportSettings();
}

function ensureReportTriggers_(on) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['rptWeeklyJob', 'rptMonthlyJob'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  if (!on) return;
  ScriptApp.newTrigger('rptWeeklyJob').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone(TZ).create();
  ScriptApp.newTrigger('rptMonthlyJob').timeBased().onMonthDay(1).atHour(8).inTimezone(TZ).create();
}

function isTriggerCall_(e) {
  return !!(e && e.triggerUid) && ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === String(e.triggerUid); });
}
function rptWeeklyJob(e) {
  if (!isTriggerCall_(e)) throw new Error('เรียกได้จากตัวตั้งเวลาเท่านั้น');
  if (settingOn_('rpt_weekly')) sendReport_('week');
}
function rptMonthlyJob(e) {
  if (!isTriggerCall_(e)) throw new Error('เรียกได้จากตัวตั้งเวลาเท่านั้น');
  if (settingOn_('rpt_monthly')) sendReport_('month');
  const cur = fiscalOf_(todayStr_());
  if (settingOn_('rpt_quarterly') && (cur.month - 1) % 3 === 0) sendReport_('quarter');
}

/** รันหนึ่งครั้งจากหน้าแก้ไขสคริปต์: อนุญาตสิทธิ์ส่งอีเมลและตั้งเวลา */
function authorizeReports() {
  ownerOnly_();
  ScriptApp.getProjectTriggers();
  Logger.log('โควตาอีเมลคงเหลือวันนี้: ' + MailApp.getRemainingDailyQuota());
  return 'อนุญาตแล้ว';
}
