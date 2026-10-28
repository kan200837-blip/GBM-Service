/**
 * ระบบรายงาน KPI ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม
 * Google Apps Script Web App + Google Sheets เป็นฐานข้อมูล
 *
 * ขั้นตอนแรก: เปิดไฟล์นี้แล้วรันฟังก์ชัน setup() หนึ่งครั้ง
 * จากนั้นเลือก Deploy > New deployment > Web app
 *
 * บทบาทผู้ใช้
 *   admin   ผู้ดูแลระบบ: ทำได้ทุกอย่าง จัดการผู้ใช้ ตัวชี้วัด ล็อกเดือน ดูประวัติ
 *   editor  ผู้บันทึกข้อมูล: บันทึก/แก้ไขเฉพาะ KPI ของหน่วยที่ได้รับสิทธิ์ ในเดือนที่ยังไม่ล็อก
 *   viewer  ผู้ดู: ดูรายงานได้อย่างเดียว (ค่าเริ่มต้นของทุกคนที่ไม่อยู่ในรายชื่อ)
 * เจ้าของสคริปต์ (บัญชีที่ Deploy) เป็นผู้ดูแลระบบเสมอ
 */

const APP_TITLE = 'KPI ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม';
const DB_PROP = 'KPI_DB_ID';
const PM_UNIT = 'ซ่อมบำรุง';
const ROLES = ['admin', 'editor', 'inspector', 'viewer'];
const DIRS = ['ge', 'gt', 'lt', 'zero', 'le_year', 'info'];
const MASTER_FIELDS = ['name', 'group', 'unit', 'target_text', 'kind', 'dir', 'thr', 'max100', 'num_label', 'den_label', 'order', 'active', 'source'];
const ITEM_FIELDS = ['name', 'discipline', 'location', 'uom', 'responsible_name', 'responsible_email', 'kpi_id', 'contractor', 'active', 'note'];

const SHEETS = {
  KPI_Master: ['kpi_id', 'name', 'group', 'unit', 'target_text', 'kind', 'dir', 'thr', 'max100', 'num_label', 'den_label', 'order', 'active', 'source'],
  KPI_Targets: ['fy', 'kpi_id', 'dir', 'thr', 'target_text', 'note', 'updated_by', 'updated_at'],
  KPI_Data: ['fy', 'month', 'kpi_id', 'num', 'den', 'note', 'updated_by', 'updated_at'],
  KPI_Quarter: ['fy', 'quarter', 'kpi_id', 'value', 'note', 'updated_by', 'updated_at'],
  Workload: ['fy', 'quarter', 'work_type', 'count', 'note'],
  PM_Systems: ['sys_code', 'name', 'discipline', 'active', 'responsible_name', 'responsible_email'],
  PM_Items: ['code', 'name', 'discipline', 'location', 'uom', 'responsible_name', 'responsible_email', 'kpi_id', 'contractor', 'active', 'note'],
  PM_Plan: ['fy', 'month', 'code', 'plan', 'actual', 'note', 'updated_by', 'updated_at'],
  Staff: ['name', 'email', 'duty', 'active', 'staff_code', 'first_name', 'last_name', 'can_login'],
  Insights: ['fy', 'order', 'type', 'title', 'text'],
  DQ_Issues: ['no', 'severity', 'where', 'issue', 'impact', 'fix', 'status'],
  Users: ['email', 'name', 'role', 'units', 'active', 'insp_types', 'insp_groups', 'staff_code'],
  Insp_Resp: ['staff_code', 'type_id', 'group', 'note'],
  Insp_Types: ['type_id', 'name', 'freq', 'pm_code', 'kpi_id', 'discipline', 'owner_name', 'owner_email', 'active', 'order', 'pm_sync', 'min_cov', 'note', 'standard'],
  Insp_Checks: ['type_id', 'check_id', 'label', 'kind', 'soft', 'min', 'max', 'tol', 'unit', 'crit', 'std', 'order'],
  Insp_Assets: ['asset_id', 'type_id', 'name', 'location', 'code', 'group', 'active', 'since', 'until', 'responsible_name', 'responsible_email', 'note'],
  Insp_Log: ['log_id', 'type_id', 'asset_id', 'date', 'period', 'inspector', 'st', 'ng_n', 'wn_n', 'results', 'note', 'source', 'created_by', 'created_at'],
  Auto_Override: ['target', 'key', 'by', 'at', 'note'],
  Insp_Actions: ['action_id', 'type_id', 'asset_id', 'check_id', 'issue', 'severity', 'found_date', 'last_seen', 'seen_count', 'assignee', 'assignee_email', 'due_date', 'status', 'action_note', 'done_date', 'verified_ok_date', 'updated_by', 'updated_at'],
  Locks: ['fy', 'month', 'locked_by', 'locked_at'],
  Audit_Log: ['time', 'user', 'role', 'action', 'sheet', 'key', 'field', 'old', 'new'],
  Settings: ['key', 'value']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** รันครั้งเดียวเพื่อสร้าง Google Sheets ฐานข้อมูลพร้อมข้อมูลตั้งต้นจาก Seed.gs */
function setup() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(DB_PROP);
  if (id) {
    const old = SpreadsheetApp.openById(id);
    Logger.log('พบฐานข้อมูลเดิมแล้ว: ' + old.getUrl());
    return old.getUrl();
  }
  const ss = SpreadsheetApp.create('KPI_FMD_ฐานข้อมูล');
  props.setProperty(DB_PROP, ss.getId());
  ensureSheets_(ss);
  const now = new Date().toISOString();
  Object.keys(SEED).forEach(function (name) { put_(ss, name, seedRows_(name, now)); });
  const owner = Session.getEffectiveUser().getEmail();
  put_(ss, 'Settings', [['org', 'ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม'], ['pm_seeded', '1']]);
  put_(ss, 'Users', [[owner, 'เจ้าของระบบ', 'admin', '*', 1]]);
  put_(ss, 'Audit_Log', [[now, owner, 'admin', 'สร้างระบบ', 'setup', '', '', '', 'ข้อมูลตั้งต้นจากไฟล์ KPI 2568 และแผนบำรุงรักษา 2567–2569']]);
  Logger.log('สร้างฐานข้อมูลแล้ว: ' + ss.getUrl());
  return ss.getUrl();
}

/**
 * ใช้ Google Sheets ที่นำเข้าจากไฟล์ KPI_FMD_Import.xlsx เป็นฐานข้อมูล
 * แก้ EXISTING_ID เป็น ID ของ Google Sheets แล้วรันหนึ่งครั้ง
 * ระบบจะเพิ่มชีต/คอลัมน์ที่ขาด (Locks, Audit_Log, Settings) และเพิ่มเจ้าของเป็นผู้ดูแลถ้ายังไม่มีผู้ดูแล
 */
function connectExistingDb(id) {
  const EXISTING_ID = (typeof id === 'string' && id) ? id : 'วาง ID ของ Google Sheets ที่นี่';
  const ss = SpreadsheetApp.openById(EXISTING_ID);
  PropertiesService.getScriptProperties().setProperty(DB_PROP, ss.getId());
  ensureSheets_(ss);
  const owner = Session.getEffectiveUser().getEmail();
  const users = read_('Users');
  if (!users.some(function (u) { return u.role === 'admin' && isActive_(u.active) && String(u.email).indexOf('@') > 0 && String(u.email).indexOf('ใส่') !== 0; })) {
    ss.getSheetByName('Users').appendRow([owner, 'เจ้าของระบบ', 'admin', '*', 1]);
  }
  if (!read_('Settings').length) put_(ss, 'Settings', [['org', 'ฝ่ายบริหารอาคารสถานที่และสิ่งแวดล้อม']]);
  audit_({ email: owner, role: 'admin' }, [{ action: 'เชื่อมฐานข้อมูล', sheet: 'setup', key: ss.getId(), field: '', old: '', new: 'นำเข้าจาก Excel' }]);
  Logger.log('เชื่อมฐานข้อมูลแล้ว: ' + ss.getUrl());
  return ss.getUrl();
}

/** สร้างชีตและคอลัมน์ที่ขาดให้ครบตาม SHEETS (ไม่ลบหรือเปลี่ยนข้อมูลเดิม) */
function ensureSheets_(ss) {
  Object.keys(SHEETS).forEach(function (name) {
    const cols = SHEETS[name];
    let sh = ss.getSheetByName(name);
    if (!sh) {
      const sheets = ss.getSheets();
      sh = (sheets.length === 1 && sheets[0].getLastRow() === 0 && sheets[0].getName().indexOf('Sheet') === 0) ? sheets[0].setName(name) : ss.insertSheet(name);
    }
    const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    cols.forEach(function (c) {
      if (have.indexOf(c) < 0) { have.push(c); sh.getRange(1, have.length).setValue(c); }
    });
    sh.getRange(1, 1, 1, have.length).setFontWeight('bold').setBackground('#e6efec');
    sh.setFrozenRows(1);
    // คอลัมน์รหัส/วันที่ของงานตรวจเช็คเก็บเป็นข้อความ กัน Google Sheets แปลงเป็นวันที่/ตัวเลข (รหัส 18 หลักจะเพี้ยน)
    (typeof INSP_TEXT_COLS !== 'undefined' && INSP_TEXT_COLS[name] || []).forEach(function (c) {
      const ci = have.indexOf(c);
      if (ci >= 0) sh.getRange(1, ci + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
}

function put_(ss, name, rows) {
  if (!rows.length) return;
  ss.getSheetByName(name).getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

var DB_READY_ = '';

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(DB_PROP);
  if (!id) throw new Error('ยังไม่ได้รัน setup() เพื่อสร้างฐานข้อมูล');
  const ss = SpreadsheetApp.openById(id);
  if (DB_READY_ !== id) { DB_READY_ = id; upgradeDb_(ss); }
  return ss;
}

/** แถวข้อมูลตั้งต้นของชีต เติมคอลัมน์ให้ครบและใส่ผู้บันทึก/เวลา */
function seedRows_(name, now) {
  const head = SHEETS[name];
  return (SEED[name] || []).map(function (r) {
    const line = r.slice();
    while (line.length < head.length) line.push('');
    if (head.indexOf('updated_by') >= 0) { line[head.indexOf('updated_by')] = 'setup'; line[head.indexOf('updated_at')] = now; }
    return line;
  });
}

function seedIsNew_() {
  return typeof SEED !== 'undefined' && !!SEED.KPI_Master && !!SEED.PM_Items && !!SEED.PM_Plan;
}

function headOf_(sh) {
  return sh && sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
}

/**
 * ปรับฐานข้อมูลเดิม (สร้างจากโค้ดรุ่นก่อน) ให้ตรงกับโค้ดปัจจุบันโดยอัตโนมัติ ทำงานเมื่อพบว่าขาดชีต
 * - สร้างชีตที่ขาดและใส่ข้อมูลตั้งต้นจาก Seed.gs
 * - แผนบำรุงรักษารุ่นเก่า (ใช้คอลัมน์ discipline) เปลี่ยนชื่อเป็น PM_Plan_เก่า แล้วสร้างใหม่
 * - ถ้าเพิ่งเพิ่ม KPI_Quarter: ย้ายข้อมูลรายเดือน K1–K6 ปี 2569 ไปชีต KPI_Data_ก่อนแทนที่
 *   และเปลี่ยน Insights/DQ_Issues เป็นชุดใหม่ (ชุดเดิมเก็บไว้ในชีต *_เก่า)
 * ไม่ลบข้อมูลเดิมทิ้ง และรันซ้ำได้ (เรียกเองได้จากเมนู Run > upgradeDb)
 */
function upgradeDb() { return upgradeDb_(db_()); }

function upgradeDb_(ss) {
  const a = upgradeBase_(ss);
  // คอลัมน์ใหม่ (บทบาทผู้ตรวจเช็ค / ผู้รับผิดชอบระบบ) ในฐานข้อมูลเดิม
  if (headOf_(ss.getSheetByName('Users')).indexOf('insp_groups') < 0 || headOf_(ss.getSheetByName('PM_Systems')).indexOf('responsible_email') < 0) ensureSheets_(ss);
  const b = inspSeed_(ss);
  const c = inspEnsureLinkItems_(ss) > 0;
  return a || b || c;
}

function upgradeBase_(ss) {
  const names = ss.getSheets().map(function (x) { return x.getName(); });
  const missing = Object.keys(SHEETS).filter(function (n) { return names.indexOf(n) < 0; });
  const oldPlan = names.indexOf('PM_Plan') >= 0 && headOf_(ss.getSheetByName('PM_Plan')).indexOf('discipline') >= 0;
  const oldMaster = names.indexOf('KPI_Master') >= 0 && headOf_(ss.getSheetByName('KPI_Master')).indexOf('source') < 0;
  // มีชีตแผนบำรุงรักษาแต่ว่างเปล่า และยังไม่เคยใส่ข้อมูลตั้งต้น (เช่น เคยอัปเกรดตอนที่ Seed.gs ยังเป็นรุ่นเก่า)
  let emptyPlan = false;
  if (!missing.length && names.indexOf('PM_Items') >= 0 && names.indexOf('Settings') >= 0 && ss.getSheetByName('PM_Items').getLastRow() < 2) {
    const flag = ss.getSheetByName('Settings').getDataRange().getValues().some(function (r) { return r[0] === 'pm_seeded'; });
    emptyPlan = !flag;
  }
  if (!missing.length && !oldPlan && !oldMaster && !emptyPlan) return false;
  if (!seedIsNew_()) {
    throw new Error('ไฟล์ Seed.gs ยังเป็นรุ่นเก่า (ไม่มีข้อมูลแผนบำรุงรักษา รหัส 7 หลัก) กรุณาเปิดไฟล์ Seed.gs ใน Apps Script ลบเนื้อหาเดิมแล้ววางไฟล์ Seed.gs รุ่นใหม่ จากนั้น Deploy > Manage deployments > Edit > New version แล้วรีเฟรช');
  }
  const now = new Date().toISOString();
  const notes = [];
  if (oldPlan) {
    ss.getSheetByName('PM_Plan').setName('PM_Plan_เก่า');
    missing.push('PM_Plan');
    notes.push('PM_Plan รุ่นเก่า > PM_Plan_เก่า');
  }
  ensureSheets_(ss);
  const fresh = missing.indexOf('KPI_Quarter') >= 0;
  if (fresh) {
    ['Insights', 'DQ_Issues'].forEach(function (n) {
      if (missing.indexOf(n) < 0 && ss.getSheetByName(n)) {
        ss.getSheetByName(n).setName(n + '_เก่า');
        missing.push(n);
        notes.push(n + ' > ' + n + '_เก่า');
      }
    });
    ensureSheets_(ss);
  }
  if (emptyPlan) {
    ['KPI_Targets', 'KPI_Quarter', 'Workload', 'PM_Systems', 'PM_Items', 'PM_Plan', 'Staff', 'Insights', 'DQ_Issues'].forEach(function (n) {
      if (ss.getSheetByName(n).getLastRow() < 2 && missing.indexOf(n) < 0) missing.push(n);
    });
    notes.push('เติมข้อมูลแผนบำรุงรักษาที่ว่างอยู่');
  }
  missing.forEach(function (n) {
    if (['Users', 'Settings', 'Audit_Log', 'Locks'].indexOf(n) >= 0) return;
    if (SEED[n]) put_(ss, n, seedRows_(n, now));
  });
  if (SEED.PM_Items && !ss.getSheetByName('Settings').getDataRange().getValues().some(function (r) { return r[0] === 'pm_seeded'; })) {
    ss.getSheetByName('Settings').appendRow(['pm_seeded', '1']);
  }
  // เพิ่มตัวชี้วัดที่ยังไม่มีในตารางเดิม (จับคู่ตามชื่อคอลัมน์)
  const ms = ss.getSheetByName('KPI_Master'), mh = headOf_(ms);
  const have = {};
  ms.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[mh.indexOf('kpi_id')])] = true; });
  const add = (SEED.KPI_Master || []).filter(function (r) { return !have[String(r[0])]; }).map(function (r) {
    return mh.map(function (h) { const i = SHEETS.KPI_Master.indexOf(h); return i >= 0 && r[i] !== undefined ? r[i] : (h === 'source' ? 'manual' : ''); });
  });
  if (add.length) { ms.getRange(ms.getLastRow() + 1, 1, add.length, mh.length).setValues(add); notes.push('เพิ่มตัวชี้วัด ' + add.length); }
  if (fresh) {
    // ข้อมูลรายเดือน K1–K6 ปี 2569 เดิมขัดกับรายงานรายไตรมาส: ย้ายไปเก็บไว้
    const ds = ss.getSheetByName('KPI_Data'), dv = ds.getDataRange().getValues(), dh = dv[0];
    const iFy = dh.indexOf('fy'), iK = dh.indexOf('kpi_id');
    const gone = [];
    for (let i = 1; i < dv.length; i++) {
      if (Number(dv[i][iFy]) === 2569 && ['K1', 'K2', 'K3', 'K4', 'K5', 'K6'].indexOf(String(dv[i][iK])) >= 0) gone.push(i + 1);
    }
    if (gone.length) {
      const ar = ss.getSheetByName('KPI_Data_ก่อนแทนที่') || ss.insertSheet('KPI_Data_ก่อนแทนที่');
      ar.getRange(1, 1, 1, dh.length).setValues([dh]);
      ar.getRange(2, 1, gone.length, dh.length).setValues(gone.map(function (n) { return dv[n - 1]; }));
      gone.slice().reverse().forEach(function (n) { ds.deleteRow(n); });
      notes.push('ย้ายข้อมูลรายเดือน K1–K6 ปี 2569 ' + gone.length + ' แถว');
    }
  }
  if (!ss.getSheetByName('Users') || ss.getSheetByName('Users').getLastRow() < 2) {
    ss.getSheetByName('Users').appendRow([Session.getEffectiveUser().getEmail(), 'เจ้าของระบบ', 'admin', '*', 1]);
  }
  const al = ss.getSheetByName('Audit_Log');
  al.appendRow([now, Session.getEffectiveUser().getEmail(), 'admin', 'อัปเกรดฐานข้อมูล', 'setup', '', '', '', 'สร้างชีต: ' + missing.join(', ') + (notes.length ? ' · ' + notes.join(' · ') : '')]);
  return true;
}

function read_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (!v.length) return [];
  const head = v.shift();
  return v.filter(function (r) { return r.some(function (c) { return c !== ''; }); }).map(function (r) {
    const o = {};
    head.forEach(function (k, i) {
      const x = r[i];
      o[k] = x === '' ? null : (x instanceof Date ? x.toISOString() : x);
    });
    return o;
  });
}

/* ---------------- ผู้ใช้และสิทธิ์ ---------------- */

function isActive_(v) {
  return !(v === 0 || v === false || String(v).toLowerCase() === 'false' || String(v) === '0' || v === 'ไม่');
}

function parseUnits_(s) {
  return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(String);
}

function currentUser_() {
  if (typeof STAFF_CTX_ !== 'undefined' && STAFF_CTX_) return STAFF_CTX_; // เข้าด้วยรหัสเจ้าหน้าที่ + ชื่อ-นามสกุล (ผ่าน api)
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const users = read_('Users');
  let u = null;
  if (email) {
    u = users.filter(function (x) { return String(x.email || '').toLowerCase() === email && isActive_(x.active); })[0];
  }
  if (u && ROLES.indexOf(u.role) >= 0) {
    const lst = function (x) { return String(x || '').split(',').map(function (y) { return y.trim(); }).filter(String); };
    const rs = inspRespOf_(email, u), uniq = function (k) { const m = {}; rs.forEach(function (r) { m[r[k]] = 1; }); return m['*'] ? [] : Object.keys(m); };
    return { email: email, name: u.name || email, role: u.role, units: parseUnits_(u.units), inspTypes: rs.length ? uniq('type_id') : lst(u.insp_types), inspGroups: rs.length ? uniq('group') : lst(u.insp_groups), inspResp: rs, staffCode: String(u.staff_code || '').trim() };
  }
  if (email && email === owner) return { email: email, name: 'เจ้าของระบบ', role: 'admin', units: ['*'], inspTypes: [], inspGroups: [], inspResp: [] };
  return { email: email, name: email || 'ผู้เข้าชม', role: 'viewer', units: [], inspTypes: [], inspGroups: [], inspResp: [] };
}

function canEditUnit_(u, unit) {
  if (u.role === 'admin') return true;
  return u.role === 'editor' && (u.units.indexOf('*') >= 0 || u.units.indexOf(unit) >= 0);
}

function requireAdmin_() {
  const u = currentUser_();
  if (u.role !== 'admin') throw new Error('ต้องเป็นผู้ดูแลระบบ');
  return u;
}

function requireEditor_(allowInspector) {
  const u = currentUser_();
  if (u.role === 'viewer') throw new Error('บัญชีนี้ดูได้อย่างเดียว ให้ผู้ดูแลระบบเพิ่มสิทธิ์ในแท็บผู้ดูแลระบบ');
  if (u.role === 'inspector' && !allowInspector) throw new Error('บทบาทผู้ตรวจเช็คบันทึกได้เฉพาะผลตรวจเช็คและงานแก้ไข');
  return u;
}

function lockSet_() {
  const s = {};
  read_('Locks').forEach(function (l) { s[Number(l.fy) + '|' + Number(l.month)] = true; });
  return s;
}

/* ---------------- ประวัติการแก้ไข ---------------- */

function audit_(user, entries) {
  if (!entries.length) return;
  const sh = db_().getSheetByName('Audit_Log');
  const now = new Date().toISOString();
  const rows = entries.map(function (e) {
    return [now, user.email || user.name, user.role, e.action, e.sheet, e.key || '', e.field || '', e.old === undefined ? '' : e.old, e.new === undefined ? '' : e.new];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
}

function norm_(x) {
  if (x === null || x === undefined || x === '') return '';
  const n = Number(x);
  return (typeof x !== 'boolean' && String(x).trim() !== '' && !isNaN(n)) ? String(n) : String(x);
}

/* ---------------- อ่านข้อมูล ---------------- */

function appUrl_() {
  try { return typeof ScriptApp !== 'undefined' ? String(ScriptApp.getService().getUrl() || '') : ''; } catch (e) { return ''; }
}

/** ข้อมูลทั้งหมดที่หน้าเว็บใช้คำนวณรายงาน */
/** ใช้ตอนเปิดหน้าเว็บ: บอกบทบาทของบัญชี Google ที่เปิดอยู่ (ไม่ส่งข้อมูลอื่น) */
function whoAmI() {
  const u = currentUser_();
  return { role: u.role, email: u.email, name: u.name };
}

function getAll() {
  const settings = {};
  read_('Settings').forEach(function (s) { settings[s.key] = s.value; });
  const user = currentUser_();
  const master = read_('KPI_Master').filter(function (k) { return isActive_(k.active); })
    .sort(function (a, b) { return Number(a.order) - Number(b.order); });
  const lite = user.role === 'inspector'; // ผู้ตรวจเช็คเห็นเฉพาะงานตรวจเช็ค ไม่ส่งตัวเลข KPI ไปที่เบราว์เซอร์
  return {
    master: master,
    data: lite ? [] : read_('KPI_Data'),
    targets: lite ? [] : read_('KPI_Targets'),
    quarter: lite ? [] : read_('KPI_Quarter'),
    workload: lite ? [] : read_('Workload'),
    systems: read_('PM_Systems'),
    items: read_('PM_Items'),
    staff: lite ? read_('Staff').map(function (s) { return { name: s.name, duty: s.duty, active: s.active, email: '' }; }) : read_('Staff'),
    plan: read_('PM_Plan'),
    insights: lite ? [] : read_('Insights'),
    issues: lite ? [] : read_('DQ_Issues'),
    locks: read_('Locks').map(function (l) { return { fy: Number(l.fy), month: Number(l.month), by: l.locked_by }; }),
    settings: settings,
    user: user,
    pmUnit: PM_UNIT,
    appUrl: appUrl_(),
    insp: inspBundle_(),
    dbUrl: user.role === 'admin' ? db_().getUrl() : ''
  };
}

/** ข้อมูลสำหรับหน้าผู้ดูแลระบบ */
function getAdminData() {
  requireAdmin_();
  // ผ่าน JSON เพื่อกันค่าที่ google.script.run ส่งไม่ได้ (เช่น Date) ซึ่งจะทำให้ผลลัพธ์กลายเป็น null ทั้งก้อน
  return JSON.parse(JSON.stringify({
    users: read_('Users'),
    locks: read_('Locks'),
    master: read_('KPI_Master'),
    staff: read_('Staff'),
    audit: readAudit_(300)
  }));
}

function readAudit_(limit) {
  const sh = db_().getSheetByName('Audit_Log');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit, last - 1);
  const v = sh.getRange(last - n + 1, 1, n, 9).getValues();
  return v.reverse().map(function (r) {
    const s = function (x) { return x instanceof Date ? x.toISOString() : (x === null || x === undefined ? '' : x); };
    return { time: s(r[0]), user: s(r[1]), role: s(r[2]), action: s(r[3]), sheet: s(r[4]), key: s(r[5]), field: s(r[6]), old: s(r[7]), new: s(r[8]) };
  });
}

function getAudit(limit) {
  requireAdmin_();
  return readAudit_(Math.min(Number(limit) || 300, 2000));
}

/* ---------------- บันทึกข้อมูล ---------------- */

function checkNumber_(label, v) {
  if (v === null || v === undefined || v === '') return;
  const n = Number(v);
  if (isNaN(n)) throw new Error(label + ' ต้องเป็นตัวเลข');
  if (n < 0) throw new Error(label + ' ต้องไม่ติดลบ');
}

/** บันทึก/แก้ไข/ลบ ข้อมูล KPI รายเดือน rows: [{fy, month, kpi_id, num, den, note}] */
function saveKpiRows(rows) {
  const u = requireEditor_();
  const unitOf = {}, srcOf = {}, ovK = [];
  read_('KPI_Master').forEach(function (k) { unitOf[k.kpi_id] = k.unit; srcOf[k.kpi_id] = k.source; });
  const locks = lockSet_();
  const res = upsert_('KPI_Data', ['fy', 'month', 'kpi_id'], rows, ['num', 'den', 'note'], {
    user: u,
    allowDelete: true,
    validate: function (r) {
      if (!unitOf[r.kpi_id]) throw new Error('ไม่พบตัวชี้วัด ' + r.kpi_id);
      if (!(r.month >= 1 && r.month <= 12)) throw new Error('เดือนไม่ถูกต้อง');
      checkNumber_('ตัวตั้ง/จำนวน (' + r.kpi_id + ')', r.num);
      checkNumber_('ตัวหาร (' + r.kpi_id + ')', r.den);
    },
    check: function (r) {
      if (srcOf[r.kpi_id] === 'pm') return 'ตัวชี้วัด ' + r.kpi_id + ' คำนวณจากแผนบำรุงรักษา ไม่ต้องกรอกผลเอง (แก้ที่แท็บแผนบำรุงรักษา)';
      if (srcOf[r.kpi_id] === 'insp' && inspKpiOwns_(r.fy, r.month)) {
        if (u.role !== 'admin') return 'ตัวชี้วัด ' + r.kpi_id + ' คำนวณจากผลตรวจเช็คอัตโนมัติ แก้ได้เฉพาะผู้ดูแลระบบ (บันทึกผลที่แท็บตรวจเช็คบำรุงรักษา / แผนบำรุงรักษา)';
        ovK.push({ target: 'KPI_Data', key: Number(r.fy) + '|' + Number(r.month) + '|' + r.kpi_id, note: 'ผู้ดูแลระบบแก้ไขผล KPI' });
      }
      if (!canEditUnit_(u, unitOf[r.kpi_id])) return 'ไม่มีสิทธิ์แก้ไขตัวชี้วัด ' + r.kpi_id + ' (หน่วย ' + unitOf[r.kpi_id] + ')';
      if (u.role !== 'admin' && locks[Number(r.fy) + '|' + Number(r.month)]) return 'เดือนนี้ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อนแก้ไข';
      return null;
    }
  });
  inspAddOverrides_(u, ovK);
  return res;
}

/**
 * บันทึกผลตัวชี้วัดที่รายงานเป็นรายไตรมาส (กรณีไม่มีข้อมูลรายเดือน) rows: [{fy, quarter, kpi_id, value, note}]
 * ใช้เฉพาะเดือนที่ยังไม่มีข้อมูลรายเดือนของ KPI นั้น
 */
function saveQuarterRows(rows) {
  const u = requireEditor_();
  const unitOf = {}, srcOf = {};
  read_('KPI_Master').forEach(function (k) { unitOf[k.kpi_id] = k.unit; srcOf[k.kpi_id] = k.source; });
  const locks = lockSet_();
  return upsert_('KPI_Quarter', ['fy', 'quarter', 'kpi_id'], rows, ['value', 'note'], {
    user: u,
    allowDelete: true,
    validate: function (r) {
      if (!unitOf[r.kpi_id]) throw new Error('ไม่พบตัวชี้วัด ' + r.kpi_id);
      if (!(Number(r.quarter) >= 1 && Number(r.quarter) <= 4)) throw new Error('ไตรมาสไม่ถูกต้อง');
      checkNumber_('ผลรายไตรมาส (' + r.kpi_id + ')', r.value);
    },
    check: function (r) {
      if (srcOf[r.kpi_id] === 'pm') return 'ตัวชี้วัด ' + r.kpi_id + ' คำนวณจากแผนบำรุงรักษา ไม่ต้องกรอกผลเอง';
      if (srcOf[r.kpi_id] === 'insp' && inspKpiOwns_(r.fy, (Number(r.quarter) - 1) * 3 + 1) && u.role !== 'admin') return 'ตัวชี้วัด ' + r.kpi_id + ' คำนวณจากผลตรวจเช็คอัตโนมัติ แก้ได้เฉพาะผู้ดูแลระบบ';
      if (!canEditUnit_(u, unitOf[r.kpi_id])) return 'ไม่มีสิทธิ์แก้ไขตัวชี้วัด ' + r.kpi_id + ' (หน่วย ' + unitOf[r.kpi_id] + ')';
      if (u.role !== 'admin') {
        for (let m = (Number(r.quarter) - 1) * 3 + 1; m <= Number(r.quarter) * 3; m++) {
          if (locks[Number(r.fy) + '|' + m]) return 'ไตรมาสนี้มีเดือนที่ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อนแก้ไข';
        }
      }
      return null;
    }
  });
}

/**
 * บันทึกปริมาณงานรายไตรมาสที่เป็นสาเหตุของงานซ่อมไม่ทัน (เช่น งานจัดสถานที่ งานออกหน่วยแพทย์เคลื่อนที่)
 * rows: [{fy, quarter, work_type, count, note}]  ล้างจำนวนแล้วบันทึกเพื่อลบ
 */
function saveWorkloadRows(rows) {
  const u = requireEditor_();
  const locks = lockSet_();
  return upsert_('Workload', ['fy', 'quarter', 'work_type'], rows, ['count', 'note'], {
    user: u,
    allowDelete: true,
    validate: function (r) {
      if (!String(r.work_type || '').trim()) throw new Error('ต้องระบุประเภทงาน');
      if (String(r.work_type).length > 80) throw new Error('ชื่อประเภทงานยาวเกิน 80 ตัวอักษร');
      r.work_type = String(r.work_type).trim();
      if (!(Number(r.quarter) >= 1 && Number(r.quarter) <= 4)) throw new Error('ไตรมาสไม่ถูกต้อง');
      checkNumber_('จำนวนงาน (' + r.work_type + ')', r.count);
    },
    check: function (r) {
      if (!canEditUnit_(u, PM_UNIT)) return 'ไม่มีสิทธิ์บันทึกปริมาณงาน (หน่วย ' + PM_UNIT + ')';
      if (u.role !== 'admin') {
        for (let m = (Number(r.quarter) - 1) * 3 + 1; m <= Number(r.quarter) * 3; m++) {
          if (locks[Number(r.fy) + '|' + m]) return 'ไตรมาสนี้มีเดือนที่ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อนแก้ไข';
        }
      }
      return null;
    }
  });
}

/** ผู้บันทึกแก้ผลจริงของรายการได้ถ้าได้สิทธิ์หน่วยซ่อมบำรุง หรือเป็นผู้รับผิดชอบรายการนั้น (อีเมลตรงกับ PM_Items) */
function canEditItem_(u, item) {
  if (u.role === 'admin') return true;
  if (u.role !== 'editor') return false;
  if (canEditUnit_(u, PM_UNIT)) return true;
  const em = String(u.email || '').toLowerCase();
  if (!!item && !!item.responsible_email && String(item.responsible_email).trim().toLowerCase() === em) return true;
  // ผู้รับผิดชอบระบบ (4 หลักแรกของรหัส) บันทึกได้ทุกรายการในระบบนั้น
  if (item && em) { const sys = read_('PM_Systems').filter(function (x) { return norm_(x.sys_code) === String(item.code).slice(0, 4); })[0]; if (sys && String(sys.responsible_email || '').trim().toLowerCase() === em) return true; }
  return false;
}

/**
 * บันทึกแผน/ผลจริงรายรหัส rows: [{fy, month, code, plan, actual, note}]
 * กำหนดค่าแผน (plan) ได้เฉพาะผู้ดูแลระบบ ส่วนผลจริง (actual) บันทึกโดยผู้รับผิดชอบรายการ
 */
function savePlanRows(rows) {
  const u = requireEditor_();
  // ช่องผลจริงของรหัสที่รับผลจากระบบตรวจเช็ค: ผู้ใช้ทั่วไปแก้ไม่ได้ ผู้ดูแลระบบแก้/ลบได้ และระบบจะไม่คำนวณทับช่องนั้น
  const linked = inspLinkedCodes_(), from = inspKpiFrom_(), cur = {}, ovs = [];
  read_('PM_Plan').forEach(function (r) { cur[Number(r.fy) + '|' + Number(r.month) + '|' + norm_(r.code)] = r; });
  (rows || []).forEach(function (r) {
    const code = norm_(r.code), k = Number(r.fy) + '|' + Number(r.month) + '|' + code;
    if (!linked[code] || !inspKpiOwns_(r.fy, r.month, from)) return;
    const o = cur[k] || {};
    const chg = norm_(o.actual) !== norm_(r.actual) || (r.plan !== undefined && norm_(o.plan) !== norm_(r.plan));
    if (!chg) return;
    if (u.role !== 'admin') throw new Error('รายการ ' + code + ' รับผลจริงจากระบบตรวจเช็คอัตโนมัติ แก้ได้เฉพาะผู้ดูแลระบบ (บันทึกผลที่แท็บ "ตรวจเช็คบำรุงรักษา")');
    ovs.push({ target: 'PM_Plan', key: k, note: 'ผู้ดูแลระบบแก้ไขในแผน: ' + norm_(o.actual) + ' → ' + norm_(r.actual) });
  });
  const r0 = savePlanRows_(rows);
  inspAddOverrides_(u, ovs);
  inspKpiAfterPlan_(rows);
  return r0;
}
/** ผลจริงของแผนที่ผูก KPI ซึ่งคำนวณจากผลตรวจเช็ค เปลี่ยน -> คำนวณ KPI เดือนนั้นใหม่ */
function inspKpiAfterPlan_(rows) {
  const ids = inspKpiIds_(); if (!ids.length) return;
  const kOf = {}; read_('PM_Items').forEach(function (i) { kOf[norm_(i.code)] = String(i.kpi_id || ''); });
  const by = {};
  (rows || []).forEach(function (r) { if (ids.indexOf(kOf[norm_(r.code)]) >= 0) { const f = Number(r.fy); (by[f] = by[f] || []).push(Number(r.month)); } });
  Object.keys(by).forEach(function (f) { syncInspKpiSafe_(Number(f), by[f], 'ระบบตรวจเช็ค (จากแผน PM)'); });
}
function savePlanRows_(rows) {
  const u = requireEditor_();
  const locks = lockSet_();
  const itemOf = {};
  read_('PM_Items').forEach(function (i) { itemOf[norm_(i.code)] = i; });
  return upsert_('PM_Plan', ['fy', 'month', 'code'], rows, ['plan', 'actual', 'note'], {
    user: u,
    allowDelete: true,
    validate: function (r) {
      r.code = norm_(r.code);
      if (!itemOf[r.code]) throw new Error('ไม่พบรหัสรายการ ' + r.code);
      if (!(r.month >= 1 && r.month <= 12)) throw new Error('เดือนไม่ถูกต้อง');
      checkNumber_('แผน (' + r.code + ')', r.plan);
      checkNumber_('ผลจริง (' + r.code + ')', r.actual);
    },
    check: function (r, changed) {
      const planTouched = changed.indexOf('plan') >= 0;
      if (planTouched && u.role !== 'admin') return 'กำหนดค่าแผนได้เฉพาะผู้ดูแลระบบ (รหัส ' + r.code + ')';
      if (!canEditItem_(u, itemOf[r.code])) return 'ไม่มีสิทธิ์บันทึกรายการ ' + r.code + ' (เฉพาะผู้รับผิดชอบหรือผู้ได้รับสิทธิ์หน่วย ' + PM_UNIT + ')';
      if (u.role !== 'admin' && locks[Number(r.fy) + '|' + Number(r.month)]) return 'เดือนนี้ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อนแก้ไข';
      return null;
    }
  });
}

/* ---------------- บันทึกหน้างานด้วยรหัส 7 หลัก ---------------- */

function settingOn_(key) {
  let v = null;
  read_('Settings').forEach(function (r) { if (r.key === key) v = r.value; });
  return v === 1 || v === true || ['1', 'true', 'on', 'เปิด'].indexOf(String(v).toLowerCase()) >= 0;
}

/** ตั้งค่าระบบ (เฉพาะผู้ดูแลระบบ) key ที่อนุญาต: code_entry = เปิดให้บันทึกผลหน้างานด้วยรหัส 7 หลัก · insp_entry = เปิดให้บันทึกผลตรวจเช็คหน้างาน · insp_mail = ส่งอีเมลแจ้งเตือน · insp_sla_high / insp_sla_normal = จำนวนวันกำหนดแก้ไข */
function saveSetting(key, value) {
  const u = requireAdmin_();
  const bools = ['code_entry'].concat(INSP_BOOL_SETTINGS);
  if (bools.indexOf(key) < 0 && INSP_NUM_SETTINGS.indexOf(key) < 0 && key !== 'insp_kpi_from') throw new Error('ไม่รู้จักการตั้งค่า ' + key);
  let val;
  if (key === 'insp_kpi_from') {
    val = String(value || '');
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(val)) throw new Error('เดือนเริ่มต้องเป็นรูปแบบ ปปปป-ดด (ค.ศ.)');
  } else if (INSP_NUM_SETTINGS.indexOf(key) >= 0) {
    const n = Number(value);
    if (!(n >= 1 && n <= 365) || Math.floor(n) !== n) throw new Error('จำนวนวันต้องเป็นจำนวนเต็ม 1–365');
    val = String(n);
  } else val = (value === true || value === 1 || value === '1') ? '1' : '0';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_().getSheetByName('Settings');
    const v = sh.getDataRange().getValues();
    let old = '', row = 0;
    for (let i = 1; i < v.length; i++) { if (v[i][0] === key) { row = i + 1; old = String(v[i][1]); break; } }
    if (row) sh.getRange(row, 2).setValue(val); else sh.appendRow([key, val]);
    if (old !== val) audit_(u, [{ action: 'ตั้งค่า', sheet: 'Settings', key: key, field: 'value', old: old, new: val }]);
  } finally { lock.releaseLock(); }
  return { key: key, value: val };
}

function fiscalNow_() {
  const d = new Date(), m = d.getMonth();
  return { fy: d.getFullYear() + 543 + (m >= 9 ? 1 : 0), month: (m + 3) % 12 + 1 };
}

/**
 * ผู้ปฏิบัติงานบันทึกผลที่ทำจริงของรายการ ด้วยรหัส 7 หลัก  e: {fy, month, code, actual, note, reporter}
 * - ผู้มีสิทธิ์ (ผู้ดูแลระบบ / ผู้บันทึกหน่วยซ่อมบำรุง / ผู้รับผิดชอบรายการ) บันทึกได้เสมอ
 * - คนอื่นบันทึกได้เมื่อผู้ดูแลระบบเปิด code_entry และต้องใส่ชื่อผู้บันทึก
 * - แก้เฉพาะ "ผลที่ทำ" และหมายเหตุ ไม่แตะค่าแผน ต้องมีแผนของเดือนนั้นอยู่แล้ว ไม่บันทึกเดือนอนาคต เดือนที่ล็อกแล้วบันทึกไม่ได้
 */
function saveFieldEntry(e) {
  const u = currentUser_();
  const code = norm_(e && e.code);
  if (!/^\d{7}$/.test(code)) throw new Error('รหัสรายการต้องเป็นตัวเลข 7 หลัก');
  const fy = Number(e.fy), month = Number(e.month);
  if (!fy || !(month >= 1 && month <= 12)) throw new Error('ปีงบหรือเดือนไม่ถูกต้อง');
  if (e.actual === null || e.actual === undefined || String(e.actual).trim() === '') throw new Error('กรุณาใส่ผลที่ทำจริง');
  checkNumber_('ผลที่ทำจริง', e.actual);
  const actual = Number(e.actual);
  if (actual > 100000) throw new Error('ผลที่ทำจริงสูงผิดปกติ ตรวจสอบตัวเลขอีกครั้ง');
  const item = read_('PM_Items').filter(function (i) { return norm_(i.code) === code && isActive_(i.active); })[0];
  if (!item) throw new Error('ไม่พบรหัสรายการ ' + code + ' (หรือถูกปิดใช้งาน)');
  const privileged = canEditItem_(u, item);
  if (!privileged && !settingOn_('code_entry')) throw new Error('ยังไม่เปิดให้บันทึกด้วยรหัส หรือไม่มีสิทธิ์รายการนี้ ให้ผู้ดูแลระบบเปิดการบันทึกหน้างาน หรือเพิ่มสิทธิ์ให้บัญชีนี้');
  const reporter = String(e.reporter || '').trim().slice(0, 60);
  if (!privileged && !reporter) throw new Error('กรุณาใส่ชื่อผู้บันทึก');
  if (u.role !== 'admin') {
    if (lockSet_()[fy + '|' + month]) throw new Error('เดือนนี้ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อนแก้ไข');
    const now = fiscalNow_();
    if (fy * 12 + month > now.fy * 12 + now.month) throw new Error('บันทึกเดือนล่วงหน้าไม่ได้');
  }
  const who = privileged ? (u.email || u.name) : ('หน้างาน: ' + reporter);
  // รหัสที่รับผลจากระบบตรวจเช็คอัตโนมัติ: บันทึกผ่านแบบตรวจเช็คแทน (ผู้ดูแลระบบแก้ได้ และค่าที่แก้จะล็อกไว้)
  const autoT = inspLinkedCodes_()[code], autoCell = autoT && inspKpiOwns_(fy, month);
  if (autoCell && u.role !== 'admin') throw new Error('รายการ ' + code + ' รับผลจริงจากระบบตรวจเช็คอัตโนมัติ กรุณาบันทึกที่แท็บ "ตรวจเช็คบำรุงรักษา" ประเภท ' + autoT.name);
  if (autoCell) inspAddOverrides_(u, [{ target: 'PM_Plan', key: fy + '|' + month + '|' + code, note: 'บันทึกหน้างานโดยผู้ดูแลระบบ' }]);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_().getSheetByName('PM_Plan');
    const v = sh.getDataRange().getValues();
    const head = v[0];
    const iFy = head.indexOf('fy'), iM = head.indexOf('month'), iC = head.indexOf('code');
    const iP = head.indexOf('plan'), iA = head.indexOf('actual'), iN = head.indexOf('note');
    let row = 0;
    for (let i = 1; i < v.length; i++) {
      if (Number(v[i][iFy]) === fy && Number(v[i][iM]) === month && norm_(v[i][iC]) === code) { row = i + 1; break; }
    }
    if (!row || !(Number(v[row - 1][iP]) > 0)) throw new Error('เดือนนี้ไม่มีแผนของรหัส ' + code + ' ติดต่อผู้ดูแลระบบเพื่อกำหนดแผนก่อน');
    const old = v[row - 1][iA];
    const note = (e.note === undefined || e.note === null) ? v[row - 1][iN] : String(e.note).slice(0, 200);
    sh.getRange(row, iA + 1).setValue(actual);
    if (iN >= 0) sh.getRange(row, iN + 1).setValue(note);
    const iBy = head.indexOf('updated_by'), iAt = head.indexOf('updated_at');
    if (iBy >= 0) sh.getRange(row, iBy + 1).setValue(who);
    if (iAt >= 0) sh.getRange(row, iAt + 1).setValue(new Date().toISOString());
    if (norm_(old) !== norm_(actual)) {
      audit_({ email: who, name: who, role: privileged ? u.role : 'field' }, [{ action: 'บันทึกหน้างาน', sheet: 'PM_Plan', key: fy + '|' + month + '|' + code, field: 'actual', old: norm_(old), new: norm_(actual) }]);
    }
    inspKpiAfterPlan_([{ fy: fy, month: month, code: code }]);
    return { code: code, fy: fy, month: month, plan: Number(v[row - 1][iP]), actual: actual, old: old === '' ? null : Number(old) };
  } finally { lock.releaseLock(); }
}

/**
 * เขียนข้อมูลแบบ upsert ตามคีย์ ตรวจสิทธิ์เฉพาะแถวที่มีการเปลี่ยนแปลง
 * ถ้ามีแถวใดไม่ผ่านสิทธิ์ จะไม่บันทึกอะไรเลย
 */
function upsert_(name, keys, rows, fields, opt) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_().getSheetByName(name);
    const values = sh.getDataRange().getValues();
    const head = values[0].map(String);
    const index = {};
    for (let i = 1; i < values.length; i++) {
      index[keys.map(function (k) { return norm_(values[i][head.indexOf(k)]); }).join('|')] = i;
    }
    const now = new Date().toISOString();
    const plans = [];
    let unchanged = 0;
    rows.forEach(function (r) {
      if (opt.validate) opt.validate(r);
      const key = keys.map(function (k) { return norm_(r[k]); }).join('|');
      const idx = index[key];
      const optional = ['note'];
      const empty = fields.every(function (f) { return optional.indexOf(f) >= 0 || norm_(r[f]) === ''; });
      const oldRow = idx === undefined ? null : values[idx];
      if (empty) {
        if (oldRow && opt.allowDelete) plans.push({ type: 'delete', idx: idx, key: key, r: r, old: oldRow, changed: fields.filter(function (f) { return norm_(oldRow[head.indexOf(f)]) !== ''; }) });
        else unchanged++;
        return;
      }
      if (!oldRow) { plans.push({ type: 'insert', key: key, r: r, changed: fields.filter(function (f) { return norm_(r[f]) !== ''; }) }); return; }
      const diff = fields.filter(function (f) { return norm_(oldRow[head.indexOf(f)]) !== norm_(r[f]); });
      if (!diff.length) { unchanged++; return; }
      plans.push({ type: 'update', idx: idx, key: key, r: r, old: oldRow, diff: diff, changed: diff });
    });

    // ตรวจสิทธิ์ก่อนเขียน
    plans.forEach(function (p) {
      const msg = opt.check && opt.check(p.r, p.changed || []);
      if (msg) throw new Error(msg);
    });

    const logs = [];
    const dels = [];
    let saved = 0, removed = 0;
    plans.forEach(function (p) {
      if (p.type === 'delete') {
        dels.push(p.idx + 1);
        removed++;
        logs.push({ action: 'ลบ', sheet: name, key: p.key, field: fields.join(','), old: fields.map(function (f) { return norm_(p.old[head.indexOf(f)]); }).join(' / '), new: '' });
        return;
      }
      const line = head.map(function (h, ci) {
        if (h === 'updated_by') return opt.user.email || opt.user.name;
        if (h === 'updated_at') return now;
        if (keys.indexOf(h) >= 0 || fields.indexOf(h) >= 0) {
          const x = p.r[h];
          return x === null || x === undefined ? '' : x;
        }
        return p.old ? p.old[ci] : '';
      });
      if (p.type === 'update') {
        sh.getRange(p.idx + 1, 1, 1, head.length).setValues([line]);
        p.diff.forEach(function (f) {
          logs.push({ action: 'แก้ไข', sheet: name, key: p.key, field: f, old: norm_(p.old[head.indexOf(f)]), new: norm_(p.r[f]) });
        });
      } else {
        sh.appendRow(line);
        fields.forEach(function (f) {
          if (norm_(p.r[f]) !== '') logs.push({ action: 'เพิ่ม', sheet: name, key: p.key, field: f, old: '', new: norm_(p.r[f]) });
        });
      }
      saved++;
    });
    dels.sort(function (a, b) { return b - a; }).forEach(function (n) { sh.deleteRow(n); });
    audit_(opt.user, logs);
    return { saved: saved, removed: removed, unchanged: unchanged };
  } finally {
    lock.releaseLock();
  }
}

/** แก้ไขสถานะปัญหาคุณภาพข้อมูล (ผู้บันทึกและผู้ดูแลระบบ) */
function setIssueStatus(no, status) {
  const u = requireEditor_();
  if (['เปิด', 'กำลังแก้', 'เสร็จ'].indexOf(status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  const sh = db_().getSheetByName('DQ_Issues');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(no)) {
      const old = v[i][6];
      if (old !== status) {
        sh.getRange(i + 1, 7).setValue(status);
        audit_(u, [{ action: 'แก้ไข', sheet: 'DQ_Issues', key: String(no), field: 'status', old: old, new: status }]);
      }
      return true;
    }
  }
  return false;
}

/* ---------------- ผู้ดูแลระบบ ---------------- */

/** แทนที่รายชื่อผู้ใช้ทั้งหมด rows: [{email, name, role, units, active}] */
function saveUsers(rows) {
  const u = requireAdmin_();
  const seen = {}, keepCode = {};
  read_('Users').forEach(function (o) { keepCode[String(o.email).toLowerCase()] = String(o.staff_code || '').trim(); });
  const clean = rows.filter(function (r) { return String(r.email || '').trim() !== ''; }).map(function (r) {
    const email = String(r.email).trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง: ' + r.email);
    if (seen[email]) throw new Error('อีเมลซ้ำ: ' + email);
    seen[email] = true;
    if (ROLES.indexOf(r.role) < 0) throw new Error('บทบาทไม่ถูกต้อง: ' + r.role);
    const units = r.role === 'admin' ? '*' : parseUnits_(r.units).join(',');
    if (r.role === 'editor' && !units) throw new Error('ผู้บันทึกข้อมูลต้องระบุหน่วยอย่างน้อยหนึ่งหน่วย: ' + email);
    const lst = function (x) { return String(x || '').split(',').map(function (y) { return y.trim(); }).filter(String).join(','); };
    const it = r.role === 'admin' || r.role === 'viewer' ? '' : lst(r.insp_types), ig = r.role === 'admin' || r.role === 'viewer' ? '' : lst(r.insp_groups);
    return [email, r.name || '', r.role, r.role === 'inspector' ? '' : units, r.active === false || r.active === 0 ? 0 : 1, it, ig, r.staff_code !== undefined ? String(r.staff_code || '').trim() : keepCode[email] || ''];
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const old = read_('Users');
    const oldMap = {};
    old.forEach(function (o) { oldMap[String(o.email).toLowerCase()] = o; });
    const sh = db_().getSheetByName('Users');
    const uh = headOf_(sh).map(String), cols = SHEETS.Users.map(function (h) { return uh.indexOf(h); });
    if (cols.some(function (i) { return i < 0; })) throw new Error('ชีต Users ขาดคอลัมน์ ให้รัน upgradeDb');
    const width = Math.max.apply(null, cols) + 1;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
    if (clean.length) sh.getRange(2, 1, clean.length, width).setValues(clean.map(function (c) { const line = []; for (let i = 0; i < width; i++) line.push(''); cols.forEach(function (ci, j) { line[ci] = c[j]; }); return line; }));
    const logs = [];
    clean.forEach(function (c) {
      const o = oldMap[c[0]];
      if (!o) logs.push({ action: 'เพิ่ม', sheet: 'Users', key: c[0], field: 'role/units', old: '', new: c[2] + ' / ' + c[3] });
      else ['name', 'role', 'units', 'active', 'insp_types', 'insp_groups'].forEach(function (f, i) {
        const nv = c[i + 1];
        if (norm_(o[f]) !== norm_(nv)) logs.push({ action: 'แก้ไข', sheet: 'Users', key: c[0], field: f, old: norm_(o[f]), new: norm_(nv) });
      });
    });
    Object.keys(oldMap).forEach(function (k) {
      if (!seen[k]) logs.push({ action: 'ลบ', sheet: 'Users', key: k, field: 'role/units', old: oldMap[k].role + ' / ' + oldMap[k].units, new: '' });
    });
    audit_(u, logs);
    return { users: clean.length, changes: logs.length };
  } finally {
    lock.releaseLock();
  }
}

/** ล็อก/ปลดล็อกเดือนของปีงบประมาณ */
function setLock(fy, month, locked) {
  const u = requireAdmin_();
  fy = Number(fy); month = Number(month);
  if (!(month >= 1 && month <= 12) || !fy) throw new Error('ปีงบหรือเดือนไม่ถูกต้อง');
  const sh = db_().getSheetByName('Locks');
  const v = sh.getDataRange().getValues();
  let row = 0;
  for (let i = 1; i < v.length; i++) { if (Number(v[i][0]) === fy && Number(v[i][1]) === month) { row = i + 1; break; } }
  if (locked && !row) {
    sh.appendRow([fy, month, u.email || u.name, new Date().toISOString()]);
    audit_(u, [{ action: 'ล็อกเดือน', sheet: 'Locks', key: fy + '|' + month, field: 'locked', old: '0', new: '1' }]);
  } else if (!locked && row) {
    sh.deleteRow(row);
    audit_(u, [{ action: 'ปลดล็อกเดือน', sheet: 'Locks', key: fy + '|' + month, field: 'locked', old: '1', new: '0' }]);
  }
  return read_('Locks');
}

function targetText_(dir, thr, kind) {
  const t = Number(thr), pct = kind === 'ratio' ? '%' : '';
  const map = { ge: '≥ ', gt: '> ', lt: '< ', le_year: '≤ ' };
  if (dir === 'zero') return '0 ครั้ง';
  if (dir === 'info') return 'ติดตาม';
  return (map[dir] || '') + t + pct + (dir === 'le_year' ? ' ครั้ง/ปี' : '');
}

function checkTarget_(r) {
  if (DIRS.indexOf(r.dir) < 0) throw new Error('ทิศทางเป้าหมายไม่ถูกต้อง: ' + r.kpi_id);
  if (r.dir !== 'zero' && r.dir !== 'info') {
    if (r.thr === null || r.thr === undefined || r.thr === '' || isNaN(Number(r.thr))) throw new Error('ต้องระบุตัวเลขเป้าหมายของ ' + r.kpi_id);
    if (Number(r.thr) < 0) throw new Error('เป้าหมายต้องไม่ติดลบ: ' + r.kpi_id);
  }
}

/** เพิ่ม/แก้ไขตัวชี้วัด rows: [{kpi_id, name, group, unit, target_text, kind, dir, thr, max100, num_label, den_label, order, active, source}] */
function saveMaster(rows) {
  const u = requireAdmin_();
  rows.forEach(function (r) {
    if (!/^[A-Za-z0-9_]{1,20}$/.test(String(r.kpi_id || ''))) throw new Error('รหัสตัวชี้วัดใช้ได้เฉพาะ A-Z, 0-9, _ : ' + r.kpi_id);
    if (!String(r.name || '').trim()) throw new Error('ต้องระบุชื่อตัวชี้วัด ' + r.kpi_id);
    if (['ratio', 'count'].indexOf(r.kind) < 0) throw new Error('ชนิดตัวชี้วัดต้องเป็น ratio หรือ count: ' + r.kpi_id);
    checkTarget_(r);
    r.source = r.source || 'manual';
    if (['manual', 'pm', 'insp'].indexOf(r.source) < 0) throw new Error('แหล่งข้อมูลต้องเป็น manual, pm หรือ insp: ' + r.kpi_id);
    if ((r.source === 'pm' || r.source === 'insp') && r.kind !== 'ratio') throw new Error('KPI ที่คำนวณจากแผน PM / ผลตรวจเช็ค ต้องเป็นชนิดร้อยละ: ' + r.kpi_id);
    if (!String(r.target_text || '').trim()) r.target_text = targetText_(r.dir, r.thr, r.kind);
  });
  return upsert_('KPI_Master', ['kpi_id'], rows, MASTER_FIELDS, { user: u, allowDelete: false });
}

/** เป้าหมายรายปีงบ rows: [{fy, kpi_id, dir, thr, target_text, note}] แถวที่ dir ว่างทั้งหมด = ลบ (กลับไปใช้ค่าเริ่มต้นใน KPI_Master) */
function saveTargets(rows) {
  const u = requireAdmin_();
  const byId = {};
  read_('KPI_Master').forEach(function (k) { byId[k.kpi_id] = k; });
  return upsert_('KPI_Targets', ['fy', 'kpi_id'], rows, ['dir', 'thr', 'target_text', 'note'], {
    user: u,
    allowDelete: true,
    validate: function (r) {
      if (!byId[r.kpi_id]) throw new Error('ไม่พบตัวชี้วัด ' + r.kpi_id);
      if (!(Number(r.fy) > 2500 && Number(r.fy) < 2700)) throw new Error('ปีงบประมาณไม่ถูกต้อง: ' + r.fy);
      if (norm_(r.dir) !== '') {
        checkTarget_(r);
        if (!String(r.target_text || '').trim()) r.target_text = targetText_(r.dir, r.thr, byId[r.kpi_id].kind);
      }
    }
  });
}

/* ---------------- ระบบและรายการบำรุงรักษา (รหัส 7 หลัก) ---------------- */

/** ระบบงาน rows: [{sys_code (4 หลัก), name, discipline, active}] */
function saveSystems(rows) {
  const u = requireAdmin_();
  return upsert_('PM_Systems', ['sys_code'], rows, ['name', 'discipline', 'active', 'responsible_name', 'responsible_email'], {
    user: u,
    allowDelete: false,
    validate: function (r) {
      r.sys_code = norm_(r.sys_code);
      if (!/^\d{4}$/.test(r.sys_code)) throw new Error('รหัสระบบต้องเป็นตัวเลข 4 หลัก: ' + r.sys_code);
      if (!String(r.name || '').trim()) throw new Error('ต้องระบุชื่อระบบ ' + r.sys_code);
      if (!String(r.discipline || '').trim()) throw new Error('ต้องระบุหมวดงานของระบบ ' + r.sys_code);
      if (r.responsible_email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(r.responsible_email).trim())) throw new Error('อีเมลผู้รับผิดชอบระบบไม่ถูกต้อง ' + r.sys_code);
    }
  });
}

/** รายการบำรุงรักษา rows: [{code (7 หลัก), name, discipline, location, uom, responsible_name, responsible_email, kpi_id, contractor, active, note}] */
function saveItems(rows) {
  const u = requireAdmin_();
  const sys = {}, kpis = {}, staff = {};
  read_('PM_Systems').forEach(function (s) { sys[norm_(s.sys_code)] = s; });
  read_('KPI_Master').forEach(function (k) { kpis[k.kpi_id] = k; });
  read_('Staff').forEach(function (s) { staff[String(s.name).trim()] = s; });
  const seen = {};
  return upsert_('PM_Items', ['code'], rows, ITEM_FIELDS, {
    user: u,
    allowDelete: false,
    validate: function (r) {
      r.code = norm_(r.code);
      if (!/^\d{7}$/.test(r.code)) throw new Error('รหัสรายการต้องเป็นตัวเลข 7 หลัก: ' + r.code);
      if (seen[r.code]) throw new Error('รหัสรายการซ้ำ: ' + r.code);
      seen[r.code] = true;
      const s = sys[r.code.slice(0, 4)];
      if (!s) throw new Error('รหัส ' + r.code + ': ไม่พบระบบรหัส ' + r.code.slice(0, 4) + ' (เพิ่มในรายการระบบก่อน)');
      if (!String(r.name || '').trim()) throw new Error('ต้องระบุชื่อรายการ ' + r.code);
      r.discipline = s.discipline;
      if (r.kpi_id && !kpis[r.kpi_id]) throw new Error('รหัส ' + r.code + ': ไม่พบตัวชี้วัด ' + r.kpi_id);
      r.responsible_name = String(r.responsible_name || '').trim();
      r.responsible_email = String(r.responsible_email || '').trim().toLowerCase();
      if (!r.responsible_email && staff[r.responsible_name] && staff[r.responsible_name].email) r.responsible_email = String(staff[r.responsible_name].email).toLowerCase();
      if (r.responsible_email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r.responsible_email)) throw new Error('รหัส ' + r.code + ': อีเมลผู้รับผิดชอบไม่ถูกต้อง');
    }
  });
}

/** รายชื่อผู้รับผิดชอบ rows: [{name, email, duty, active}] */
function saveStaff(rows) {
  const u = requireAdmin_();
  return upsert_('Staff', ['name'], rows, ['email', 'duty', 'active'], {
    user: u,
    allowDelete: false,
    validate: function (r) {
      r.name = String(r.name || '').trim();
      if (!r.name) throw new Error('ต้องระบุชื่อผู้รับผิดชอบ');
      r.email = String(r.email || '').trim().toLowerCase();
      if (r.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r.email)) throw new Error('อีเมลไม่ถูกต้อง: ' + r.email);
    }
  });
}

/** คัดลอกค่าแผน (ไม่รวมผลจริง) จากปีงบหนึ่งไปอีกปีงบ เฉพาะรหัส/เดือนที่ปลายทางยังไม่มีแถว */
function copyPlanFromFY(fromFy, toFy) {
  const u = requireAdmin_();
  fromFy = Number(fromFy); toFy = Number(toFy);
  if (!fromFy || !toFy || fromFy === toFy) throw new Error('ปีงบต้นทางและปลายทางไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = read_('PM_Plan');
    const exist = {};
    all.forEach(function (r) { if (Number(r.fy) === toFy) exist[Number(r.month) + '|' + norm_(r.code)] = true; });
    const active = {};
    read_('PM_Items').forEach(function (i) { if (isActive_(i.active)) active[norm_(i.code)] = true; });
    const now = new Date().toISOString();
    const out = [];
    all.forEach(function (r) {
      if (Number(r.fy) !== fromFy || r.plan === null || r.plan === '') return;
      const code = norm_(r.code);
      if (!active[code] || exist[Number(r.month) + '|' + code]) return;
      out.push([toFy, Number(r.month), code, r.plan, '', '', u.email || u.name, now]);
    });
    if (out.length) {
      const sh = db_().getSheetByName('PM_Plan');
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = SHEETS.PM_Plan;
      const rows = out.map(function (r) { return head.map(function (h) { const i = idx.indexOf(h); return i >= 0 ? r[i] : ''; }); });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
    }
    audit_(u, [{ action: 'คัดลอกแผน', sheet: 'PM_Plan', key: fromFy + '→' + toFy, field: 'plan', old: '', new: out.length + ' แถว' }]);
    return { copied: out.length };
  } finally {
    lock.releaseLock();
  }
}

// ==INSP-CORE-START==
/* ====== แกนคำนวณระบบตรวจเช็ค (ใช้ร่วมกันทั้งฝั่งเซิร์ฟเวอร์และหน้าเว็บ) — ไม่ผูกกับ DOM/Sheets ====== */
var INSP_NG_WORDS = ['ไม่ผ่าน', 'ไม่ปกติ', 'ผิดปกติ', 'ไม่ได้', 'ใช้งานไม่', 'ชำรุด', 'เสีย', 'รั่ว', 'ไม่สะอาด', 'ไม่เรียบร้อย', 'ขัดข้อง', 'ไม่ดี', 'ไม่ครบ', 'ขัดข้อง'];
var INSP_OK_WORDS = ['ปกติ', 'เรียบร้อย', 'สะอาด', 'ใช้งานได้', 'ไม่มี', 'ผ่าน', 'ครบ', 'ดี'];

/** แปลงข้อความผลตรวจเป็น 'ok' | 'ng' | '' (ไม่ระบุ) */
function inspNormStatus(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).replace(/[​\s]+/g, ' ').trim().toLowerCase();
  if (!s || s === '-' || s === 'n/a' || s === 'na') return '';
  if (/^(ng|fail|bad|no)$/.test(s)) return 'ng';
  if (/^(ok|pass|good|yes)$/.test(s)) return 'ok';
  var i;
  for (i = 0; i < INSP_NG_WORDS.length; i++) if (s.indexOf(INSP_NG_WORDS[i]) >= 0) return 'ng';
  for (i = 0; i < INSP_OK_WORDS.length; i++) if (s.indexOf(INSP_OK_WORDS[i]) >= 0) return 'ok';
  return '';
}
/** ดึงตัวเลขจากข้อความ เช่น "36.2 hrs" -> 36.2  (คืน null ถ้าไม่มีตัวเลข) */
function inspNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var str = String(v);
  if (str.indexOf('=') >= 0) str = str.slice(str.lastIndexOf('=') + 1); // "17.6+123.2=140.8" -> ผลลัพธ์หลังเครื่องหมาย =
  if (/^\s*-?\d+,\d{1,2}\s*$/.test(str)) str = str.replace(',', '.'); // "20,8" = ทศนิยม
  var m = /-?\d[\d,]*\.?\d*/.exec(str);
  if (!m) return null;
  var n = Number(m[0].replace(/,/g, ''));
  return isFinite(n) ? n : null;
}
function inspRange(c) {
  var tol = Number(c.tol) || 0;
  var lo = (c.min === null || c.min === undefined || c.min === '') ? null : Number(c.min) * (1 - tol / 100);
  var hi = (c.max === null || c.max === undefined || c.max === '') ? null : Number(c.max) * (1 + tol / 100);
  return { lo: lo, hi: hi, has: lo !== null || hi !== null };
}
/**
 * ประเมินค่าที่ตรวจ 1 รายการ  c = นิยามรายการ {kind: s|n|m|t, min, max, tol}
 * คืน {v: ค่าที่เก็บ, st: 'ok'|'ng'|'wn'|''}  (wn = เกินเกณฑ์แต่เป็นรายการ "เฝ้าระวัง" soft ไม่ถือว่าไม่ปกติ)   ค่าที่เก็บ: s -> 'ok'/'ng'  n -> ตัวเลข  m -> ตัวเลขหรือ 'ok'/'ng'  t -> ข้อความ
 */
function inspEvalCheck(c, raw) {
  var kind = c.kind || 's';
  if (kind === 't') { var t = (raw === null || raw === undefined) ? '' : String(raw).trim().slice(0, 200); return { v: t, st: '' }; }
  if (kind === 's') { var s = inspNormStatus(raw); return { v: s, st: (s === 'ng' && Number(c.soft)) ? 'wn' : s }; }
  var n = inspNum(raw), r = inspRange(c);
  if (kind === 'm' && n === null) { var s2 = inspNormStatus(raw); return { v: s2, st: (s2 === 'ng' && Number(c.soft)) ? 'wn' : s2 }; }
  if (n === null) return { v: '', st: '' };
  if (kind === 'm' && !r.has) return { v: n, st: 'ok' };
  if (!r.has) return { v: n, st: '' };
  var ok = (r.lo === null || n >= r.lo - 1e-9) && (r.hi === null || n <= r.hi + 1e-9);
  return { v: n, st: ok ? 'ok' : (Number(c.soft) ? 'wn' : 'ng') };
}
/** ประเมินผลทั้งใบ checks = นิยามรายการของประเภทนั้น, raw = {check_id: ค่าดิบ} */
function inspEvalEntry(checks, raw) {
  var res = {}, ng = [], wn = [], filled = 0;
  (checks || []).forEach(function (c) {
    var e = inspEvalCheck(c, raw ? raw[c.check_id] : null);
    if (e.v !== '' && e.v !== null) { res[c.check_id] = e.v; filled++; }
    if (e.st === 'ng') ng.push(c.check_id);
    if (e.st === 'wn') wn.push(c.check_id);
  });
  return { res: res, ng: ng, wn: wn, ngN: ng.length, wnN: wn.length, filled: filled, st: ng.length ? 'ng' : (wn.length ? 'wn' : 'ok') };
}

/* ---- วันที่ (ข้อความ YYYY-MM-DD ล้วน ไม่พึ่งเขตเวลา) ---- */
function inspPad(n) { return n < 10 ? '0' + n : String(n); }
function inspParse(s) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || '')); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; }
function inspDays(s) { var p = inspParse(s); return p ? Math.floor(Date.UTC(p.y, p.m - 1, p.d) / 86400000) : null; }
function inspDiff(a, b) { return inspDays(a) - inspDays(b); }
function inspAdd(s, n) { var t = new Date((inspDays(s) + n) * 86400000); return t.getUTCFullYear() + '-' + inspPad(t.getUTCMonth() + 1) + '-' + inspPad(t.getUTCDate()); }
function inspEom(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
function inspFyOf(s) { var p = inspParse(s); return p ? p.y + 543 + (p.m >= 10 ? 1 : 0) : null; }
function inspMiOf(m) { return ((m + 2) % 12) + 1; }
function inspYmOf(fy, mi) { var m = ((mi + 8) % 12) + 1; var y = fy - 543 - (m >= 10 ? 1 : 0); return { y: y, m: m, key: y + '-' + inspPad(m) }; }
/** รอบของวันที่: รายวัน 'YYYY-MM-DD'  รายเดือน 'YYYY-MM'  รายสัปดาห์ 'YYYY-MM-Wn' (W1=วันที่ 1-7, W2=8-14, W3=15-21, W4=22-สิ้นเดือน) */
function inspPeriodOf(freq, ymd, weekHint) {
  var p = inspParse(ymd); if (!p) return '';
  var ym = p.y + '-' + inspPad(p.m);
  if (freq === 'D') return ym + '-' + inspPad(p.d);
  if (freq !== 'W') return ym;
  var w = Number(weekHint); if (!(w >= 1 && w <= 4)) w = Math.min(4, Math.ceil(p.d / 7));
  return ym + '-W' + w;
}
function inspBounds(period) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(period)) return [period, period];
  var m = /^(\d{4})-(\d{2})(?:-W(\d))?$/.exec(period); if (!m) return null;
  var y = +m[1], mo = +m[2], e = inspEom(y, mo), ym = m[1] + '-' + m[2];
  if (!m[3]) return [ym + '-01', ym + '-' + inspPad(e)];
  var w = +m[3];
  return [ym + '-' + inspPad((w - 1) * 7 + 1), ym + '-' + inspPad(w === 4 ? e : w * 7)];
}
function inspPeriodLabel(period) {
  var MN = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  var d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(period);
  if (d) return (+d[3]) + ' ' + MN[+d[2] - 1] + ' ' + String((+d[1] + 543) % 100);
  var m = /^(\d{4})-(\d{2})(?:-W(\d))?$/.exec(period); if (!m) return period;
  return (m[3] ? 'สัปดาห์ ' + m[3] + ' ' : '') + MN[+m[2] - 1] + ' ' + String((+m[1] + 543) % 100);
}

/**
 * ตารางรอบการตรวจของ 1 ประเภท ตลอดปีงบ fy  (asOf = วันนี้ YYYY-MM-DD)
 * type {freq}, assets = ทรัพย์สินของประเภทนี้ {asset_id, active, since, until}, logs = บันทึกของประเภทนี้ {asset_id, period, st}
 * คืนรายการช่อง {fy, mi, period, state, exp, done, cov, ng, n, missing:[asset_id]}
 * state: done | partial | missed | open | future | na
 */
function inspMatrix(type, assets, logs, fy, asOf, opt) {
  var minCov = (type.min_cov !== undefined && type.min_cov !== null && type.min_cov !== '' && Number(type.min_cov) > 0) ? Number(type.min_cov) : ((opt && opt.minCov !== undefined && opt.minCov !== null) ? Number(opt.minCov) : 1);
  var weekly = type.freq === 'W', daily = type.freq === 'D';
  var byP = {};
  (logs || []).forEach(function (l) { (byP[l.period] = byP[l.period] || []).push(l); });
  var cells = [];
  for (var mi = 1; mi <= 12; mi++) {
    var ym = inspYmOf(fy, mi);
    var periods = weekly ? [ym.key + '-W1', ym.key + '-W2', ym.key + '-W3', ym.key + '-W4'] : [ym.key];
    if (daily) { periods = []; for (var dd = 1, de = inspEom(ym.y, ym.m); dd <= de; dd++) periods.push(ym.key + '-' + inspPad(dd)); }
    var off = !!(type._planOff && type._planOff[fy + '|' + mi]); // แผน PM เดือนนี้ = 0 / ไม่มีแผน -> ไม่ต้องตรวจ
    periods.forEach(function (p) {
      var b = inspBounds(p);
      var exp = (assets || []).filter(function (a) {
        if (!(a.active === undefined || a.active === null || isInspActive(a.active))) return false;
        if (a.since && String(a.since).slice(0, 7) > ym.key) return false;
        if (a.until && String(a.until).slice(0, 7) < ym.key) return false;
        return true;
      });
      var ls = byP[p] || [], seen = {}, ng = 0, wn = 0;
      ls.forEach(function (l) { seen[l.asset_id] = true; if (l.st === 'ng') ng++; else if (l.st === 'wn') wn++; });
      var done = exp.filter(function (a) { return seen[a.asset_id]; }).length;
      var cov = exp.length ? done / exp.length : null;
      var started = b[0] <= asOf, ended = b[1] < asOf;
      var state;
      if (!started) state = 'future';
      else if (cov === null || off) state = 'na';
      else if (cov >= minCov - 1e-9) state = 'done';
      else if (!ended) state = 'open';
      else state = done > 0 ? 'partial' : 'missed';
      cells.push({ fy: fy, mi: mi, period: p, start: b[0], end: b[1], state: state, planOff: off, exp: exp.length, done: done, cov: cov, ng: ng, wn: wn, n: ls.length, minCov: minCov,
        missing: exp.filter(function (a) { return !seen[a.asset_id]; }).map(function (a) { return a.asset_id; }) });
    });
  }
  return cells;
}
function isInspActive(v) { return !(v === 0 || v === false || v === '0' || v === 'false' || v === 'ไม่' || v === 'FALSE'); }

/** สรุปความสม่ำเสมอของการตรวจ: รอบที่ครบ / รอบที่ถึงกำหนด (ครบแล้ว หรือสิ้นสุดแล้ว) */
function inspCompliance(cells) {
  var due = 0, done = 0, missed = 0, partial = 0, open = 0;
  cells.forEach(function (c) {
    if (c.state === 'done') { due++; done++; }
    else if (c.state === 'missed') { due++; missed++; }
    else if (c.state === 'partial') { due++; partial++; }
    else if (c.state === 'open') open++;
  });
  return { due: due, done: done, missed: missed, partial: partial, open: open, pct: due ? done / due * 100 : null };
}

/** ตรวจสอบกำหนดวันแก้ไขตามความรุนแรง (วัน) */
function inspSla(sev, opt) { return sev === 'high' ? ((opt && opt.slaHigh) || 7) : ((opt && opt.slaNormal) || 30); }

/** รหัสงานแก้ไข กำหนดแน่นอนจาก (ทรัพย์สิน, รายการ, วันที่พบ) เพื่อให้นำเข้าซ้ำได้โดยไม่เกิดซ้ำ */
function inspActionId(assetId, checkId, found) {
  var s = assetId + '|' + checkId, h = 5381;
  for (var i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
  return 'A' + String(found || '').replace(/-/g, '').slice(2) + '-' + h.toString(36).slice(0, 5).toUpperCase();
}

/**
 * สร้าง/อัปเดตงานแก้ไขจากบันทึกที่เรียงตามวันที่ (ใช้ทั้งตอนนำเข้าและตอนบันทึกใหม่)
 * logs เรียงเก่า->ใหม่ ต่อ (asset, check): ต่อเนื่อง ng = 1 งาน; ปกติหลังจากนั้น = ตรวจซ้ำปกติ (verified_ok_date)
 * opt.closeOnOk=true (ตอนนำเข้าประวัติ): ปิดงานให้เองเมื่อพบปกติถัดมา
 * existing = งานเดิม {action_id, asset_id, check_id, status,...} ใช้ต่อยอดงานที่ยังเปิดอยู่
 */
function inspDeriveActions(checksById, assetsById, logs, existing, opt) {
  opt = opt || {};
  var open = {}; // asset|check -> action
  var out = {};
  (existing || []).forEach(function (a) {
    out[a.action_id] = a;
    if (a.status === 'open' || a.status === 'in_progress') open[a.asset_id + '|' + a.check_id] = a;
  });
  var sorted = (logs || []).slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.log_id < b.log_id ? -1 : 1); });
  sorted.forEach(function (l) {
    var res = l.results; if (typeof res === 'string') { try { res = JSON.parse(res); } catch (e) { res = {}; } }
    var a0 = assetsById[l.asset_id] || {};
    Object.keys(checksById[l.type_id] || {}).forEach(function (cid) {
      var c = checksById[l.type_id][cid];
      if (c.kind === 't') return;
      var key = l.asset_id + '|' + cid, cur = open[key];
      var ev = inspEvalCheck(c, res ? res[cid] : null), bad = ev.st === 'ng';
      if (bad) {
        var v = res ? res[cid] : '';
        if (cur) { cur.last_seen = l.date; cur.seen_count = (Number(cur.seen_count) || 1) + 1; if (l.note && /ทำเรื่อง|แทงชำรุด|แทนชำรุด|ส่งซ่อม|แจ้งซ่อม|เบิก/.test(l.note) && cur.status === 'open') { cur.status = 'in_progress'; if (!cur.action_note) cur.action_note = l.note; } cur.verified_ok_date = ''; return; }
        var sev = Number(c.crit) ? 'high' : 'normal';
        var proc = l.note && /ทำเรื่อง|แทงชำรุด|แทนชำรุด|ส่งซ่อม|แจ้งซ่อม|เบิก/.test(l.note);
        var act = { action_id: inspActionId(l.asset_id, cid, l.date), type_id: l.type_id, asset_id: l.asset_id, check_id: cid,
          issue: (c.label || cid) + ((c.kind !== 's' && v !== '' && v !== undefined && v !== 'ng' && v !== 'ok') ? ' = ' + v : ' ไม่ปกติ') + (l.note ? ' — ' + l.note : ''),
          severity: sev, found_date: l.date, last_seen: l.date, seen_count: 1, assignee: '', due_date: inspAdd(l.date, inspSla(sev, opt)),
          status: proc ? 'in_progress' : 'open', action_note: proc ? l.note : '', done_date: '', verified_ok_date: '' };
        out[act.action_id] = act; open[key] = act;
      } else if (cur && ev.st === 'ok') {
        cur.verified_ok_date = l.date;
        if (opt.closeOnOk) { cur.status = 'done'; cur.done_date = l.date; cur.action_note = cur.action_note || 'ปิดอัตโนมัติ: ตรวจซ้ำพบปกติ'; delete open[key]; }
      }
    });
  });
  return Object.keys(out).map(function (k) { return out[k]; });
}

function inspActive_(x) { return x && isInspActive(x.active); }
function inspIsOpen(a) { return a.status === 'open' || a.status === 'in_progress'; }

/**
 * รายการแจ้งเตือน ณ วันที่ asOf
 * kind: missed (เลยรอบไม่ได้ตรวจ/ตรวจไม่ครบ) | due (ใกล้ครบกำหนดรอบ) | action_overdue | action_due | repeat (พบซ้ำหลายรอบ) | verify (ตรวจซ้ำปกติ รอยืนยันปิดงาน)
 * level: bad | warn | info
 */
function inspAlerts(types, assets, logs, actions, asOf, opt) {
  opt = opt || {};
  var out = [], fyNow = inspFyOf(asOf);
  (types || []).filter(inspActive_).forEach(function (t) {
    var as = (assets || []).filter(function (a) { return a.type_id === t.type_id; });
    var ls = (logs || []).filter(function (l) { return l.type_id === t.type_id; });
    var weekly = t.freq === 'W', look = weekly ? 42 : 75, win = weekly ? 2 : (opt.warnDays || 7);
    var cells = inspMatrix(t, as, ls, fyNow, asOf).concat(inspMatrix(t, as, ls, fyNow - 1, asOf));
    if (t.freq === 'D') { // รายวัน: รวมเป็นแจ้งเตือนเดียวต่อประเภท (วันที่ไม่ได้ตรวจ/ตรวจไม่ครบ ใน 14 วันล่าสุด)
      var miss = cells.filter(function (c) { return (c.state === 'missed' || c.state === 'partial') && inspDiff(asOf, c.end) <= 14; }).sort(function (a, b) { return a.period < b.period ? 1 : -1; });
      if (miss.length) {
        var mset = {}; miss.forEach(function (c) { c.missing.forEach(function (x) { mset[x] = 1; }); });
        out.push({ kind: 'missed', level: 'bad', type_id: t.type_id, period: miss[0].period, missing: Object.keys(mset), done: miss[0].done, exp: miss[0].exp, days: inspDiff(asOf, miss[0].end), state: miss[0].state, nDays: miss.length, win: 14 });
      }
      return;
    }
    cells.forEach(function (c) {
      if (c.state === 'future' || c.state === 'na' || c.state === 'done') return;
      if (c.state === 'open') {
        var left = inspDiff(c.end, asOf);
        if (left <= win) out.push({ kind: 'due', level: 'warn', type_id: t.type_id, period: c.period, missing: c.missing, done: c.done, exp: c.exp, days: left });
      } else {
        var late = inspDiff(asOf, c.end);
        if (late <= look) out.push({ kind: 'missed', level: 'bad', type_id: t.type_id, period: c.period, missing: c.missing, done: c.done, exp: c.exp, days: late, state: c.state });
      }
    });
  });
  (actions || []).forEach(function (a) {
    if (a.status === 'done' || a.status === 'cancelled') return;
    var base = { type_id: a.type_id, asset_id: a.asset_id, action_id: a.action_id, check_id: a.check_id };
    if (a.due_date && a.due_date < asOf) out.push(Object.assign({ kind: 'action_overdue', level: 'bad', days: inspDiff(asOf, a.due_date) }, base));
    else if (a.due_date && inspDiff(a.due_date, asOf) <= 3) out.push(Object.assign({ kind: 'action_due', level: 'warn', days: inspDiff(a.due_date, asOf) }, base));
    if ((Number(a.seen_count) || 0) >= 3) out.push(Object.assign({ kind: 'repeat', level: 'warn', days: inspDiff(asOf, a.found_date), count: Number(a.seen_count) }, base));
    if (a.verified_ok_date) out.push(Object.assign({ kind: 'verify', level: 'info', days: inspDiff(asOf, a.verified_ok_date) }, base));
  });
  var rank = { bad: 0, warn: 1, info: 2 };
  return out.sort(function (x, y) { return (rank[x.level] - rank[y.level]) || ((y.days || 0) - (x.days || 0)); });
}

/** สรุปภาพรวมสำหรับผู้บริหาร ของปีงบ fy */
function inspSummary(types, assets, logs, actions, fy, asOf, opt) {
  var per = [], monthly = [], i;
  for (i = 1; i <= 12; i++) monthly.push({ mi: i, done: 0, exp: 0, logs: 0, ng: 0 });
  var tot = { due: 0, done: 0, covDone: 0, covExp: 0, logs: 0, ng: 0, wn: 0, pcts: [] };
  (types || []).filter(inspActive_).sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); }).forEach(function (t) {
    var as = (assets || []).filter(function (a) { return a.type_id === t.type_id; });
    var ls = (logs || []).filter(function (l) { return l.type_id === t.type_id; });
    var cells = inspMatrix(t, as, ls, fy, asOf, opt), comp = inspCompliance(cells);
    var cd = 0, ce = 0;
    cells.forEach(function (c) {
      if (c.state === 'done' || c.state === 'partial' || c.state === 'missed') { cd += c.done; ce += c.exp; monthly[c.mi - 1].done += c.done; monthly[c.mi - 1].exp += c.exp; }
    });
    var inFy = ls.filter(function (l) { return inspFyOf(l.date) === fy; });
    var ng = inFy.filter(function (l) { return l.st === 'ng'; }).length, wn = inFy.filter(function (l) { return l.st === 'wn'; }).length;
    inFy.forEach(function (l) { var m = inspMiOf(+String(l.date).slice(5, 7)); monthly[m - 1].logs++; if (l.st === 'ng') monthly[m - 1].ng++; });
    var acts = (actions || []).filter(function (a) { return a.type_id === t.type_id; });
    var open = acts.filter(inspIsOpen), over = open.filter(function (a) { return a.due_date && a.due_date < asOf; });
    var last = ''; ls.forEach(function (l) { if (l.date > last) last = l.date; });
    per.push({ type: t, cells: cells, comp: comp, covDone: cd, covExp: ce, covPct: ce ? cd / ce * 100 : null, logs: inFy.length, ng: ng, wn: wn,
      normalPct: inFy.length ? (inFy.length - ng) / inFy.length * 100 : null, open: open.length, overdue: over.length, last: last, lastAge: last ? inspDiff(asOf, last) : null,
      assets: as.filter(inspActive_).length });
    tot.due += comp.due; tot.done += comp.done; if (comp.pct !== null) tot.pcts.push(comp.pct); tot.covDone += cd; tot.covExp += ce; tot.logs += inFy.length; tot.ng += ng; tot.wn += wn;
  });
  var oa = (actions || []).filter(inspIsOpen), ages = oa.map(function (a) { return inspDiff(asOf, a.found_date); });
  var closed = (actions || []).filter(function (a) { return a.status === 'done' && a.done_date && a.found_date; });
  return { per: per, monthly: monthly, total: {
    // ค่าเฉลี่ยของแต่ละประเภท (ถ่วงเท่ากัน) — ไม่ให้งานรายวัน (365 รอบ/ปี) กลบงานรายเดือน
    compPct: tot.pcts.length ? tot.pcts.reduce(function (a, b) { return a + b; }, 0) / tot.pcts.length : null, cyclesDue: tot.due, cyclesDone: tot.done, covPct: tot.covExp ? tot.covDone / tot.covExp * 100 : null, logs: tot.logs, ng: tot.ng, wn: tot.wn,
    normalPct: tot.logs ? (tot.logs - tot.ng) / tot.logs * 100 : null,
    open: oa.length, inProgress: oa.filter(function (a) { return a.status === 'in_progress'; }).length,
    overdue: oa.filter(function (a) { return a.due_date && a.due_date < asOf; }).length,
    high: oa.filter(function (a) { return a.severity === 'high'; }).length,
    avgAge: ages.length ? ages.reduce(function (s, x) { return s + x; }, 0) / ages.length : null,
    maxAge: ages.length ? Math.max.apply(null, ages) : null,
    closedN: closed.length, avgClose: closed.length ? closed.reduce(function (s, a) { return s + inspDiff(a.done_date, a.found_date); }, 0) / closed.length : null
  } };
}


/**
 * ภาพรวมรายอาคาร (group ของจุดตรวจ) x ประเภท ของปีงบ fy
 * คืน {groups:[{name, assets, byType:{type_id:{pct, due, done, assets, lastAge}}, open, overdue, pct}], types:[...]}
 */
function inspByGroup(types, assets, logs, actions, fy, asOf, opt) {
  var act = (types || []).filter(inspActive_).sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  var gmap = {}, assetG = {};
  (assets || []).forEach(function (a) { if (!inspActive_(a)) return; var g = String(a.group || '').trim() || '(ไม่ระบุอาคาร)'; assetG[a.asset_id] = g; (gmap[g] = gmap[g] || []).push(a); });
  var logsBy = {}; (logs || []).forEach(function (l) { var g = assetG[l.asset_id]; if (!g) return; var k = g + '|' + l.type_id; (logsBy[k] = logsBy[k] || []).push(l); });
  var groups = Object.keys(gmap).map(function (g) {
    var row = { name: g, assets: gmap[g].length, byType: {}, open: 0, overdue: 0, due: 0, done: 0 };
    act.forEach(function (t) {
      var as = gmap[g].filter(function (a) { return a.type_id === t.type_id; }); if (!as.length) return;
      var ls = logsBy[g + '|' + t.type_id] || [], comp = inspCompliance(inspMatrix(t, as, ls, fy, asOf, opt));
      var last = ''; ls.forEach(function (l) { if (l.date > last) last = l.date; });
      row.byType[t.type_id] = { pct: comp.pct, due: comp.due, done: comp.done, assets: as.length, last: last, lastAge: last ? inspDiff(asOf, last) : null };
      row.due += comp.due; row.done += comp.done; if (comp.pct !== null) (row.pcts = row.pcts || []).push(comp.pct);
    });
    (actions || []).forEach(function (x) { if (assetG[x.asset_id] === g && inspIsOpen(x)) { row.open++; if (x.due_date && x.due_date < asOf) row.overdue++; } });
    row.pct = row.pcts && row.pcts.length ? row.pcts.reduce(function (a, b) { return a + b; }, 0) / row.pcts.length : null;
    return row;
  }).sort(function (a, b) { return b.assets - a.assets; });
  return { groups: groups, types: act };
}

/**
 * ตัวชี้วัดของฝ่ายที่คำนวณจากผลตรวจเช็ค (KPI_Master.source = 'insp')  สูตร "งานที่บำรุงรักษาได้ / งานที่ต้องบำรุงรักษา" รายเดือน
 * - งานตรวจเช็คของประเภทที่ผูก kpi_id นี้: จุดตรวจ × รอบ ที่เริ่มแล้ว (เดือนปัจจุบัน = ความคืบหน้า ณ วันนี้)
 * - รายการแผน PM ที่ผูก kpi_id นี้แต่ยังไม่มีระบบตรวจเช็ค (รหัสไม่ตรงกับ pm_code ของประเภทใด): min(ผลจริง, แผน) / แผน
 * D = {types, assets, logs, items, plan}  plan = แถว PM_Plan {fy, month(1=ต.ค.), code, plan, actual}
 * คืน array 12 เดือน (index 0 = ต.ค.) ช่องละ null หรือ {num, den, pct, inProgress, parts:[{src:'insp'|'pm', id, label, done, exp}]}
 */
function inspKpiYear(kid, fy, D, asOf) {
  var out = [], mi;
  for (mi = 1; mi <= 12; mi++) out.push(null);
  var nz = function (x) { return x === null || x === undefined || x === '' ? null : Number(x); };
  var types = (D.types || []).filter(function (t) { return inspActive_(t) && String(t.kpi_id || '') === kid; });
  var covered = {};
  (D.types || []).forEach(function (t) { if (inspActive_(t) && t.pm_code) covered[String(t.pm_code)] = t.type_id; });
  var acc = []; for (mi = 1; mi <= 12; mi++) acc.push({ num: 0, den: 0, parts: [], any: false });
  types.forEach(function (t) {
    var as = (D.assets || []).filter(function (a) { return a.type_id === t.type_id; });
    var ls = (D.logs || []).filter(function (l) { return l.type_id === t.type_id; });
    var cells = inspMatrix(t, as, ls, fy, asOf), by = {};
    cells.forEach(function (c) {
      if (c.state === 'future' || c.state === 'na') return; // รอบที่ยังไม่สิ้นสุดนับตามที่ทำได้จริง (ค่าเดือนปัจจุบันจะเพิ่มขึ้นจนสิ้นเดือน)
      var b = by[c.mi] || (by[c.mi] = { done: 0, exp: 0 });
      b.done += Math.min(c.done, c.exp); b.exp += c.exp;
    });
    Object.keys(by).forEach(function (k) {
      if (D.ovPlan && t.pm_code && D.ovPlan[fy + '|' + k + '|' + t.pm_code]) return; // ผู้ดูแลระบบแก้ผลในแผนแล้ว -> ใช้ค่าจากแผนแทน
      var a = acc[k - 1], b = by[k]; if (!b.exp) return;
      a.num += b.done; a.den += b.exp; a.any = true;
      a.parts.push({ src: 'insp', id: t.type_id, label: t.name, done: b.done, exp: b.exp });
    });
  });
  var items = (D.items || []).filter(function (i) { return String(i.kpi_id || '') === kid && (i.active === undefined || i.active === null || isInspActive(i.active)); });
  if (items.length) {
    var ix = {};
    (D.plan || []).forEach(function (r) { if (Number(r.fy) === fy) ix[String(r.code) + '|' + Number(r.month)] = r; });
    items.forEach(function (it) {
      for (var m = 1; m <= 12; m++) {
        if (covered[String(it.code)] && !(D.ovPlan && D.ovPlan[fy + '|' + m + '|' + String(it.code)])) continue;
        var r = ix[String(it.code) + '|' + m]; if (!r) continue;
        var p = nz(r.plan), a = nz(r.actual); if (!(p > 0)) continue;
        var ym = inspYmOf(fy, m), end = ym.key + '-' + inspPad(inspEom(ym.y, ym.m));
        if (ym.key + '-01' > asOf) continue;
        if (a === null && end >= asOf) continue; // เดือนปัจจุบันที่ยังไม่รายงานผล
        var d = Math.min(Math.max(a || 0, 0), p), A = acc[m - 1];
        A.num += d; A.den += p; A.any = true;
        A.parts.push({ src: covered[String(it.code)] ? 'override' : 'pm', id: String(it.code), label: it.name, done: d, exp: p });
      }
    });
  }
  for (mi = 1; mi <= 12; mi++) {
    var A = acc[mi - 1]; if (!A.any || !A.den) continue;
    var ym2 = inspYmOf(fy, mi);
    out[mi - 1] = { num: Math.round(A.num * 100) / 100, den: Math.round(A.den * 100) / 100, pct: A.num / A.den * 100, inProgress: ym2.key + '-' + inspPad(inspEom(ym2.y, ym2.m)) >= asOf, parts: A.parts };
  }
  return out;
}
/** ข้อความสรุปที่มาของค่า (เก็บในหมายเหตุ KPI_Data) */
function inspKpiNote(cell) {
  if (!cell) return '';
  var a = cell.parts.filter(function (p) { return p.src === 'insp'; }), b = cell.parts.filter(function (p) { return p.src === 'pm'; }), o = cell.parts.filter(function (p) { return p.src === 'override'; });
  var s = 'คำนวณจากผลตรวจเช็ค: ' + a.map(function (p) { return p.label.replace(/\s*\(.*\)$/, '') + ' ' + p.done + '/' + p.exp; }).join(', ');
  if (b.length) { var d = 0, e = 0; b.forEach(function (p) { d += p.done; e += p.exp; }); s += ' · แผน PM ที่ยังไม่มีระบบตรวจเช็ค ' + b.length + ' รายการ ' + d + '/' + e; }
  if (o.length) s += ' · ค่าที่ผู้ดูแลระบบแก้ในแผน: ' + o.map(function (p) { return p.id + ' ' + p.done + '/' + p.exp; }).join(', ');
  if (cell.inProgress) s += ' · เดือนนี้ยังไม่สิ้นสุด';
  return s.slice(0, 480);
}

/**
 * แผน -> ตรวจเช็ค: ทำเครื่องหมายเดือนที่แผน PM ของรหัสที่ผูกกับประเภทงาน "ไม่มีแผน/แผน = 0" (เฉพาะปีงบที่มีการกำหนดแผนของรหัสนั้นแล้ว)
 * เดือนเหล่านั้นไม่ถือว่าต้องตรวจ (ไม่แจ้งเตือน ไม่นับใน KPI)
 */
function inspAttachPlan(types, plan) {
  var by = {};
  (plan || []).forEach(function (r) { var c = String(r.code); (by[c] = by[c] || []).push(r); });
  (types || []).forEach(function (t) {
    t._planOff = {};
    var rows = t.pm_code ? by[String(t.pm_code)] : null; if (!rows) return;
    var fys = {}, has = {};
    rows.forEach(function (r) { var f = Number(r.fy); fys[f] = 1; if (Number(r.plan) > 0) has[f + '|' + Number(r.month)] = 1; });
    Object.keys(fys).forEach(function (f) { for (var m = 1; m <= 12; m++) if (!has[f + '|' + m]) t._planOff[f + '|' + m] = true; });
  });
  return types;
}

/**
 * ตรวจเช็ค -> แผน: ผลจริงของเดือน (fy, mi) ของประเภทงาน ตามวิธีนับ
 * assets = จำนวนจุดตรวจที่ตรวจแล้ว / ที่ต้องตรวจ (ค่าเริ่มต้น) · logs = จุดตรวจ × รอบ · cycles = จำนวนรอบที่ครบ / จำนวนรอบ
 * คืน null ถ้าเดือนนั้นยังไม่ถึง/ไม่ต้องตรวจ
 */
function inspPlanMonth(type, assets, logs, fy, mi, asOf, mode) {
  var cells = inspMatrix(type, assets, logs, fy, asOf).filter(function (c) { return c.mi === mi && c.state !== 'future' && c.state !== 'na'; });
  if (!cells.length) return null;
  mode = mode || 'assets';
  if (mode === 'cycles') return { exp: cells.length, done: cells.filter(function (c) { return c.state === 'done'; }).length };
  if (mode === 'logs') { var e = 0, d = 0; cells.forEach(function (c) { e += c.exp; d += Math.min(c.done, c.exp); }); return { exp: e, done: d }; }
  var ym = inspYmOf(fy, mi).key, due = {}, n = 0, done = 0;
  (assets || []).forEach(function (a) {
    if (!isInspActive(a.active === undefined ? 1 : a.active)) return;
    if (a.since && String(a.since).slice(0, 7) > ym) return;
    if (a.until && String(a.until).slice(0, 7) < ym) return;
    due[a.asset_id] = 1; n++;
  });
  var seen = {};
  (logs || []).forEach(function (l) { if (String(l.period).slice(0, 7) === ym && due[l.asset_id] && !seen[l.asset_id]) { seen[l.asset_id] = 1; done++; } });
  return n ? { exp: n, done: done } : null;
}
// ==INSP-CORE-END==

// ==INSP-SERVER-START==
/* ====================================================================
 * ระบบตรวจเช็คบำรุงรักษา (ถังดับเพลิง ตู้น้ำดับเพลิง ประตูหนีไฟ ไฟฉุกเฉิน Generator ระบบเสียงตามสาย)
 * ชีต: Insp_Types, Insp_Checks, Insp_Assets, Insp_Log, Insp_Actions
 * ==================================================================== */
const INSP_TEXT_COLS = {
  Insp_Types: ['type_id', 'pm_code'],
  Insp_Checks: ['type_id', 'check_id'],
  Insp_Assets: ['asset_id', 'type_id', 'code', 'since', 'until'],
  Insp_Log: ['log_id', 'type_id', 'asset_id', 'date', 'period'],
  Insp_Actions: ['action_id', 'type_id', 'asset_id', 'check_id', 'found_date', 'last_seen', 'due_date', 'done_date', 'verified_ok_date'],
  Auto_Override: ['key']
};
const INSP_BOOL_SETTINGS = ['insp_entry', 'insp_mail', 'insp_report_on'];
const INSP_NUM_SETTINGS = ['insp_sla_high', 'insp_sla_normal'];
const INSP_ACTION_STATUS = ['open', 'in_progress', 'done', 'cancelled'];

function todayStr_() { return new Date(Date.now() + 7 * 3600 * 1000).toISOString().slice(0, 10); }

/** เติมข้อมูลตั้งต้นของระบบตรวจเช็ค (ทำครั้งเดียวต่อฐานข้อมูล) */
function inspSeed_(ss) {
  const st = ss.getSheetByName('Settings');
  const ti = ss.getSheetByName('Insp_Types');
  if (!st || !ti) return false;
  const sv = st.getDataRange().getValues();
  let at = -1;
  sv.forEach(function (r, i) { if (r[0] === 'insp_seeded') at = i; });
  const cur = at >= 0 ? (Number(sv[at][1]) || 1) : 0;
  if (typeof INSP_SEED === 'undefined') {
    if (cur) return false;
    throw new Error('ไม่พบไฟล์ InspSeed.gs ในโปรเจกต์ กรุณาเพิ่มไฟล์ InspSeed.gs (สร้างไฟล์สคริปต์ใหม่ชื่อ InspSeed แล้ววางเนื้อหา) จากนั้น Deploy > Manage deployments > Edit > New version');
  }
  const ver = Number(INSP_SEED.version) || 1;
  if (cur >= ver) return false;
  // เพิ่มเฉพาะแถวที่ยังไม่มี (ไม่ทับค่าที่ผู้ดูแลระบบแก้ไว้)  ทรัพย์สินเดิมที่ยังไม่มีอาคาร (group) เติมให้
  const keyCols = { Insp_Types: ['type_id'], Insp_Checks: ['type_id', 'check_id'], Insp_Assets: ['asset_id'] };
  const added = {};
  Object.keys(keyCols).forEach(function (n) {
    const sh = ss.getSheetByName(n), rows = INSP_SEED[n] || [], head = SHEETS[n];
    if (!rows.length) return;
    if (sh.getLastRow() < 2) { put_(ss, n, rows); added[n] = rows.length; return; }
    const v = sh.getDataRange().getValues(), hdr = v[0].map(String);
    const ki = keyCols[n].map(function (k) { return hdr.indexOf(k); }), si = keyCols[n].map(function (k) { return head.indexOf(k); });
    const have = {};
    for (let i = 1; i < v.length; i++) have[ki.map(function (c) { return String(v[i][c]); }).join('|')] = i;
    const gi = hdr.indexOf('group'), sgi = head.indexOf('group');
    const add = [];
    rows.forEach(function (r) {
      const k = si.map(function (c) { return String(r[c]); }).join('|');
      if (have[k] === undefined) { add.push(hdr.map(function (h) { const j = head.indexOf(h); return j < 0 ? '' : r[j]; })); return; }
      if (n === 'Insp_Assets' && gi >= 0 && sgi >= 0 && r[sgi] && !String(v[have[k]][gi]).trim()) sh.getRange(have[k] + 1, gi + 1).setValue(r[sgi]);
      if (n === 'Insp_Types') { const di = hdr.indexOf('standard'), sdi = head.indexOf('standard'); if (di >= 0 && sdi >= 0 && r[sdi] && !String(v[have[k]][di]).trim()) sh.getRange(have[k] + 1, di + 1).setValue(r[sdi]); }
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, hdr.length).setValues(add);
    added[n] = add.length;
  });
  // ตั้งให้ KPI ที่กำหนดใน INSP_SEED.kpiSource คำนวณจากผลตรวจเช็ค (ครั้งเดียว และเฉพาะที่ยังเป็น "กรอกเอง")
  const ksrc = INSP_SEED.kpiSource || {}, kv = Object.keys(ksrc);
  if (kv.length && cur < 3) {
    const km = ss.getSheetByName('KPI_Master'), mv = km.getDataRange().getValues(), mh = mv[0].map(String), iK = mh.indexOf('kpi_id'), iS = mh.indexOf('source');
    if (iS >= 0) for (let i = 1; i < mv.length; i++) {
      const id = String(mv[i][iK]);
      if (ksrc[id] && ['', 'manual'].indexOf(String(mv[i][iS])) >= 0) { km.getRange(i + 1, iS + 1).setValue(ksrc[id]); added.kpi = (added.kpi || 0) + 1; }
    }
  }
  // รุ่น 4: เชื่อม ตรวจเช็ค -> แผน อัตโนมัติ (pm_sync จาก off เป็น assets) และเพิ่มรายการแผน "ตรวจเช็คไฟฉุกเฉิน" ให้ประเภทที่ยังไม่มีรหัส
  if (cur < 4 && INSP_SEED.link) {
    const L = INSP_SEED.link;
    added.items = inspEnsureLinkItems_(ss);
    const haveI = {}; read_('PM_Items').forEach(function (i) { haveI[norm_(i.code)] = 1; });
    const tsh = ss.getSheetByName('Insp_Types'), tv = tsh.getDataRange().getValues(), th = tv[0].map(String);
    const iS = th.indexOf('pm_sync'), iC = th.indexOf('pm_code'), iT = th.indexOf('type_id');
    for (let i = 1; i < tv.length; i++) {
      const id = String(tv[i][iT]);
      if (L.pmCode && L.pmCode[id] && !String(tv[i][iC]).trim() && haveI[L.pmCode[id]]) { tsh.getRange(i + 1, iC + 1).setValue(L.pmCode[id]); tv[i][iC] = L.pmCode[id]; }
      if (L.sync && iS >= 0 && ['', 'off'].indexOf(String(tv[i][iS])) >= 0 && String(tv[i][iC]).trim()) { tsh.getRange(i + 1, iS + 1).setValue(L.sync); added.link = (added.link || 0) + 1; }
    }
  }
  if (at >= 0) st.getRange(at + 1, 2).setValue(String(ver)); else st.appendRow(['insp_seeded', String(ver)]);
  ss.getSheetByName('Audit_Log').appendRow([new Date().toISOString(), Session.getEffectiveUser().getEmail(), 'admin', cur ? 'ปรับปรุงข้อมูลตั้งต้นระบบตรวจเช็ค v' + ver : 'ตั้งต้นระบบตรวจเช็ค', 'setup', '', '', '', 'เพิ่มประเภท ' + (added.Insp_Types || 0) + ' · รายการตรวจ ' + (added.Insp_Checks || 0) + ' · จุดตรวจ ' + (added.Insp_Assets || 0) + (added.kpi ? ' · KPI ใช้ผลตรวจเช็ค ' + added.kpi : '') + (added.link ? ' · เชื่อมเข้าแผน ' + added.link + ' ประเภท' : '') + (added.items ? ' · รายการแผนใหม่ ' + added.items : '')]);
  return true;
}

/** รายการแผนที่ระบบตรวจเช็คต้องใช้ (เช่น 2004003 ตรวจเช็คไฟฉุกเฉิน) ต้องมีอยู่เสมอ — เติมให้ถ้าหายไป (รันทุกครั้งที่เปิดฐานข้อมูล) */
function inspEnsureLinkItems_(ss) {
  if (typeof INSP_SEED === 'undefined' || !INSP_SEED.link || !(INSP_SEED.link.items || []).length) return 0;
  const ish = ss.getSheetByName('PM_Items'); if (!ish || ish.getLastRow() < 2) return 0;
  const iv = ish.getDataRange().getValues(), ih = iv[0].map(String), have = {}, ci = ih.indexOf('code');
  for (let i = 1; i < iv.length; i++) have[norm_(iv[i][ci])] = 1;
  const add = INSP_SEED.link.items.filter(function (o) { return !have[o.code]; }).map(function (o) { return ih.map(function (h) { return o[h] === undefined ? '' : o[h]; }); });
  if (add.length) ish.getRange(ish.getLastRow() + 1, 1, add.length, ih.length).setValues(add);
  return add.length;
}

function inspSettings_() {
  const s = {};
  read_('Settings').forEach(function (r) { s[r.key] = r.value; });
  const on = function (v) { return v === 1 || v === true || ['1', 'true', 'on', 'เปิด'].indexOf(String(v).toLowerCase()) >= 0; };
  return { entryOpen: on(s.insp_entry), mail: on(s.insp_mail), slaHigh: Number(s.insp_sla_high) || 7, slaNormal: Number(s.insp_sla_normal) || 30 };
}

function inspTypes_(plan) {
  const ts = read_('Insp_Types').map(function (t) { t.type_id = String(t.type_id); t.pm_code = t.pm_code ? norm_(t.pm_code) : ''; t.pm_sync = t.pm_sync || 'off'; return t; })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  return inspAttachPlan(ts, (plan || read_('PM_Plan')).map(function (r) { return { fy: r.fy, month: r.month, code: norm_(r.code), plan: r.plan }; }));
}
function inspChecks_() {
  return read_('Insp_Checks').map(function (c) { c.type_id = String(c.type_id); c.check_id = String(c.check_id); return c; })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
}
function inspAssets_() {
  return read_('Insp_Assets').map(function (a) { a.asset_id = String(a.asset_id); a.type_id = String(a.type_id); a.code = a.code === null ? '' : String(a.code); a.since = a.since ? String(a.since) : ''; a.until = a.until ? String(a.until) : ''; return a; });
}
function inspChecksById_(checks) {
  const m = {};
  checks.forEach(function (c) { (m[c.type_id] = m[c.type_id] || {})[c.check_id] = c; });
  return m;
}

/** ชุดข้อมูลตรวจเช็คที่หน้าเว็บใช้ (บันทึกย้อนหลังประมาณ 25 เดือน งานแก้ไขที่ยังเปิดทั้งหมด) */
function inspBundle_() {
  const today = todayStr_(), cut = inspAdd(today, -760), cutRes = inspAdd(today, -62);
  // ผลรายข้อส่งเฉพาะ 62 วันล่าสุด (บันทึกรายวัน/ข้อมูลหลายอาคารทำให้ข้อมูลใหญ่) ที่เก่ากว่าใช้เพียงสถานะสรุป
  const raw = read_('Insp_Log');
  const logs = raw.filter(function (l) { return String(l.date) >= cut; }).map(function (l) {
    return { log_id: l.log_id, type_id: String(l.type_id), asset_id: String(l.asset_id), date: String(l.date), period: String(l.period), inspector: l.inspector || '', st: l.st, ng_n: Number(l.ng_n) || 0, wn_n: Number(l.wn_n) || 0, results: String(l.date) >= cutRes ? (l.results || '{}') : '', note: l.note || '', source: l.source || '' };
  });
  const actions = read_('Insp_Actions').filter(function (a) { return a.status === 'open' || a.status === 'in_progress' || String(a.found_date) >= cut; });
  const set = inspSettings_();
  let types = inspTypes_(), assets = inspAssets_(), logsOut = logs, actsOut = actions, scoped = false;
  const cu = currentUser_();
  if (cu.role === 'inspector') {
    // ผู้ตรวจเช็ค: ส่งเฉพาะอาคาร×ระบบที่ตนรับผิดชอบ (ภาพรวม/ผลตรวจ/งานแก้ไขจะถูกจำกัดตามไปด้วย)
    const tById = {}; types.forEach(function (t) { tById[t.type_id] = t; });
    const okA = {};
    assets = assets.filter(function (a) { const ok = canInspect_(cu, tById[a.type_id], null, a); if (ok) okA[a.asset_id] = 1; return ok; });
    const okT = {}; assets.forEach(function (a) { okT[a.type_id] = 1; });
    types = types.filter(function (t) { return okT[t.type_id]; });
    logsOut = logs.filter(function (l) { return okA[l.asset_id]; });
    actsOut = actions.filter(function (a) { return okA[a.asset_id]; });
    scoped = true;
  }
  return { scoped: scoped, today: today, types: types, checks: inspChecks_(), assets: assets, logs: logsOut, actions: actsOut, sla: { high: set.slaHigh, normal: set.slaNormal }, entryOpen: set.entryOpen, mail: set.mail, resp: inspResp_(), kpiFrom: inspKpiFrom_(), kpiEvid: scoped ? {} : inspKpiEvidence_(raw, today), overrides: read_('Auto_Override').map(function (r) { return { target: r.target, key: norm_(r.key), by: r.by, at: r.at, note: r.note || '' }; }) };
}

/**
 * หลักฐานจากผลตรวจเช็คสำหรับ KPI ที่ยังกรอกเอง (ไม่เขียนทับ): รายเดือน ปีงบปัจจุบันและปีก่อน
 * K4 = ค่าน้ำทิ้งที่วัดหน้างานผ่านเกณฑ์ (pH น้ำออก 5–9, คลอรีน 0.5–2) · K5 = ทดสอบเดินเครื่อง Auto แล้วจ่ายไฟช้ากว่า 15 วินาที
 */
const INSP_KPI_EVID = [
  { kpi_id: 'K4', type_id: 'T09', checks: ['w9', 'w16'], mode: 'pass', label: 'ค่าน้ำทิ้งที่วัดหน้างานผ่านเกณฑ์ (pH น้ำออก, คลอรีน)' },
  { kpi_id: 'K5', type_id: 'T05', checks: ['g12'], mode: 'over', limit: 15, label: 'ทดสอบ Generator ระบบ Auto จ่ายไฟช้ากว่า 15 วินาที (ครั้ง)' }
];
function inspKpiEvidence_(raw, today) {
  const fyNow = inspFyOf(today), ck = {}, out = { defs: INSP_KPI_EVID, data: {} };
  inspChecks_().forEach(function (c) { ck[c.type_id + '|' + c.check_id] = c; });
  INSP_KPI_EVID.forEach(function (e) {
    const box = out.data[e.kpi_id] = {};
    raw.forEach(function (l) {
      if (String(l.type_id) !== e.type_id) return;
      const d = String(l.date), fy = inspFyOf(d); if (fy !== fyNow && fy !== fyNow - 1) return;
      let r; try { r = JSON.parse(l.results || '{}'); } catch (x) { return; }
      const k = fy + '|' + inspMiOf(+d.slice(5, 7)), b = box[k] || (box[k] = { n: 0, hit: 0 });
      e.checks.forEach(function (cid) {
        const v = r[cid]; if (typeof v !== 'number') return;
        b.n++;
        if (e.mode === 'pass') { if (inspEvalCheck(ck[e.type_id + '|' + cid] || { kind: 'n' }, v).st === 'ok') b.hit++; }
        else if (v > e.limit) b.hit++;
      });
    });
  });
  return out;
}

/**
 * สิทธิ์บันทึกผลตรวจ/แก้งานของประเภท (และจุดตรวจ ถ้าระบุ)
 * ผู้ดูแลระบบ: ทั้งหมด · ผู้ตรวจเช็ค: ทุกประเภท/อาคาร หรือเฉพาะที่กำหนดในผู้ใช้ (insp_types / insp_groups)
 * ผู้บันทึกข้อมูล: หน่วยซ่อมบำรุง หรือเป็นเจ้าของประเภท/ผู้รับผิดชอบรายการแผน/ระบบ หรือได้รับกำหนดประเภทไว้ · ทุกบทบาท: ผู้รับผิดชอบจุดตรวจนั้น
 */
function canInspect_(u, type, pmItem, asset) {
  if (u.role === 'admin') return true;
  if (u.role !== 'editor' && u.role !== 'inspector') return false;
  const em = String(u.email || '').toLowerCase();
  if (asset && em && String(asset.responsible_email || '').trim().toLowerCase() === em) return true;
  const rs = u.inspResp || [];
  if (rs.length) {
    const gg = asset ? String(asset.group || '').trim() : '';
    return rs.some(function (r) { return (r.type_id === '*' || !type || r.type_id === type.type_id) && (r.group === '*' || !asset || r.group === gg); });
  }
  const tScope = u.inspTypes || [], gScope = u.inspGroups || [];
  if (type && tScope.length && tScope.indexOf(type.type_id) < 0) return false;
  if (asset && gScope.length && gScope.indexOf(String(asset.group || '').trim()) < 0) return false;
  if (u.role === 'inspector') return true;
  if (tScope.length || gScope.length) return true;
  if (canEditUnit_(u, PM_UNIT)) return true;
  if (!em) return false;
  if (type && String(type.owner_email || '').trim().toLowerCase() === em) return true;
  return !!pmItem && String(pmItem.responsible_email || '').trim().toLowerCase() === em;
}

function inspPmItem_(type) {
  return type && type.pm_code ? read_('PM_Items').filter(function (i) { return norm_(i.code) === norm_(type.pm_code); })[0] : null;
}

/**
 * ผู้ปฏิบัติงานบันทึกผลตรวจเช็ค
 * p: {type_id, date:'YYYY-MM-DD', week?:1-4 (ประเภทรายสัปดาห์), inspector, entries:[{asset_id, results:{check_id:ค่า}, note}]}
 * - ผู้มีสิทธิ์ (ผู้ดูแลระบบ / หน่วยซ่อมบำรุง / เจ้าของประเภทหรือผู้รับผิดชอบรายการแผน) บันทึกได้เสมอ
 * - คนอื่นบันทึกได้เมื่อผู้ดูแลเปิด insp_entry และต้องใส่ชื่อผู้ตรวจ
 * - รายการที่ไม่ปกติต้องมีหมายเหตุ ระบบสร้าง/ต่อยอด "งานแก้ไข" ให้อัตโนมัติ และตรวจซ้ำปกติจะติดป้าย "รอยืนยันปิดงาน"
 */
function saveInspection(p) {
  const u = currentUser_();
  const type = inspTypes_().filter(function (t) { return t.type_id === String(p && p.type_id) && isActive_(t.active); })[0];
  if (!type) throw new Error('ไม่พบประเภทการตรวจ หรือถูกปิดใช้งาน');
  const pm = inspPmItem_(type);
  const allAs = inspAssets_().filter(function (a) { return a.type_id === type.type_id; }), asBy = {};
  allAs.forEach(function (a) { asBy[a.asset_id] = a; });
  const ents = (p.entries || []);
  // มีสิทธิ์เมื่อมีสิทธิ์กับทุกจุดที่ส่งมา (ผู้ตรวจที่จำกัดอาคารบันทึกได้เฉพาะอาคารนั้น)
  const privileged = ents.length ? ents.every(function (e) { return canInspect_(u, type, pm, asBy[String(e.asset_id)]); }) : canInspect_(u, type, pm);
  const set = inspSettings_();
  if (!privileged && !set.entryOpen) throw new Error('ยังไม่เปิดให้บันทึกผลตรวจหน้างาน หรือไม่มีสิทธิ์ ให้ผู้ดูแลระบบเปิด "บันทึกผลตรวจหน้างาน" หรือเพิ่มสิทธิ์ให้บัญชีนี้');
  const inspector = String(p.inspector || '').trim().slice(0, 60);
  if (!privileged && !inspector) throw new Error('กรุณาใส่ชื่อผู้ตรวจ');
  const who = privileged ? (u.email || u.name) : ('หน้างาน: ' + inspector);
  const insName = inspector || u.name || u.email || '';
  const date = String(p.date || '');
  if (!inspParse(date)) throw new Error('วันที่ตรวจไม่ถูกต้อง');
  const today = todayStr_();
  if (date > today) throw new Error('บันทึกวันที่ล่วงหน้าไม่ได้');
  if (u.role !== 'admin') {
    if (inspDiff(today, date) > 62) throw new Error('บันทึกย้อนหลังได้ไม่เกิน 62 วัน ถ้าจำเป็นให้ผู้ดูแลระบบบันทึกให้');
    const fy = inspFyOf(date), mi = inspMiOf(+date.slice(5, 7));
    if (lockSet_()[fy + '|' + mi]) throw new Error('เดือนของวันที่นี้ถูกล็อกแล้ว ขอให้ผู้ดูแลระบบปลดล็อกก่อน');
  }
  const entries = (p.entries || []);
  if (!entries.length) throw new Error('ไม่มีรายการที่จะบันทึก');
  if (entries.length > 300) throw new Error('บันทึกได้ครั้งละไม่เกิน 300 รายการ');
  const checks = inspChecks_().filter(function (c) { return c.type_id === type.type_id; });
  const allChecks = inspChecksById_(checks);
  const assets = inspAssets_().filter(function (a) { return a.type_id === type.type_id; });
  const assetBy = {}; assets.forEach(function (a) { assetBy[a.asset_id] = a; });
  const period = inspPeriodOf(type.freq, date, p.week);
  const required = checks.filter(function (c) { return c.kind === 's' || c.kind === 'm' || (c.kind === 'n' && inspRange(c).has); });
  // รายการที่ใช้กับจุดนั้น = รายการที่มีผลในบันทึกล่าสุดของจุดนั้น (เช่น บ่อบำบัดบางชุดไม่มีเครื่องโอโซน) ถ้าไม่เคยบันทึกใช้ทุกรายการ
  const want = {}; entries.forEach(function (e) { want[String(e.asset_id)] = 1; });
  const lastRes = {};
  read_('Insp_Log').forEach(function (l) {
    const aid = String(l.asset_id); if (!want[aid] || String(l.type_id) !== type.type_id) return;
    const d = String(l.date); if (lastRes[aid] && lastRes[aid].date >= d) return;
    let r = null; try { r = JSON.parse(l.results || '{}'); } catch (x) { r = null; }
    if (r && Object.keys(r).length) lastRes[aid] = { date: d, ids: Object.keys(r) };
  });
  const seen = {}, rows = [];
  entries.forEach(function (e) {
    const a = assetBy[String(e.asset_id)];
    if (!a) throw new Error('ไม่พบทรัพย์สิน ' + e.asset_id + ' ในประเภทนี้');
    if (!isActive_(a.active)) throw new Error('ทรัพย์สิน "' + a.name + '" ถูกปิดใช้งาน');
    if (seen[a.asset_id]) throw new Error('ทรัพย์สิน "' + a.name + '" ซ้ำในรายการ');
    seen[a.asset_id] = 1;
    const ev = inspEvalEntry(checks, e.results || {});
    const lr = lastRes[String(e.asset_id)];
    const miss = required.filter(function (c) { const v = ev.res[c.check_id]; return (!lr || lr.ids.indexOf(c.check_id) >= 0) && (v === undefined || v === ''); });
    if (miss.length) throw new Error('"' + a.name + '" ยังไม่ได้ตรวจ: ' + miss.slice(0, 4).map(function (c) { return c.label; }).join(', ') + (miss.length > 4 ? ' ฯลฯ' : ''));
    const note = String(e.note || '').trim().slice(0, 200);
    if (ev.ngN && !note) throw new Error('"' + a.name + '" พบรายการไม่ปกติ กรุณาระบุรายละเอียดที่พบในช่องหมายเหตุ');
    rows.push({ a: a, ev: ev, note: note });
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = db_();
    const sh = ss.getSheetByName('Insp_Log'), head = SHEETS.Insp_Log;
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
    const iId = head.indexOf('log_id'), iAs = head.indexOf('asset_id'), iDt = head.indexOf('date');
    const idx = {}, lastDate = {};
    vals.forEach(function (r, i) {
      idx[String(r[iId])] = i;
      const a = String(r[iAs]), d = String(r[iDt]);
      if (!lastDate[a] || d > lastDate[a]) lastDate[a] = d;
    });
    const now = new Date().toISOString();
    const appendRows = [], logsForActions = [], audits = [];
    let updated = 0, ngTotal = 0;
    rows.forEach(function (r) {
      const id = r.a.asset_id + '|' + date;
      const rec = { log_id: id, type_id: type.type_id, asset_id: r.a.asset_id, date: date, period: period, inspector: insName, st: r.ev.st, ng_n: r.ev.ngN, wn_n: r.ev.wnN,
        results: JSON.stringify(r.ev.res), note: r.note, source: privileged ? 'app' : 'field', created_by: who, created_at: now };
      const line = head.map(function (h) { return rec[h] === undefined || rec[h] === null ? '' : rec[h]; });
      if (idx[id] !== undefined) { sh.getRange(idx[id] + 2, 1, 1, head.length).setValues([line]); updated++; } else appendRows.push(line);
      ngTotal += r.ev.ngN;
      if (!lastDate[r.a.asset_id] || date >= lastDate[r.a.asset_id]) logsForActions.push({ log_id: id, type_id: type.type_id, asset_id: r.a.asset_id, date: date, note: r.note, results: r.ev.res });
    });
    if (appendRows.length) sh.getRange(sh.getLastRow() + 1, 1, appendRows.length, head.length).setValues(appendRows);
    audits.push({ action: 'บันทึกผลตรวจ', sheet: 'Insp_Log', key: type.type_id + '|' + date, field: 'รายการ', old: '', new: rows.length + ' ทรัพย์สิน' + (ngTotal ? ' · ไม่ปกติ ' + ngTotal : '') + (updated ? ' · แก้ของเดิม ' + updated : '') });

    // ---- งานแก้ไข ----
    const act = inspApplyActions_(ss, type, pm, assetBy, allChecks, logsForActions, set, who);
    audit_({ email: who, name: who, role: privileged ? u.role : 'field' }, audits.concat(act.audit));
    const done = inspPeriodDone_(type, assets, period);
    const link = syncLinks_(inspFyOf(date), [inspMiOf(+date.slice(5, 7))], who);
    const sync = link.plan, kpi = link.kpi;
    return { saved: rows.length, updated: updated, ng: ngTotal, period: period, newActions: act.created, touched: act.touched, done: done.done, exp: done.exp, sync: sync, kpi: kpi };
  } finally { lock.releaseLock(); }
}

/** จำนวนทรัพย์สินที่ตรวจแล้ว/ที่ต้องตรวจ ในรอบนั้น */
function inspPeriodDone_(type, assets, period) {
  const logs = read_('Insp_Log').filter(function (l) { return String(l.type_id) === type.type_id && String(l.period) === period; });
  const cells = inspMatrix(type, assets, logs.map(function (l) { return { asset_id: String(l.asset_id), period: String(l.period), st: l.st }; }), inspFyOf(period.slice(0, 7) + '-01'), '9999-12-31');
  const c = cells.filter(function (x) { return x.period === period; })[0];
  return c ? { done: c.done, exp: c.exp } : { done: 0, exp: 0 };
}

function inspApplyActions_(ss, type, pm, assetBy, allChecks, newLogs, set, who) {
  const sh = ss.getSheetByName('Insp_Actions'), head = SHEETS.Insp_Actions;
  const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, ci) { o[h] = r[ci] === '' ? '' : r[ci]; });
    return o;
  }) : [];
  const snap = function (a) { return JSON.stringify([a.seen_count, a.last_seen, a.status, a.action_note, a.verified_ok_date, a.severity, a.issue]); };
  const before = {};
  existing.forEach(function (a) { a.seen_count = Number(a.seen_count) || 1; before[a.action_id] = snap(a); });
  const all = inspDeriveActions(allChecks, assetBy, newLogs, existing.filter(function (a) { return a.type_id === type.type_id; }), { slaHigh: set.slaHigh, slaNormal: set.slaNormal });
  const now = new Date().toISOString();
  const adds = [], audits = [];
  let touched = 0;
  all.forEach(function (a) {
    if (a._row) {
      if (before[a.action_id] !== snap(a)) {
        a.updated_by = who; a.updated_at = now;
        sh.getRange(a._row, 1, 1, head.length).setValues([head.map(function (h) { return a[h] === undefined || a[h] === null ? '' : a[h]; })]);
        touched++;
      }
      return;
    }
    const as = assetBy[a.asset_id] || {};
    const w = inspAssignees_(type, as, pm);
    a.assignee = w.name; a.assignee_email = w.email;
    a.issue = (as.name ? as.name + ': ' : '') + a.issue;
    a.updated_by = who; a.updated_at = now;
    adds.push(head.map(function (h) { return a[h] === undefined || a[h] === null ? '' : a[h]; }));
    audits.push({ action: 'สร้างงานแก้ไข', sheet: 'Insp_Actions', key: a.action_id, field: 'issue', old: '', new: a.issue });
  });
  if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, head.length).setValues(adds);
  return { created: adds.length, touched: touched, audit: audits };
}

/** แก้ไขงานแก้ไข rows: [{action_id, assignee, assignee_email, due_date, status, action_note, severity}] */
function saveActions(rows) {
  const u = requireEditor_(true);
  const types = {}; inspTypes_().forEach(function (t) { types[t.type_id] = t; });
  const assetBy = {}; inspAssets_().forEach(function (a) { assetBy[a.asset_id] = a; });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = db_().getSheetByName('Insp_Actions'), head = SHEETS.Insp_Actions;
    const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
    const iId = head.indexOf('action_id');
    const idx = {}; vals.forEach(function (r, i) { idx[String(r[iId])] = i; });
    const today = todayStr_(), now = new Date().toISOString();
    const plan = [];
    (rows || []).forEach(function (r) {
      const i = idx[String(r.action_id)];
      if (i === undefined) throw new Error('ไม่พบงานแก้ไข ' + r.action_id);
      const cur = {}; head.forEach(function (h, ci) { cur[h] = vals[i][ci]; });
      const type = types[String(cur.type_id)];
      const mine = (u.email && String(cur.assignee_email || '').toLowerCase().split(/[,;\s]+/).indexOf(String(u.email).toLowerCase()) >= 0) || (u.staffName && String(cur.assignee || '').split(/\s*,\s*/).indexOf(u.staffName) >= 0);
      if (!mine && !canInspect_(u, type, inspPmItem_(type), assetBy[String(cur.asset_id)])) throw new Error('ไม่มีสิทธิ์แก้ไขงานของ "' + (type ? type.name : cur.type_id) + '"');
      const nx = Object.assign({}, cur);
      ['assignee', 'assignee_email', 'action_note'].forEach(function (f) { if (r[f] !== undefined) nx[f] = String(r[f] === null ? '' : r[f]).trim().slice(0, f === 'action_note' ? 300 : 200); });
      if (r.severity !== undefined) { if (['high', 'normal'].indexOf(r.severity) < 0) throw new Error('ความรุนแรงไม่ถูกต้อง'); nx.severity = r.severity; }
      if (r.due_date !== undefined) { if (r.due_date && !inspParse(r.due_date)) throw new Error('วันครบกำหนดไม่ถูกต้อง'); nx.due_date = r.due_date || ''; }
      if (nx.assignee_email && String(nx.assignee_email).split(/[,;]+/).some(function (e) { return e.trim() && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e.trim()); })) throw new Error('อีเมลผู้รับผิดชอบไม่ถูกต้อง');
      if (r.status !== undefined) {
        if (INSP_ACTION_STATUS.indexOf(r.status) < 0) throw new Error('สถานะไม่ถูกต้อง');
        nx.status = r.status;
        if (r.status === 'done') {
          if (!String(nx.action_note || '').trim()) throw new Error('กรุณาระบุผลการแก้ไขก่อนปิดงาน (' + cur.action_id + ')');
          if (!nx.done_date) nx.done_date = today;
        } else if (r.status === 'open' || r.status === 'in_progress') nx.done_date = '';
      }
      const diff = head.filter(function (h) { return ['updated_by', 'updated_at'].indexOf(h) < 0 && norm_(cur[h]) !== norm_(nx[h]); });
      if (diff.length) plan.push({ i: i, nx: nx, diff: diff, cur: cur });
    });
    const logs = [];
    plan.forEach(function (p) {
      p.nx.updated_by = u.email || u.name; p.nx.updated_at = now;
      sh.getRange(p.i + 2, 1, 1, head.length).setValues([head.map(function (h) { return p.nx[h] === undefined || p.nx[h] === null ? '' : p.nx[h]; })]);
      p.diff.forEach(function (f) { logs.push({ action: 'แก้ไข', sheet: 'Insp_Actions', key: p.cur.action_id, field: f, old: norm_(p.cur[f]), new: norm_(p.nx[f]) }); });
    });
    audit_(u, logs);
    return { saved: plan.length, unchanged: (rows || []).length - plan.length };
  } finally { lock.releaseLock(); }
}

/* ====== การไหลของข้อมูล: ตรวจเช็ค -> แผนบำรุงรักษา -> KPI  (ผู้ดูแลระบบแก้/ลบได้ ค่าที่แก้จะ "ล็อกไว้" ไม่ถูกคำนวณทับ) ====== */
const INSP_SYNC_MODES = ['off', 'assets', 'logs', 'cycles'];
const INSP_SYS_WHO = 'ระบบตรวจเช็ค';
/** ช่องที่ผู้ดูแลระบบแก้เอง: target = PM_Plan | KPI_Data, key = fy|month|code หรือ fy|month|kpi_id */
function inspOverrides_() {
  const o = {};
  read_('Auto_Override').forEach(function (r) { o[r.target + '|' + norm_(r.key)] = r; });
  return o;
}
function inspAddOverrides_(u, list) {
  if (!list.length) return;
  const have = inspOverrides_(), sh = db_().getSheetByName('Auto_Override'), now = new Date().toISOString(), add = [];
  list.forEach(function (x) { const k = x.target + '|' + x.key; if (!have[k]) { have[k] = 1; add.push([x.target, x.key, u.email || u.name, now, x.note || '']); } });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
}
/** รหัสแผนที่รับผลจากระบบตรวจเช็ค: code -> ประเภทงาน */
function inspLinkedCodes_(types) {
  const m = {};
  (types || inspTypes_()).forEach(function (t) { if (isActive_(t.active) && t.pm_code && t.pm_sync !== 'off') m[t.pm_code] = t; });
  return m;
}

/** ตรวจเช็ค -> แผน: เขียน PM_Plan.actual (และสร้างแถวแผนถ้ายังไม่มี) ของเดือนในปีงบ fy  months = null ทุกเดือนที่ถึงแล้ว */
function syncInspPlan_(fy, months, who) {
  fy = Number(fy);
  const today = todayStr_(), from = inspKpiFrom_(), ov = inspOverrides_(), locks = lockSet_();
  const sh = db_().getSheetByName('PM_Plan'), v = sh.getDataRange().getValues(), head = v[0].map(String);
  const iF = head.indexOf('fy'), iM = head.indexOf('month'), iC = head.indexOf('code'), iP = head.indexOf('plan'), iA = head.indexOf('actual'), iN = head.indexOf('note'), iB = head.indexOf('updated_by'), iT = head.indexOf('updated_at');
  const plan = []; for (let i = 1; i < v.length; i++) plan.push({ fy: v[i][iF], month: v[i][iM], code: norm_(v[i][iC]), plan: v[i][iP] });
  const types = inspTypes_(plan), linked = inspLinkedCodes_(types);
  const itemOk = {}; read_('PM_Items').forEach(function (i) { itemOk[norm_(i.code)] = isActive_(i.active === null ? 1 : i.active); });
  const at = {}; for (let i = 1; i < v.length; i++) at[Number(v[i][iF]) + '|' + Number(v[i][iM]) + '|' + norm_(v[i][iC])] = i;
  const allLogs = read_('Insp_Log'), assets = inspAssets_(), now = new Date().toISOString();
  const res = { fy: fy, updated: 0, created: 0, unchanged: 0, overridden: 0, locked: 0, beforeStart: 0 }, logsA = [], adds = [];
  Object.keys(linked).forEach(function (code) {
    const t = linked[code]; if (!itemOk[code]) return;
    const as = assets.filter(function (a) { return a.type_id === t.type_id; });
    const ls = allLogs.filter(function (l) { return String(l.type_id) === t.type_id && inspFyOf(String(l.period).slice(0, 7) + '-01') === fy; })
      .map(function (l) { return { asset_id: String(l.asset_id), period: String(l.period), st: l.st }; });
    for (let mi = 1; mi <= 12; mi++) {
      if (months && months.indexOf(mi) < 0) continue;
      const k = fy + '|' + mi + '|' + code;
      const r = inspPlanMonth(t, as, ls, fy, mi, today, t.pm_sync); if (!r) continue;
      if (!inspKpiOwns_(fy, mi, from)) { res.beforeStart++; continue; }
      if (locks[fy + '|' + mi]) { res.locked++; continue; }
      if (ov['PM_Plan|' + k]) { res.overridden++; continue; }
      const i = at[k];
      if (i === undefined) {
        const line = head.map(function (h) { return { fy: fy, month: mi, code: code, plan: r.exp, actual: r.done, note: 'สร้างจากระบบตรวจเช็ค: ' + t.name, updated_by: INSP_SYS_WHO, updated_at: now }[h]; }).map(function (x) { return x === undefined ? '' : x; });
        adds.push(line); at[k] = -1; res.created++;
        logsA.push({ action: 'ตรวจเช็ค→แผน', sheet: 'PM_Plan', key: k, field: 'plan/actual', old: '', new: r.exp + '/' + r.done });
        continue;
      }
      if (i < 0) continue;
      const row = v[i], sysRow = /^สร้างจากระบบตรวจเช็ค/.test(String(row[iN] || ''));
      const newPlan = sysRow ? r.exp : row[iP];
      if (norm_(row[iA]) === norm_(r.done) && norm_(row[iP]) === norm_(newPlan)) { res.unchanged++; continue; }
      logsA.push({ action: 'ตรวจเช็ค→แผน', sheet: 'PM_Plan', key: k, field: 'actual', old: norm_(row[iA]), new: norm_(r.done) });
      row[iA] = r.done; row[iP] = newPlan; if (iB >= 0) row[iB] = INSP_SYS_WHO; if (iT >= 0) row[iT] = now;
      sh.getRange(i + 1, 1, 1, head.length).setValues([row]); res.updated++;
    }
  });
  if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, head.length).setValues(adds);
  if (logsA.length) audit_({ email: who || INSP_SYS_WHO, role: 'system' }, logsA);
  return res;
}
/** ตรวจเช็ค -> แผน -> KPI ตามลำดับ */
function syncLinks_(fy, months, who) {
  const out = { fy: Number(fy) };
  try { out.plan = syncInspPlan_(fy, months, who); } catch (e) { out.plan = { error: String(e.message || e) }; }
  out.kpi = syncInspKpiSafe_(fy, months, who);
  return out;
}
/** ผู้ดูแลระบบ: ซิงก์ทั้งปีงบ (ตรวจเช็ค -> แผน -> KPI) */
function syncInspToPlan(fy) {
  const u = requireAdmin_();
  fy = Number(fy); if (!fy) throw new Error('ปีงบไม่ถูกต้อง');
  return syncLinks_(fy, null, u.email);
}
/** ผู้ดูแลระบบ: ยกเลิกค่าที่แก้เอง ให้กลับไปใช้ค่าจากการคำนวณ  items = [{target, key}] หรือ 'all' */
function revertOverrides(items) {
  const u = requireAdmin_();
  const sh = db_().getSheetByName('Auto_Override'), v = sh.getDataRange().getValues();
  const want = {}; if (items !== 'all') (items || []).forEach(function (x) { want[x.target + '|' + norm_(x.key)] = 1; });
  const rows = [], touched = {};
  for (let i = v.length - 1; i >= 1; i--) {
    const k = v[i][0] + '|' + norm_(v[i][1]);
    if (items === 'all' || want[k]) { rows.push(i + 1); const p = String(v[i][1]).split('|'); (touched[p[0]] = touched[p[0]] || {})[Number(p[1])] = 1; }
  }
  rows.forEach(function (r) { sh.deleteRow(r); });
  if (rows.length) audit_(u, [{ action: 'คืนค่าอัตโนมัติ', sheet: 'Auto_Override', key: items === 'all' ? 'ทั้งหมด' : Object.keys(want).join(', ').slice(0, 200), field: '', old: rows.length + ' ช่อง', new: '' }]);
  const synced = Object.keys(touched).map(function (f) { return syncLinks_(Number(f), Object.keys(touched[f]).map(Number), u.email); });
  return { reverted: rows.length, synced: synced };
}
/** ลบบันทึกผลตรวจหลายรายการ (ผู้ดูแลระบบ) แล้วคำนวณแผน/KPI ของเดือนนั้นใหม่ */
function deleteInspLogs(ids) {
  const u = requireAdmin_();
  const want = {}; (ids || []).forEach(function (x) { want[String(x)] = 1; });
  if (!Object.keys(want).length) throw new Error('ไม่ได้เลือกรายการ');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  const touched = {};
  let n = 0;
  try {
    const sh = db_().getSheetByName('Insp_Log'), head = SHEETS.Insp_Log, v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
    const iD = head.indexOf('date'), iS = head.indexOf('st'), logsA = [];
    for (let i = v.length - 1; i >= 0; i--) {
      if (!want[String(v[i][0])]) continue;
      const d = String(v[i][iD]), f = inspFyOf(d); (touched[f] = touched[f] || {})[inspMiOf(+d.slice(5, 7))] = 1;
      sh.deleteRow(i + 2); n++;
      logsA.push({ action: 'ลบ', sheet: 'Insp_Log', key: String(v[i][0]), field: 'ผลตรวจ', old: String(v[i][iS]), new: '' });
    }
    audit_(u, logsA);
  } finally { lock.releaseLock(); }
  const synced = Object.keys(touched).map(function (f) { return syncLinks_(Number(f), Object.keys(touched[f]).map(Number), u.email); });
  return { removed: n, synced: synced };
}
/** อ่านบันทึกผลตรวจเต็ม (ผู้ดูแลระบบ ใช้เปิดแก้ไข) */
function getInspLog(logId) {
  requireAdmin_();
  const l = read_('Insp_Log').filter(function (x) { return String(x.log_id) === String(logId); })[0];
  if (!l) throw new Error('ไม่พบบันทึก ' + logId);
  return { log_id: String(l.log_id), type_id: String(l.type_id), asset_id: String(l.asset_id), date: String(l.date), period: String(l.period), inspector: l.inspector || '', results: l.results || '{}', note: l.note || '', source: l.source || '' };
}
/** ลบงานแก้ไข (ผู้ดูแลระบบ) */
function deleteInspActions(ids) {
  const u = requireAdmin_();
  const want = {}; (ids || []).forEach(function (x) { want[String(x)] = 1; });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = db_().getSheetByName('Insp_Actions'), head = SHEETS.Insp_Actions, v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
    const logsA = []; let n = 0;
    for (let i = v.length - 1; i >= 0; i--) {
      if (!want[String(v[i][0])]) continue;
      sh.deleteRow(i + 2); n++;
      logsA.push({ action: 'ลบ', sheet: 'Insp_Actions', key: String(v[i][0]), field: 'issue', old: String(v[i][head.indexOf('issue')]).slice(0, 120), new: '' });
    }
    audit_(u, logsA);
    return { removed: n };
  } finally { lock.releaseLock(); }
}

/** ผู้ดูแลระบบ: นำเข้าประวัติ (ทีละชุดไม่เกิน ~1500 แถว) payload: {assets, logs, actions} — นำเข้าซ้ำได้ ไม่สร้างข้อมูลซ้ำ */
function importInspection(payload) {
  const u = requireAdmin_();
  const logsIn = (payload && payload.logs) || [], actsIn = (payload && payload.actions) || [], assetsIn = (payload && payload.assets) || [];
  if (logsIn.length > 2500 || actsIn.length > 2500) throw new Error('ส่งข้อมูลครั้งละไม่เกิน 2,500 แถว');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = db_(), now = new Date().toISOString();
    const types = {}; inspTypes_().forEach(function (t) { types[t.type_id] = t; });
    const checks = inspChecks_(), byType = {};
    checks.forEach(function (c) { (byType[c.type_id] = byType[c.type_id] || []).push(c); });
    // ทรัพย์สิน
    const sa = ss.getSheetByName('Insp_Assets'), ha = SHEETS.Insp_Assets;
    const haveA = {}; inspAssets_().forEach(function (a) { haveA[a.asset_id] = true; });
    const addA = [];
    assetsIn.forEach(function (a) {
      if (!types[String(a.type_id)]) throw new Error('ไม่พบประเภท ' + a.type_id);
      if (haveA[String(a.asset_id)]) return;
      haveA[String(a.asset_id)] = true;
      addA.push(ha.map(function (h) { return a[h] === undefined || a[h] === null ? '' : a[h]; }));
    });
    if (addA.length) sa.getRange(sa.getLastRow() + 1, 1, addA.length, ha.length).setValues(addA);
    // บันทึกตรวจ
    const sl = ss.getSheetByName('Insp_Log'), hl = SHEETS.Insp_Log;
    const haveL = {};
    if (sl.getLastRow() > 1) sl.getRange(2, 1, sl.getLastRow() - 1, 1).getValues().forEach(function (r) { haveL[String(r[0])] = true; });
    const today = todayStr_(), addL = [];
    let skipped = 0, bad = 0;
    logsIn.forEach(function (l) {
      if (!types[String(l.type_id)] || !inspParse(l.date) || String(l.date) > today || !haveA[String(l.asset_id)]) { bad++; return; }
      const id = String(l.asset_id) + '|' + l.date;
      if (haveL[id]) { skipped++; return; }
      haveL[id] = true;
      let res = {}; try { res = JSON.parse(l.results || '{}'); } catch (e) { res = {}; }
      const ev = inspEvalEntry(byType[String(l.type_id)] || [], res);
      const rec = { log_id: id, type_id: String(l.type_id), asset_id: String(l.asset_id), date: l.date, period: l.period || inspPeriodOf(types[String(l.type_id)].freq, l.date), inspector: String(l.inspector || '').slice(0, 60),
        st: ev.st, ng_n: ev.ngN, wn_n: ev.wnN, results: JSON.stringify(ev.res), note: String(l.note || '').slice(0, 200), source: 'import', created_by: u.email, created_at: now };
      addL.push(hl.map(function (h) { return rec[h] === undefined || rec[h] === null ? '' : rec[h]; }));
    });
    if (addL.length) sl.getRange(sl.getLastRow() + 1, 1, addL.length, hl.length).setValues(addL);
    // งานแก้ไข
    const sc = ss.getSheetByName('Insp_Actions'), hc = SHEETS.Insp_Actions;
    const haveC = {};
    if (sc.getLastRow() > 1) sc.getRange(2, 1, sc.getLastRow() - 1, 1).getValues().forEach(function (r) { haveC[String(r[0])] = true; });
    const addC = [];
    actsIn.forEach(function (a) {
      if (!a.action_id || haveC[a.action_id] || !haveA[String(a.asset_id)]) return;
      haveC[a.action_id] = true;
      const as = inspAssets_().filter(function (x) { return x.asset_id === String(a.asset_id); })[0] || {};
      const t = types[String(a.type_id)] || {};
      const pm = inspPmItem_(t);
      const w = inspAssignees_(t, as, pm);
      a.assignee = a.assignee || w.name; a.assignee_email = a.assignee_email || w.email;
      a.issue = (as.name ? as.name + ': ' : '') + a.issue;
      a.updated_by = u.email; a.updated_at = now;
      addC.push(hc.map(function (h) { return a[h] === undefined || a[h] === null ? '' : a[h]; }));
    });
    if (addC.length) sc.getRange(sc.getLastRow() + 1, 1, addC.length, hc.length).setValues(addC);
    audit_(u, [{ action: 'นำเข้าผลตรวจ', sheet: 'Insp_Log', key: '', field: 'logs/assets/actions', old: '', new: addL.length + ' / ' + addA.length + ' / ' + addC.length + (skipped ? ' · ข้ามที่มีแล้ว ' + skipped : '') + (bad ? ' · ไม่ผ่าน ' + bad : '') }]);
    return { logs: addL.length, assets: addA.length, actions: addC.length, skipped: skipped, bad: bad };
  } finally { lock.releaseLock(); }
}

/** ผู้ดูแลระบบ: แก้ไขนิยามประเภท/รายการตรวจ/ทรัพย์สิน  p: {types, checks, assets} */
function saveInspMaster(p) {
  const u = requireAdmin_();
  p = p || {};
  const typeIds = {}; inspTypes_().forEach(function (t) { typeIds[t.type_id] = true; });
  const kpis = {}; read_('KPI_Master').forEach(function (k) { kpis[k.kpi_id] = true; });
  const pms = {}; read_('PM_Items').forEach(function (i) { pms[norm_(i.code)] = true; });
  const ym = function (s, label) { if (s && !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(s))) throw new Error(label + ' ต้องเป็นรูปแบบ ปปปป-ดด เช่น 2026-10'); };
  const out = {};
  // ประเภทใหม่ (type_id ว่าง) -> รหัส T## ถัดไป
  if (p.types && p.types.length) {
    let mx = 0; Object.keys(typeIds).forEach(function (id) { const m = /^T(\d+)$/.exec(id); if (m) mx = Math.max(mx, +m[1]); });
    p.types.forEach(function (t) {
      if (String(t.type_id || '').trim()) return;
      if (!String(t.name || '').trim()) throw new Error('ประเภทใหม่ต้องมีชื่อ');
      mx++; t.type_id = 'T' + (mx < 10 ? '0' + mx : mx); typeIds[t.type_id] = true; t._new = 1;
      if (t.active === undefined || t.active === '') t.active = 1;
      if (!t.order) t.order = mx;
      out.newTypeId = t.type_id;
    });
  }
  if (p.types && p.types.length) {
    out.types = upsert_('Insp_Types', ['type_id'], p.types, ['name', 'freq', 'pm_code', 'kpi_id', 'discipline', 'owner_name', 'owner_email', 'active', 'order', 'pm_sync', 'min_cov', 'note', 'standard'], {
      user: u,
      validate: function (r) {
        r.type_id = String(r.type_id);
        if (!typeIds[r.type_id]) throw new Error('ไม่พบประเภท ' + r.type_id);
        if (!String(r.name || '').trim()) throw new Error('ต้องระบุชื่อประเภท ' + r.type_id);
        if (['M', 'W', 'D'].indexOf(r.freq) < 0) throw new Error('ความถี่ต้องเป็น D (รายวัน) M (รายเดือน) หรือ W (รายสัปดาห์)');
        r.pm_code = r.pm_code ? norm_(r.pm_code) : '';
        if (r.pm_code && !pms[r.pm_code]) throw new Error('ไม่พบรหัสแผน ' + r.pm_code);
        if (r.kpi_id && !kpis[r.kpi_id]) throw new Error('ไม่พบตัวชี้วัด ' + r.kpi_id);
        if (INSP_SYNC_MODES.indexOf(r.pm_sync || 'off') < 0) throw new Error('pm_sync ต้องเป็น off / assets / logs / cycles');
        r.pm_sync = r.pm_sync || 'off';
        if (r.min_cov !== '' && r.min_cov !== null && r.min_cov !== undefined && !(Number(r.min_cov) > 0 && Number(r.min_cov) <= 1)) throw new Error('สัดส่วนที่นับว่าครบรอบต้องอยู่ระหว่าง 0.1–1');
        if (r.owner_email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r.owner_email)) throw new Error('อีเมลเจ้าของประเภทไม่ถูกต้อง');
        if (r.pm_sync !== 'off' && !r.pm_code) throw new Error('ต้องระบุรหัสแผนก่อนเปิดซิงก์เข้าแผน');
      },
      check: function () { return null; }
    });
  }
  // รายการตรวจใหม่ (check_id ว่าง) -> c## ถัดไปของประเภทนั้น ; รายการที่ลบ p.removeChecks = [{type_id, check_id}]
  if (p.checks && p.checks.length) {
    const used = {}; inspChecks_().forEach(function (c) { (used[c.type_id] = used[c.type_id] || {})[c.check_id] = 1; });
    p.checks.forEach(function (c) {
      c.type_id = String(c.type_id || out.newTypeId || '');
      if (!String(c.label || '').trim()) throw new Error('รายการตรวจต้องมีชื่อ');
      if (String(c.check_id || '').trim()) return;
      const u2 = used[c.type_id] = used[c.type_id] || {}; let n = 1; while (u2['c' + n]) n++;
      c.check_id = 'c' + n; u2[c.check_id] = 1;
    });
  }
  if (p.removeChecks && p.removeChecks.length) {
    const sh = db_().getSheetByName('Insp_Checks'), v = sh.getDataRange().getValues(), h = v[0].map(String), iT = h.indexOf('type_id'), iC = h.indexOf('check_id');
    const want = {}; p.removeChecks.forEach(function (x) { want[String(x.type_id) + '|' + String(x.check_id)] = 1; });
    const logsA = [];
    for (let i = v.length - 1; i >= 1; i--) { const k = String(v[i][iT]) + '|' + String(v[i][iC]); if (want[k]) { sh.deleteRow(i + 1); logsA.push({ action: 'ลบ', sheet: 'Insp_Checks', key: k, field: 'label', old: String(v[i][h.indexOf('label')]), new: '' }); } }
    audit_(u, logsA); out.removedChecks = logsA.length;
  }
  if (p.checks && p.checks.length) {
    out.checks = upsert_('Insp_Checks', ['type_id', 'check_id'], p.checks, ['label', 'kind', 'soft', 'min', 'max', 'tol', 'unit', 'crit', 'std', 'order'], {
      user: u,
      validate: function (r) {
        r.type_id = String(r.type_id); r.check_id = String(r.check_id);
        if (!typeIds[r.type_id]) throw new Error('ไม่พบประเภท ' + r.type_id);
        if (['s', 'n', 'm', 't'].indexOf(r.kind) < 0) throw new Error('ชนิดรายการต้องเป็น s / n / m / t');
        if (!String(r.label || '').trim()) throw new Error('รายการตรวจต้องมีชื่อ');
        ['min', 'max', 'tol'].forEach(function (f) { if (r[f] !== '' && r[f] !== null && r[f] !== undefined) checkNumber_(f, r[f]); });
      },
      check: function () { return null; }
    });
  }
  if (p.assets && p.assets.length) {
    const have = inspAssets_(), maxNo = {};
    have.forEach(function (a) { const m = /-(\d+)$/.exec(a.asset_id); if (m) maxNo[a.type_id] = Math.max(maxNo[a.type_id] || 0, +m[1]); });
    p.assets.forEach(function (a) {
      if (!a.asset_id) {
        if (!typeIds[String(a.type_id)]) throw new Error('ระบุประเภทของทรัพย์สินใหม่');
        const n = (maxNo[a.type_id] = (maxNo[a.type_id] || 0) + 1);
        a.asset_id = String(a.type_id) + '-' + (n < 10 ? '0' + n : n);
        if (!a.code) a.code = 'N' + a.asset_id;
      }
    });
    out.assets = upsert_('Insp_Assets', ['asset_id'], p.assets, ['type_id', 'name', 'location', 'code', 'group', 'active', 'since', 'until', 'responsible_name', 'responsible_email', 'note'], {
      user: u,
      validate: function (r) {
        r.asset_id = String(r.asset_id); r.type_id = String(r.type_id);
        if (!typeIds[r.type_id]) throw new Error('ไม่พบประเภท ' + r.type_id);
        if (!String(r.name || '').trim()) throw new Error('ต้องมีชื่อทรัพย์สิน (' + r.asset_id + ')');
        ym(r.since, 'เริ่มตรวจ'); ym(r.until, 'สิ้นสุด');
        if (r.responsible_email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r.responsible_email)) throw new Error('อีเมลผู้รับผิดชอบไม่ถูกต้อง (' + r.asset_id + ')');
      },
      check: function () { return null; }
    });
  }
  // ประเภท/จุดตรวจ/รหัสแผนเปลี่ยน -> คำนวณการไหลของปีงบปัจจุบันใหม่
  if ((p.types && p.types.length) || (p.assets && p.assets.length)) out.synced = syncLinks_(inspFyOf(todayStr_()), null, u.email);
  return out;
}

/** ผู้ดูแลระบบ: ลบบันทึกตรวจที่บันทึกผิด (มีประวัติ) */
function deleteInspLog(logId) {
  const r = deleteInspLogs([logId]);
  if (!r.removed) throw new Error('ไม่พบบันทึก ' + logId);
  return r;
}

/* ---------------- แจ้งเตือนทางอีเมล ---------------- */

function inspAlertText_(a, types, assets) {
  const t = types[a.type_id] || { name: a.type_id }, as = assets[a.asset_id];
  if (a.kind === 'missed' && a.nDays) return '[ขาดบันทึก] ' + t.name + ': ไม่ได้บันทึก/บันทึกไม่ครบ ' + a.nDays + ' วันใน ' + a.win + ' วันล่าสุด (ล่าสุด ' + inspPeriodLabel(a.period) + ')';
  if (a.kind === 'missed') return '[เลยรอบ] ' + t.name + ' ' + inspPeriodLabel(a.period) + ': ตรวจแล้ว ' + a.done + '/' + a.exp + ' (เลยกำหนด ' + a.days + ' วัน)';
  if (a.kind === 'due') return '[ใกล้ครบรอบ] ' + t.name + ' ' + inspPeriodLabel(a.period) + ': ตรวจแล้ว ' + a.done + '/' + a.exp + ' เหลือ ' + a.days + ' วัน';
  if (a.kind === 'action_overdue') return '[งานแก้ไขเกินกำหนด] ' + t.name + (as ? ' · ' + as.name : '') + ' เกิน ' + a.days + ' วัน (' + a.action_id + ')';
  if (a.kind === 'action_due') return '[งานแก้ไขใกล้ครบกำหนด] ' + t.name + (as ? ' · ' + as.name : '') + ' อีก ' + a.days + ' วัน (' + a.action_id + ')';
  if (a.kind === 'repeat') return '[พบซ้ำ ' + a.count + ' รอบ] ' + t.name + (as ? ' · ' + as.name : '') + ' ค้างมา ' + a.days + ' วัน';
  return '';
}

/** ตัวตั้งเวลารายวันเรียกฟังก์ชันนี้ ส่งอีเมลสรุปได้วันละครั้ง และต้องเปิด insp_mail */
function sendInspAlertsJob() {
  const today = todayStr_();
  // ทุกวัน: คำนวณ KPI จากผลตรวจเช็คของปีงบปัจจุบัน (รอบที่เพิ่งสิ้นสุดจะนับเป็นขาดการตรวจ) และเดือนสุดท้ายของปีงบก่อนในเดือนแรกของปี
  const fyNow = inspFyOf(today);
  syncLinks_(fyNow, null, 'ระบบตรวจเช็ค (รายวัน)');
  if (inspMiOf(+today.slice(5, 7)) === 1) syncLinks_(fyNow - 1, [12], 'ระบบตรวจเช็ค (รายวัน)');
  const sh = db_().getSheetByName('Settings'), v = sh.getDataRange().getValues();
  let row = 0, last = '';
  for (let i = 1; i < v.length; i++) { if (v[i][0] === 'insp_mail_last') { row = i + 1; last = String(v[i][1]); } }
  if (last === today) return { sent: 0, skipped: 'ส่งแล้ววันนี้' };
  const r = inspSendMail_();
  if (r.sent || !r.skipped) { if (row) sh.getRange(row, 2).setValue(today); else sh.appendRow(['insp_mail_last', today]); }
  return r;
}

/** ผู้ดูแลระบบกดส่งสรุปแจ้งเตือนทันที (ทดสอบ) */
function sendInspAlertsNow() {
  requireAdmin_();
  return inspSendMail_();
}

function inspSendMail_() {
  if (!inspSettings_().mail) return { sent: 0, skipped: 'ปิดการส่งอีเมล' };
  const b = inspBundle_();
  const alerts = inspAlerts(b.types, b.assets, b.logs, b.actions, b.today).filter(function (a) { return a.level !== 'info'; });
  if (!alerts.length) return { sent: 0, alerts: 0 };
  const types = {}, assets = {}, items = {};
  b.types.forEach(function (t) { types[t.type_id] = t; });
  b.assets.forEach(function (a) { assets[a.asset_id] = a; });
  read_('PM_Items').forEach(function (i) { items[norm_(i.code)] = i; });
  const actions = {}; b.actions.forEach(function (a) { actions[a.action_id] = a; });
  const valid = function (e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); };
  const to = {};
  const push = function (em, line) { em = String(em || '').trim().toLowerCase(); if (!valid(em)) return; (to[em] = to[em] || []).push(line); };
  const admins = read_('Users').filter(function (x) { return x.role === 'admin' && isActive_(x.active); }).map(function (x) { return x.email; });
  admins.push(Session.getEffectiveUser().getEmail());
  alerts.forEach(function (a) {
    const line = inspAlertText_(a, types, assets);
    if (!line) return;
    const t = types[a.type_id] || {}, pm = items[norm_(t.pm_code)] || {}, as = assets[a.asset_id] || {}, ac = actions[a.action_id] || {};
    [t.owner_email, pm.responsible_email, as.responsible_email].concat(String(ac.assignee_email || '').split(/[,;]+/), inspRespEmails_(a.type_id, as.group)).forEach(function (em) { push(em, line); });
    admins.forEach(function (em) { push(em, line); });
  });
  let sent = 0;
  const link = appUrl_();
  Object.keys(to).forEach(function (em) {
    const lines = to[em].filter(function (x, i, arr) { return arr.indexOf(x) === i; });
    const body = 'สรุปแจ้งเตือนงานตรวจเช็คบำรุงรักษา ณ ' + b.today + '\n\n' + lines.slice(0, 60).join('\n') + (lines.length > 60 ? '\n… และอีก ' + (lines.length - 60) + ' รายการ' : '') + (link ? '\n\nเปิดระบบ: ' + link : '');
    if (typeof MailApp !== 'undefined') { MailApp.sendEmail({ to: em, subject: '[แจ้งเตือน] งานตรวจเช็คบำรุงรักษา ' + lines.length + ' รายการ', body: body }); sent++; }
  });
  return { sent: sent, alerts: alerts.length };
}

/** รันครั้งเดียวเพื่อตั้งเวลาส่งแจ้งเตือนทุกวันเวลา 07:00 (ต้องกดอนุญาตสิทธิ์ส่งอีเมล) */
function installInspTrigger() {
  requireAdmin_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendInspAlertsJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendInspAlertsJob').timeBased().everyDays(1).atHour(7).create();
  return 'ตั้งเวลาส่งแจ้งเตือนทุกวัน 07:00 แล้ว';
}
function removeInspTrigger() {
  requireAdmin_();
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendInspAlertsJob') { ScriptApp.deleteTrigger(t); n++; } });
  return 'ลบตัวตั้งเวลา ' + n + ' รายการ';
}

/* ====== ผลตรวจเช็ค -> ตัวชี้วัดของฝ่าย (KPI_Data) ====== */
/** เดือนแรกที่ให้ระบบคำนวณ KPI จากผลตรวจเช็ค ('YYYY-MM') ค่าเริ่มต้น = ต้นปีงบปัจจุบัน (เดือนก่อนหน้านั้นคงค่าที่รายงานไว้) */
function inspKpiFrom_() {
  let v = '';
  read_('Settings').forEach(function (r) { if (r.key === 'insp_kpi_from') v = String(r.value || ''); });
  if (/^\d{4}-\d{2}$/.test(v)) return v;
  const fy = inspFyOf(todayStr_());
  return inspYmOf(fy, 1).key;
}
function inspKpiIds_() {
  return read_('KPI_Master').filter(function (k) { return k.source === 'insp' && isActive_(k.active); }).map(function (k) { return String(k.kpi_id); });
}
/** เดือน (fy, mi) อยู่ในช่วงที่ระบบคำนวณ KPI จากผลตรวจเช็คหรือไม่ */
function inspKpiOwns_(fy, mi, from) { return inspYmOf(Number(fy), Number(mi)).key >= (from || inspKpiFrom_()); }

/**
 * คำนวณและเขียนผล KPI ที่ตั้งแหล่งข้อมูลเป็นผลตรวจเช็ค ลง KPI_Data  (ไม่ล็อกเอง ผู้เรียกจัดการ)
 * months = [1..12] หรือ null = ทุกเดือนที่ถึงแล้ว · ข้ามเดือนที่ล็อก และเดือนก่อน insp_kpi_from
 */
function syncInspKpi_(fy, months, who) {
  fy = Number(fy);
  const ids = inspKpiIds_();
  if (!ids.length) return { fy: fy, skipped: 'ไม่มีตัวชี้วัดที่ใช้ผลตรวจเช็ค' };
  const today = todayStr_(), from = inspKpiFrom_(), ov = inspOverrides_();
  const logs = read_('Insp_Log').filter(function (l) { return inspFyOf(String(l.period).slice(0, 7) + '-01') === fy; })
    .map(function (l) { return { asset_id: String(l.asset_id), type_id: String(l.type_id), period: String(l.period), date: String(l.date), st: l.st }; });
  const ovPlan = {}; Object.keys(ov).forEach(function (k) { if (k.indexOf('PM_Plan|') === 0) ovPlan[k.slice(8)] = 1; });
  const D = { ovPlan: ovPlan, types: inspTypes_(), assets: inspAssets_(), logs: logs, items: read_('PM_Items').map(function (i) { i.code = norm_(i.code); return i; }), plan: read_('PM_Plan').map(function (r) { r.code = norm_(r.code); return r; }) };
  const locks = lockSet_();
  const sh = db_().getSheetByName('KPI_Data'), v = sh.getDataRange().getValues(), head = v[0].map(String);
  const iF = head.indexOf('fy'), iM = head.indexOf('month'), iK = head.indexOf('kpi_id'), iN = head.indexOf('num'), iD = head.indexOf('den'), iNo = head.indexOf('note'), iB = head.indexOf('updated_by'), iA = head.indexOf('updated_at');
  const at = {};
  for (let i = 1; i < v.length; i++) at[Number(v[i][iF]) + '|' + Number(v[i][iM]) + '|' + String(v[i][iK])] = i;
  const now = new Date().toISOString(), logsA = [], adds = [];
  const res = { fy: fy, written: 0, unchanged: 0, locked: 0, beforeStart: 0, from: from, kpis: ids };
  ids.forEach(function (kid) {
    const yr = inspKpiYear(kid, fy, D, today);
    for (let mi = 1; mi <= 12; mi++) {
      if (months && months.indexOf(mi) < 0) continue;
      const c = yr[mi - 1]; if (!c) continue;
      if (!inspKpiOwns_(fy, mi, from)) { res.beforeStart++; continue; }
      if (locks[fy + '|' + mi]) { res.locked++; continue; }
      if (ov['KPI_Data|' + fy + '|' + mi + '|' + kid]) { res.overridden = (res.overridden || 0) + 1; continue; }
      const note = inspKpiNote(c), k = fy + '|' + mi + '|' + kid, i = at[k];
      if (i !== undefined) {
        const r = v[i];
        if (norm_(r[iN]) === norm_(c.num) && norm_(r[iD]) === norm_(c.den) && String(r[iNo]) === note) { res.unchanged++; continue; }
        logsA.push({ action: 'คำนวณจากผลตรวจเช็ค', sheet: 'KPI_Data', key: k, field: 'num/den', old: norm_(r[iN]) + '/' + norm_(r[iD]), new: c.num + '/' + c.den });
        r[iN] = c.num; r[iD] = c.den; r[iNo] = note; if (iB >= 0) r[iB] = 'ระบบตรวจเช็ค'; if (iA >= 0) r[iA] = now;
        sh.getRange(i + 1, 1, 1, head.length).setValues([r]);
      } else {
        const line = head.map(function (h) { return { fy: fy, month: mi, kpi_id: kid, num: c.num, den: c.den, note: note, updated_by: 'ระบบตรวจเช็ค', updated_at: now }[h]; }).map(function (x) { return x === undefined ? '' : x; });
        adds.push(line);
        logsA.push({ action: 'คำนวณจากผลตรวจเช็ค', sheet: 'KPI_Data', key: k, field: 'num/den', old: '', new: c.num + '/' + c.den });
      }
      res.written++;
    }
  });
  if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, head.length).setValues(adds);
  if (logsA.length) audit_({ email: who || 'ระบบตรวจเช็ค', role: 'system' }, logsA);
  return res;
}
/** เรียกหลังบันทึกข้อมูลที่กระทบ KPI โดยไม่ทำให้การบันทึกหลักล้มเหลว */
function syncInspKpiSafe_(fy, months, who) {
  try { return syncInspKpi_(fy, months, who); } catch (e) { return { error: String(e.message || e) }; }
}
/** ผู้ดูแลระบบสั่งคำนวณ KPI จากผลตรวจเช็คใหม่ (fy ว่าง = ทุกปีงบตั้งแต่ insp_kpi_from ถึงปัจจุบัน) */
function syncInspKpi(fy) {
  const u = requireAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = inspFyOf(todayStr_()), from = inspKpiFrom_(), f0 = inspFyOf(from + '-01');
    const fys = fy ? [Number(fy)] : [];
    if (!fy) for (let y = f0; y <= now; y++) fys.push(y);
    return fys.map(function (y) { const r = syncLinks_(y, null, u.email); const k = r.kpi || {}; k.plan = r.plan; return k; });
  } finally { lock.releaseLock(); }
}


/* ===== ผู้รับผิดชอบหลายคนต่อระบบ/อาคาร + รหัสเจ้าหน้าที่ 7 หลัก (รุ่น 9) =====
 * Insp_Resp: 1 แถว = เจ้าหน้าที่ (staff_code) × ระบบ (type_id หรือ '*') × อาคาร (group หรือ '*')
 * เจ้าหน้าที่ 1 คนมีได้หลายแถว  ระบบ/อาคารหนึ่งมีได้หลายคน */
function inspResp_() {
  return read_('Insp_Resp').map(function (r) {
    return { staff_code: String(r.staff_code || '').trim(), type_id: String(r.type_id || '*').trim() || '*', group: String(r.group || '*').trim() || '*', note: r.note || '' };
  }).filter(function (r) { return r.staff_code; });
}
function inspStaffName_(s) {
  const f = String(s.first_name || '').trim(), l = String(s.last_name || '').trim();
  return (f || l) ? (f + ' ' + l).trim() : String(s.name || '').trim();
}
/** รหัสเจ้าหน้าที่ของผู้ใช้ (จากชีต Users และ Staff ที่อีเมลตรงกัน) -> แถวมอบหมาย */
function inspRespOf_(email, userRow) {
  email = String(email || '').toLowerCase();
  const codes = {};
  if (userRow && String(userRow.staff_code || '').trim()) codes[String(userRow.staff_code).trim()] = 1;
  read_('Staff').forEach(function (s) {
    if (email && String(s.email || '').toLowerCase() === email && String(s.staff_code || '').trim() && isActive_(s.active)) codes[String(s.staff_code).trim()] = 1;
  });
  if (!Object.keys(codes).length) return [];
  return inspResp_().filter(function (r) { return codes[r.staff_code]; }).map(function (r) { return { type_id: r.type_id, group: r.group }; });
}
/** เจ้าหน้าที่ที่ใช้งานอยู่และรับผิดชอบ (ระบบ type_id, อาคาร group) */
function inspRespStaff_(typeId, group) {
  const by = {}; read_('Staff').forEach(function (s) { const c = String(s.staff_code || '').trim(); if (c && isActive_(s.active)) by[c] = s; });
  const seen = {}, out = [];
  inspResp_().forEach(function (r) {
    if ((r.type_id === '*' || r.type_id === String(typeId)) && (r.group === '*' || r.group === String(group || '').trim()) && by[r.staff_code] && !seen[r.staff_code]) {
      seen[r.staff_code] = 1; out.push(by[r.staff_code]);
    }
  });
  return out;
}
function inspRespEmails_(typeId, group) {
  return inspRespStaff_(typeId, group).map(function (s) { return String(s.email || '').trim(); }).filter(String);
}
/** ผู้รับผิดชอบงานแก้ไข: ผู้รับผิดชอบจุดตรวจ > เจ้าหน้าที่ตามระบบ×อาคาร (หลายคน) > เจ้าของประเภท > ผู้รับผิดชอบรายการแผน */
function inspAssignees_(type, asset, pm) {
  if (asset && (asset.responsible_name || asset.responsible_email)) return { name: asset.responsible_name || '', email: asset.responsible_email || '' };
  const st = inspRespStaff_(type && type.type_id, asset && asset.group);
  if (st.length) return { name: st.map(inspStaffName_).join(', ').slice(0, 200), email: st.map(function (s) { return String(s.email || '').trim(); }).filter(String).join(',').slice(0, 200) };
  return { name: (type && type.owner_name) || (pm && pm.responsible_name) || '', email: (type && type.owner_email) || (pm && pm.responsible_email) || '' };
}

/**
 * ผู้ดูแลระบบกำหนดเจ้าหน้าที่ + สิทธิ์ตรวจเช็ค
 * p: {staff_code:'7 หลัก', first_name, last_name, email?, duty?, active, grant (อนุญาตเข้าตรวจด้วยรหัส+ชื่อ ค่าเริ่มต้น 1) — อีเมลเป็นทางเลือก ใช้ส่งแจ้งเตือน/รายงาน,
 *     orig_name? (แถวเดิมที่ยังไม่มีรหัส), assignments:[{type_id:'*'|id, group:'*'|ชื่ออาคาร}]}
 */
function saveInspStaff(p) {
  const u = requireAdmin_();
  p = p || {};
  const code = String(p.staff_code || '').trim();
  if (!/^\d{7}$/.test(code)) throw new Error('รหัสเจ้าหน้าที่ต้องเป็นตัวเลข 7 หลัก');
  const first = String(p.first_name || '').trim().slice(0, 60), last = String(p.last_name || '').trim().slice(0, 60);
  if (!first || !last) throw new Error('กรุณาระบุชื่อและนามสกุล');
  const email = String(p.email || '').trim().toLowerCase();
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  const full = first + ' ' + last;
  const typeIds = {}; inspTypes_().forEach(function (t) { typeIds[t.type_id] = 1; });
  const seen = {}, asg = [];
  (p.assignments || []).forEach(function (a) {
    const t = String(a.type_id || '*').trim() || '*', g = String(a.group || '*').trim() || '*';
    if (t !== '*' && !typeIds[t]) throw new Error('ไม่พบระบบ ' + t);
    const k = t + '|' + g; if (seen[k]) return; seen[k] = 1; asg.push([code, t, g, '']);
  });
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  const logs = [];
  try {
    const ss = db_();
    // --- Staff
    const sh = ss.getSheetByName('Staff'), v = sh.getDataRange().getValues(), h = v[0].map(String);
    const iCL = h.indexOf('can_login'), ic = h.indexOf('staff_code'), iN = h.indexOf('name'), iE = h.indexOf('email'), iD = h.indexOf('duty'), iA = h.indexOf('active'), iF = h.indexOf('first_name'), iL = h.indexOf('last_name');
    if ([ic, iN, iF, iL].some(function (i) { return i < 0; })) throw new Error('ชีต Staff ขาดคอลัมน์ ให้รัน upgradeDb');
    let at = -1;
    for (let i = 1; i < v.length; i++) if (String(v[i][ic]).trim() === code) at = i;
    if (at < 0 && p.orig_name) for (let i = 1; i < v.length; i++) if (!String(v[i][ic]).trim() && String(v[i][iN]).trim() === String(p.orig_name).trim()) at = i;
    for (let i = 1; i < v.length; i++) if (i !== at && String(v[i][iN]).trim() === full) throw new Error('มีชื่อ "' + full + '" ในรายชื่อเจ้าหน้าที่แล้ว (คนละรหัส)');
    const act = p.active === false || p.active === 0 || p.active === '0' ? 0 : 1;
    const row = h.map(function (_, ci) { return at >= 0 ? v[at][ci] : ''; });
    row[ic] = code; row[iN] = full; row[iF] = first; row[iL] = last; if (iE >= 0) row[iE] = email; if (iD >= 0 && p.duty !== undefined) row[iD] = String(p.duty || '').trim().slice(0, 80); if (iA >= 0) row[iA] = act; if (iCL >= 0) row[iCL] = (p.grant === false || p.grant === 0 || p.grant === '0') ? 0 : 1;
    if (at >= 0) sh.getRange(at + 1, 1, 1, h.length).setValues([row]); else sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setValues([row]);
    logs.push({ action: at >= 0 ? 'แก้ไขเจ้าหน้าที่' : 'เพิ่มเจ้าหน้าที่', sheet: 'Staff', key: code, field: 'ชื่อ/อีเมล', old: at >= 0 ? v[at][iN] : '', new: full + (email ? ' <' + email + '>' : '') });
    // --- Insp_Resp: แทนที่แถวของรหัสนี้ทั้งหมด
    const rs = ss.getSheetByName('Insp_Resp'), rv = rs.getDataRange().getValues(), rh = rv[0].map(String), rc = rh.indexOf('staff_code');
    const keep = []; let oldN = 0;
    for (let i = 1; i < rv.length; i++) { if (String(rv[i][rc]).trim() === code) oldN++; else keep.push(rv[i]); }
    const all = keep.concat(asg.map(function (a) { return rh.map(function (hh) { return { staff_code: a[0], type_id: a[1], group: a[2], note: a[3] }[hh] || ''; }); }));
    if (rs.getLastRow() > 1) rs.getRange(2, 1, rs.getLastRow() - 1, rh.length).clearContent();
    if (all.length) rs.getRange(2, 1, all.length, rh.length).setValues(all);
    logs.push({ action: 'กำหนดผู้รับผิดชอบระบบ×อาคาร', sheet: 'Insp_Resp', key: code, field: 'รายการ', old: String(oldN), new: String(asg.length) });
  } finally { lock.releaseLock(); }
  audit_(u, logs);
  return { staff_code: code, name: full, assignments: asg.length };
}


/* ===== เข้าระบบตรวจเช็คด้วย รหัสเจ้าหน้าที่ 7 หลัก + ชื่อ-นามสกุล (ไม่ใช้อีเมล) ===== */
var STAFF_CTX_ = null;
function inspNameKey_(x) { return String(x || '').replace(/[\s\u200b]+/g, ' ').trim().toLowerCase(); }
function staffUser_(s) {
  const code = String(s.staff_code || '').trim();
  const rs = inspResp_().filter(function (r) { return r.staff_code === code; }).map(function (r) { return { type_id: r.type_id, group: r.group }; });
  const full = inspStaffName_(s);
  // ไม่มีการมอบหมาย = ไม่มีสิทธิ์ตรวจอะไรเลย (ไม่ให้ตกไปเป็น "ทุกอาคาร")
  if (!rs.length) rs.push({ type_id: '-', group: '-' });
  const uniq = function (k) { const m = {}; rs.forEach(function (r) { m[r[k]] = 1; }); return m['*'] ? [] : Object.keys(m); };
  return { email: String(s.email || '').trim().toLowerCase(), name: full + ' (' + code + ')', staffName: full, role: 'inspector', units: [], inspTypes: uniq('type_id'), inspGroups: uniq('group'), inspResp: rs, staffCode: code, viaCode: true };
}
function inspLoginFail_(code, bump) {
  const cache = CacheService.getScriptCache(), k = 'lf_' + code, n = Number(cache.get(k) || 0);
  if (bump) cache.put(k, String(n + 1), 900);
  return n + (bump ? 1 : 0);
}
/** เข้าสู่ระบบ: ตรวจรหัส 7 หลัก + ชื่อ + นามสกุล คืน token (อายุ 6 ชั่วโมง) */
function loginStaff(code, first, last) {
  code = String(code || '').trim();
  if (!/^\d{7}$/.test(code)) throw new Error('รหัสเจ้าหน้าที่ต้องเป็นตัวเลข 7 หลัก');
  if (inspLoginFail_(code, false) >= 5) throw new Error('ใส่ข้อมูลผิดหลายครั้ง กรุณารอ 15 นาทีแล้วลองใหม่ หรือติดต่อผู้ดูแลระบบ');
  const s = read_('Staff').filter(function (x) { return String(x.staff_code || '').trim() === code; })[0];
  const ok = s && isActive_(s.active) && isActive_(s.can_login === null || s.can_login === undefined ? 1 : s.can_login) &&
    inspNameKey_(s.first_name) === inspNameKey_(first) && inspNameKey_(s.last_name) === inspNameKey_(last);
  if (!ok) { inspLoginFail_(code, true); throw new Error('รหัสเจ้าหน้าที่หรือชื่อ-นามสกุลไม่ถูกต้อง (หรือยังไม่ได้รับสิทธิ์) กรุณาติดต่อผู้ดูแลระบบ'); }
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('stf_' + token, code, 21600);
  audit_({ email: 'รหัส ' + code, name: inspStaffName_(s), role: 'inspector' }, [{ action: 'เข้าสู่ระบบ', sheet: 'Staff', key: code, field: 'login', old: '', new: inspStaffName_(s) }]);
  return { token: token, name: inspStaffName_(s), staff_code: code };
}
function logoutStaff(token) { try { CacheService.getScriptCache().remove('stf_' + String(token || '')); } catch (e) {} return true; }
/** ตัวกลางเรียกฟังก์ชันด้วย token ของเจ้าหน้าที่ (เฉพาะฟังก์ชันของผู้ตรวจเช็ค) */
function api(token, fn, args) {
  const ALLOW = { getAll: 1, saveInspection: 1, saveActions: 1 };
  if (!ALLOW[fn]) throw new Error('ฟังก์ชันนี้ไม่เปิดให้ใช้ด้วยรหัสเจ้าหน้าที่');
  const cache = CacheService.getScriptCache(), code = cache.get('stf_' + String(token || ''));
  if (!code) throw new Error('SESSION_EXPIRED: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const s = read_('Staff').filter(function (x) { return String(x.staff_code || '').trim() === code; })[0];
  if (!s || !isActive_(s.active) || !isActive_(s.can_login === null || s.can_login === undefined ? 1 : s.can_login)) throw new Error('SESSION_EXPIRED: สิทธิ์ถูกปิดใช้งาน');
  cache.put('stf_' + String(token), code, 21600);
  STAFF_CTX_ = staffUser_(s);
  try { return globalThis[fn].apply(null, args || []); } finally { STAFF_CTX_ = null; }
}

/* ===== รายงานผลการบำรุงรักษา (อีเมล) รายสัปดาห์ / เมื่อผู้ดูแลต้องการ ===== */
function inspSettingStr_(key) { let v = ''; read_('Settings').forEach(function (r) { if (r.key === key) v = String(r.value || ''); }); return v; }
function saveInspReportTo(text) {
  const u = requireAdmin_();
  const list = String(text || '').split(/[,;\s]+/).map(function (x) { return x.trim().toLowerCase(); }).filter(String);
  list.forEach(function (e) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('อีเมลไม่ถูกต้อง: ' + e); });
  const val = list.filter(function (e, i, a) { return a.indexOf(e) === i; }).join(',');
  const sh = db_().getSheetByName('Settings'), v = sh.getDataRange().getValues();
  let row = 0, old = ''; for (let i = 1; i < v.length; i++) if (v[i][0] === 'insp_report_to') { row = i + 1; old = String(v[i][1]); }
  if (row) sh.getRange(row, 2).setValue(val); else sh.appendRow(['insp_report_to', val]);
  if (old !== val) audit_(u, [{ action: 'ตั้งค่า', sheet: 'Settings', key: 'insp_report_to', field: 'value', old: old, new: val }]);
  return { to: val };
}
function inspScopeFilter_(resp, a) { return !resp || resp.some(function (r) { return (r.type_id === '*' || r.type_id === a.type_id) && (r.group === '*' || r.group === String(a.group || '').trim()); }); }
/** ข้อมูลรายงานของขอบเขต resp (null = ทั้งหมด) ย้อนหลัง 7 วันถึง asOf */
function inspReportData_(resp, asOf) {
  const types = inspTypes_(), tBy = {}; types.forEach(function (t) { tBy[t.type_id] = t; });
  const assets = inspAssets_().filter(function (a) { return inspScopeFilter_(resp, a); }), aBy = {}; assets.forEach(function (a) { aBy[a.asset_id] = a; });
  const from = inspAdd(asOf, -6), cut = inspAdd(asOf, -120);
  const logs = read_('Insp_Log').filter(function (l) { return String(l.date) >= cut && aBy[String(l.asset_id)]; }).map(function (l) { return { log_id: l.log_id, type_id: String(l.type_id), asset_id: String(l.asset_id), date: String(l.date), period: String(l.period), st: l.st, ng_n: Number(l.ng_n) || 0, wn_n: Number(l.wn_n) || 0 }; });
  const acts = read_('Insp_Actions').filter(function (a) { return aBy[String(a.asset_id)]; });
  const tIn = {}; assets.forEach(function (a) { tIn[a.type_id] = 1; });
  const ty = types.filter(function (t) { return tIn[t.type_id] && isActive_(t.active); });
  const alerts = inspAlerts(ty, assets.filter(function (a) { return isActive_(a.active); }), logs, acts, asOf);
  const per = ty.map(function (t) {
    const ls = logs.filter(function (l) { return l.type_id === t.type_id && l.date >= from; });
    const ac = acts.filter(function (a) { return a.type_id === t.type_id && (a.status === 'open' || a.status === 'in_progress'); });
    return { name: t.name, assets: assets.filter(function (a) { return a.type_id === t.type_id && isActive_(a.active); }).length, week: ls.length, ng: ls.filter(function (l) { return l.st === 'ng'; }).length,
      open: ac.length, over: ac.filter(function (a) { return a.due_date && String(a.due_date) < asOf; }).length, missed: alerts.filter(function (x) { return x.kind === 'missed' && x.type_id === t.type_id; }).length };
  });
  const closed = acts.filter(function (a) { return a.status === 'done' && String(a.done_date || '') >= from; }).length;
  const lines = alerts.filter(function (a) { return a.level === 'bad' || a.kind === 'repeat'; }).slice(0, 25)
    .map(function (a) { return inspAlertText_(a, tBy, aBy); }).filter(String);
  const sum = function (k) { return per.reduce(function (s, x) { return s + x[k]; }, 0); };
  return { asOf: asOf, from: from, per: per, lines: lines, closed: closed, tot: { week: sum('week'), ng: sum('ng'), open: sum('open'), over: sum('over'), missed: sum('missed'), assets: sum('assets') } };
}
function inspReportHtml_(title, who, d) {
  const esc = function (x) { return String(x === null || x === undefined ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const th = 'style="text-align:left;padding:8px 10px;background:#e3eefd;color:#0f2347;font-size:13px"', td = 'style="padding:7px 10px;border-bottom:1px solid #dfe7f4;font-size:14px"';
  const tile = function (n, l, c) { return '<td style="padding:10px 14px;background:#f5f8fe;border-radius:8px;text-align:center"><div style="font-size:24px;font-weight:700;color:' + c + '">' + n + '</div><div style="font-size:12px;color:#46587a">' + l + '</div></td><td width="8"></td>'; };
  const T = d.tot, link = appUrl_();
  let h = '<div style="font-family:Tahoma,Arial,sans-serif;max-width:760px;color:#0f2347"><h2 style="margin:0 0 4px;color:#1d4ed8">' + esc(title) + '</h2>' +
    '<div style="color:#46587a;font-size:13px;margin-bottom:12px">' + esc(who) + ' · สัปดาห์ ' + esc(d.from) + ' ถึง ' + esc(d.asOf) + '</div>' +
    '<table cellspacing="0" cellpadding="0"><tr>' + tile(T.week, 'ตรวจในสัปดาห์ (ครั้ง)', '#1d4ed8') + tile(T.ng, 'พบไม่ปกติ (ครั้ง)', T.ng ? '#d23c3c' : '#15803d') + tile(T.open, 'งานแก้ไขค้าง', T.open ? '#b45309' : '#15803d') + tile(T.over, 'เกินกำหนด', T.over ? '#d23c3c' : '#15803d') + tile(d.closed, 'ปิดงานในสัปดาห์', '#15803d') + '</tr></table>' +
    '<h3 style="margin:16px 0 6px;font-size:15px">สรุปรายระบบ</h3><table cellspacing="0" width="100%"><tr><th ' + th + '>ระบบ</th><th ' + th + '>จุดตรวจ</th><th ' + th + '>ตรวจสัปดาห์นี้</th><th ' + th + '>ไม่ปกติ</th><th ' + th + '>งานค้าง</th><th ' + th + '>เกินกำหนด</th><th ' + th + '>เลยรอบ</th></tr>';
  d.per.forEach(function (x) { h += '<tr><td ' + td + '>' + esc(x.name) + '</td><td ' + td + '>' + x.assets + '</td><td ' + td + '>' + x.week + '</td><td ' + td + '>' + x.ng + '</td><td ' + td + '>' + x.open + '</td><td ' + td + (x.over ? ' bgcolor="#fdeeee"' : '') + '>' + x.over + '</td><td ' + td + (x.missed ? ' bgcolor="#fdeeee"' : '') + '>' + x.missed + '</td></tr>'; });
  if (!d.per.length) h += '<tr><td ' + td + ' colspan="7">ไม่มีระบบในขอบเขตนี้</td></tr>';
  h += '</table><h3 style="margin:16px 0 6px;font-size:15px">รายการที่ต้องติดตาม</h3>';
  h += d.lines.length ? '<ul style="margin:0;padding-left:20px;font-size:14px;line-height:1.7">' + d.lines.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>' : '<div style="font-size:14px;color:#15803d">ไม่มีรายการเลยรอบหรือเกินกำหนด</div>';
  if (link) h += '<p style="margin-top:16px"><a href="' + esc(link) + '" style="background:#2563eb;color:#fff;padding:9px 16px;border-radius:8px;text-decoration:none">เปิดระบบ</a></p>';
  return h + '<div style="margin-top:18px;font-size:12px;color:#7a8aa6">อีเมลนี้สร้างอัตโนมัติจากระบบตรวจเช็คบำรุงรักษา</div></div>';
}
function inspReportText_(d) {
  return 'ตรวจในสัปดาห์ ' + d.tot.week + ' ครั้ง · พบไม่ปกติ ' + d.tot.ng + ' · งานค้าง ' + d.tot.open + ' (เกินกำหนด ' + d.tot.over + ') · ปิดงาน ' + d.closed + '\n' + d.lines.join('\n');
}
/** ดูตัวอย่างรายงาน (ผู้ดูแลระบบ) code='' = ภาพรวมทั้งหมด หรือรหัสเจ้าหน้าที่ 7 หลัก */
function previewInspReport(code) {
  requireAdmin_();
  code = String(code || '').trim();
  const asOf = todayStr_();
  if (!code) return { html: inspReportHtml_('รายงานผลการบำรุงรักษารายสัปดาห์', 'ภาพรวมทุกอาคาร ทุกระบบ', inspReportData_(null, asOf)) };
  const s = read_('Staff').filter(function (x) { return String(x.staff_code || '').trim() === code; })[0];
  if (!s) throw new Error('ไม่พบรหัสเจ้าหน้าที่ ' + code);
  const rs = inspResp_().filter(function (r) { return r.staff_code === code; });
  return { html: inspReportHtml_('รายงานผลการบำรุงรักษารายสัปดาห์', 'ของ ' + inspStaffName_(s) + ' (' + code + ')', inspReportData_(rs, asOf)) };
}
/** ส่งรายงานทางอีเมล: เจ้าหน้าที่ทุกคนที่มีอีเมล (เฉพาะขอบเขตของตน) + ผู้ดูแลระบบ + อีเมลผู้รับเพิ่ม (ภาพรวมทั้งหมด) */
function inspSendReports_(only) {
  const asOf = todayStr_(), valid = function (e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); };
  const sent = {}, out = { sent: 0, noEmail: 0, staff: 0 };
  const mail = function (to, subject, html, text) { to = String(to).trim().toLowerCase(); if (sent[to] || !valid(to)) return; sent[to] = 1; if (typeof MailApp !== 'undefined') { MailApp.sendEmail({ to: to, subject: subject, body: text, htmlBody: html }); out.sent++; } };
  const subj = '[รายงาน] ผลการบำรุงรักษาประจำสัปดาห์ ' + asOf;
  if (only !== 'admins') {
    read_('Staff').forEach(function (s) {
      const code = String(s.staff_code || '').trim(); if (!code || !isActive_(s.active)) return;
      out.staff++;
      if (!valid(s.email)) { out.noEmail++; return; }
      const rs = inspResp_().filter(function (r) { return r.staff_code === code; }); if (!rs.length) return;
      const d = inspReportData_(rs, asOf);
      mail(s.email, subj, inspReportHtml_('รายงานผลการบำรุงรักษา', 'ของ ' + inspStaffName_(s), d), inspReportText_(d));
    });
  }
  const dAll = inspReportData_(null, asOf), htmlAll = inspReportHtml_('รายงานผลการบำรุงรักษารายสัปดาห์', 'ภาพรวมทุกอาคาร ทุกระบบ', dAll), txtAll = inspReportText_(dAll);
  read_('Users').filter(function (x) { return x.role === 'admin' && isActive_(x.active); }).forEach(function (x) { mail(x.email, subj, htmlAll, txtAll); });
  inspSettingStr_('insp_report_to').split(',').forEach(function (e) { mail(e, subj, htmlAll, txtAll); });
  try { mail(Session.getEffectiveUser().getEmail(), subj, htmlAll, txtAll); } catch (e) {}
  return out;
}
function sendInspReportNow(only) { requireAdmin_(); return inspSendReports_(only === 'admins' ? 'admins' : 'all'); }
/** ตัวตั้งเวลารายสัปดาห์ */
function sendInspWeeklyJob() {
  const on = inspSettingStr_('insp_report_on');
  if (!(on === '1' || on === 'true')) return { skipped: 'ปิดการส่งรายงานรายสัปดาห์' };
  return inspSendReports_('all');
}
function installInspWeeklyTrigger() {
  requireAdmin_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendInspWeeklyJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendInspWeeklyJob').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  return 'ตั้งเวลาส่งรายงานทุกวันจันทร์ 07:00 แล้ว (ต้องเปิด "ส่งรายงานรายสัปดาห์" ด้วย)';
}
function removeInspWeeklyTrigger() {
  requireAdmin_(); let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendInspWeeklyJob') { ScriptApp.deleteTrigger(t); n++; } });
  return 'ลบตัวตั้งเวลา ' + n + ' รายการ';
}
// ==INSP-SERVER-END==
