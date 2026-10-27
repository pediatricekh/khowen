/**
 * แอปขอเวร PN/NA  –  ไฟล์ Code.gs
 * ผูกกับ Google Sheet 1 ไฟล์ (Extensions > Apps Script)
 * แท็บที่ใช้: Staff, Config, Requests, Notices  (สร้างให้อัตโนมัติด้วยฟังก์ชัน setup)
 */
const TZ = 'Asia/Bangkok';
const TYPE_TH = { X: 'เวรหยุด', S: 'เวรสั้น', V: 'ลาพักร้อน' };
const DOWN = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
const EXCL = [['60064', '53133']];   // คู่ที่ห้ามขอเวรหยุดตรงกัน
// วันหยุดประจำ (Fix) ของบางคน ใช้เมื่อคอลัมน์ G ในแท็บ Staff ว่าง – ใส่ "-" ในคอลัมน์ G เพื่อยกเลิก
const FIXED_OFF_DEFAULT = { '63003': [0, 6] };   // หัวหน้าแผนก: หยุด อา + ส เสมอ
const DAY_ALIAS = { 'อาทิตย์': 0, 'จันทร์': 1, 'อังคาร': 2, 'พุธ': 3, 'พฤหัส': 4, 'พฤหัสบดี': 4, 'ศุกร์': 5, 'เสาร์': 6 };
const TYPE_KEY = { 'เวรหยุด': 'X', 'เวรสั้น': 'S', 'ลาพักร้อน': 'V' };

/* ---------- หน้าเว็บ ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ขอเวร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- API สำหรับแอปที่ติดตั้งบนหน้าจอ (PWA) ---------- */
const API_FNS_ = { login, save, monthInfo, setReqHours, adminData, setQuota, adminSetRequest, generateRoster, checkRoster, notifyAdmin };
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!Object.prototype.hasOwnProperty.call(API_FNS_, req.fn)) throw new Error('unknown function');
    out = { result: API_FNS_[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- รันครั้งเดียวตอนเริ่มต้น ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, headers) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#e3f3ef');
      s.setFrozenRows(1);
    }
    return s;
  };

  const staff = make('Staff', ['รหัสพนักงาน', 'ชื่อ-สกุล', 'ตำแหน่ง (RN / PN / NA)', 'PIN (ไม่บังคับ)']);
  staff.getRange('A:A').setNumberFormat('@');
  staff.getRange('D:F').setNumberFormat('@');
  staff.getRange('E1').setValue('แอดมิน (ใส่ Y)').setFontWeight('bold').setBackground('#e3f3ef');
  staff.getRange('F1').setValue('พาสทาม (ใส่ Y)').setFontWeight('bold').setBackground('#e3f3ef');
  staff.getRange('G1').setValue('วันหยุดประจำ (เช่น อา,ส)').setFontWeight('bold').setBackground('#e3f3ef');
  staff.getRange('G:G').setNumberFormat('@');
  if (staff.getLastRow() === 1) staff.appendRow(['00000', 'ตัวอย่าง (ลบแถวนี้ทิ้ง)', 'NA', '']);

  const cfg = make('Config', ['รายการ', 'ค่า']);
  if (cfg.getLastRow() === 1) {
    cfg.getRange(2, 1, 3, 2).setValues([
      ['โควต้าเวรหยุดต่อเดือน (วัน) – ลาพักร้อนหักจากโควต้านี้', 8],
      ['วันที่เริ่มเปิดให้ขอ (ของเดือนก่อนหน้า)', 5],
      ['วันที่ปิดรับ (ของเดือนก่อนหน้า)', 15]
    ]);
  }
  cfg.getRange('B5').setNumberFormat('@');
  if (!String(cfg.getRange('B5').getValue()).trim()) {
    cfg.getRange('A5:B5').setValues([['PIN กลางสำหรับทุกคน', '123456']]);
  }
  if (!String(cfg.getRange('B6').getValue()).trim()) {
    cfg.getRange('A6:B8').setValues([
      ['ชม.บังคับต่อเดือน – NA (แก้ทุกเดือน)', 192],
      ['ชม.บังคับต่อเดือน – PN (แก้ทุกเดือน)', 184],
      ['ชม.บังคับต่อเดือน – RN (แก้ทุกเดือน)', 184]
    ]);
  }

  const req = make('Requests', ['เวลาบันทึก', 'รหัส', 'ชื่อ-สกุล', 'ตำแหน่ง', 'เดือน', 'วันที่', 'ประเภท']);
  req.getRange('B:B').setNumberFormat('@');
  req.getRange('E:E').setNumberFormat('@');

  const quota = make('Quota', ['กลุ่ม', 'วัน', 'โควต้าหยุด/วัน (คน)', 'โควต้าเวรสั้น/วัน (คน)', 'เวลาเวรสั้น (ข้อมูลประกอบ)']);
  quota.getRange('A:B').setNumberFormat('@');
  quota.getRange('E:E').setNumberFormat('@');
  if (quota.getLastRow() === 1) {
    const q = [];
    DOWN.forEach((d, i) => {
      const wk = i >= 1 && i <= 5;
      q.push(['RN', d, 1, wk ? 2 : 1, i === 0 ? '9-18 x1' : i === 6 ? '7-17 x1' : '7-17 x1, 8-18 x1']);
    });
    DOWN.forEach(d => q.push(['PN/NA', d, 1, 3, '7-17 x1, 8-17 x2']));
    quota.getRange(2, 1, q.length, 5).setValues(q);
  }

  const slots = make('Slots', ['กลุ่ม', 'เวลาทำงาน', 'ชม.นับ', 'อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส']);
  slots.getRange('A:B').setNumberFormat('@');
  if (slots.getLastRow() === 1) {
    const S = [
      ['PN/NA', '7-17', [1, 1, 1, 1, 1, 1, 1]], ['PN/NA', '8-17', [3, 2, 2, 2, 2, 2, 2]],
      ['PN/NA', '7-20', [2, 2, 2, 2, 2, 2, 2]], ['PN/NA', '9-20', [2, 2, 2, 2, 2, 2, 2]],
      ['PN/NA', '10-16', [0, 1, 1, 1, 1, 0, 0]], ['PN/NA', '12-16', [0, 0, 0, 0, 0, 1, 0]],
      ['PN/NA', '8-20', [0, 0, 0, 0, 0, 0, 1]],
      ['RN', '7-17', [1, 1, 1, 1, 1, 1, 1]], ['RN', '8-18', [0, 1, 1, 1, 1, 1, 0]],
      ['RN', '7-20', [2, 2, 2, 2, 2, 2, 2]], ['RN', '9-20', [0, 1, 1, 1, 1, 1, 1]],
      ['RN', '8-20', [1, 1, 1, 1, 1, 1, 1]], ['RN', '9-18', [1, 0, 0, 0, 0, 0, 0]]
    ].map(r => [r[0], r[1], hoursDef_(r[1])].concat(r[2]));
    slots.getRange(2, 1, S.length, 10).setValues(S);
  }
  const pt = make('PartTime', ['เดือน', 'รหัส', 'วันที่', 'เวลาทำงาน']);
  pt.getRange('A:B').setNumberFormat('@');
  pt.getRange('D:D').setNumberFormat('@');

  const qo = make('QuotaOverride', ['เดือน', 'วันที่', 'กลุ่ม', 'โควต้าหยุด', 'โควต้าเวรสั้น', 'แก้โดย (รหัส)', 'เวลาแก้']);
  qo.getRange('A:A').setNumberFormat('@');
  qo.getRange('C:C').setNumberFormat('@');

  const rh = make('RequiredHours', ['เดือน', 'ตำแหน่ง (RN / PN / NA)', 'ชม.บังคับ', 'แก้โดย (รหัส)', 'เวลาแก้']);
  rh.getRange('A:B').setNumberFormat('@');

  make('Notices', ['เวลา', 'รหัส', 'ชื่อ-สกุล', 'เดือนที่ขอแก้', 'ข้อความถึงแอดมิน', 'สถานะ']);
  Logger.log('สร้างแท็บเรียบร้อยแล้ว');
}

/* ---------- ตัวช่วย ---------- */
function sh_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('ไม่พบแท็บ "' + name + '" ในชีต – ให้รันฟังก์ชัน setup ก่อน');
  return s;
}

function cfg_() {
  const v = sh_('Config').getRange('B2:B5').getValues();
  return {
    quota: Number(v[0][0]) || 8,
    openDay: Number(v[1][0]) || 5,
    closeDay: Number(v[2][0]) || 15,
    pin: String(v[3][0] || '123456').trim()   // PIN กลางของทุกคน
  };
}

function today_() {
  const p = Utilities.formatDate(new Date(), TZ, 'yyyy-M-d').split('-').map(Number);
  return { y: p[0], m: p[1], d: p[2] };
}

function mkey_(y, m) { return y + '-' + ('0' + m).slice(-2); }
function daysIn_(y, m) { return new Date(y, m, 0).getDate(); }

function months_() {
  const t = today_();
  let ny = t.y, nm = t.m + 1;
  if (nm > 12) { nm = 1; ny++; }
  return { t: t, cur: mkey_(t.y, t.m), next: mkey_(ny, nm), ny: ny, nm: nm };
}

/** ตรวจรหัส + PIN ฝั่งเซิร์ฟเวอร์ทุกครั้ง (ถ้าไม่มี PIN เฉพาะคน ใช้ PIN กลางจากแท็บ Config) */
function auth_(code, pin) {
  code = String(code || '').trim();
  if (!code) return null;
  const rows = sh_('Staff').getDataRange().getValues().slice(1);
  for (const r of rows) {
    if (String(r[0]).trim() !== code) continue;
    const need = String(r[3] || '').trim() || cfg_().pin;
    if (need !== String(pin || '').trim()) return null;
    return {
      code: code,
      name: String(r[1]).trim(),
      role: String(r[2] || 'PN/NA').trim().toUpperCase(),
      admin: String(r[4] || '').trim().toUpperCase() === 'Y',
      fixed: fixedOff_(code, r[6])
    };
  }
  return null;
}

/** คอลัมน์ G ของแท็บ Staff (เช่น "อา,ส") → [0..6]; ว่าง = ใช้ค่าเริ่มต้นของรหัสนั้น; "-" = ไม่มี */
function fixedOff_(code, cell) {
  const t = String(cell === undefined || cell === null ? '' : cell).trim();
  if (!t) return (FIXED_OFF_DEFAULT[code] || []).slice();
  const out = [];
  t.split(/[,\s/;]+/).forEach(w => {
    w = w.replace(/\./g, '').trim();
    const i = w in DAY_ALIAS ? DAY_ALIAS[w] : DOWN.indexOf(w);
    if (i >= 0 && out.indexOf(i) < 0) out.push(i);
  });
  return out;
}

function grp_(role) { return role === 'RN' ? 'RN' : 'PN/NA'; }

/** โควต้าจากแท็บ Quota → { 'RN': { 0..6: {off, short, note} }, 'PN/NA': {...} } */
function caps_() {
  const s = sh_('Quota'), last = s.getLastRow(), out = {};
  if (last < 2) return out;
  s.getRange(2, 1, last - 1, 5).getValues().forEach(r => {
    const g = String(r[0]).trim(), i = DOWN.indexOf(String(r[1]).trim());
    if (!g || i < 0) return;
    (out[g] = out[g] || {})[i] = { off: Number(r[2]) || 0, short: Number(r[3]) || 0, note: String(r[4] || '') };
  });
  return out;
}

/** คำขอของทุกคนในเดือนนั้น */
function allReq_(key) {
  const s = sh_('Requests'), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 7).getValues()
    .filter(r => String(r[4]).trim() === key)
    .map(r => ({ code: String(r[1]).trim(), name: String(r[2]), role: String(r[3]).trim().toUpperCase(),
                 day: Number(r[5]), type: TYPE_KEY[String(r[6]).trim()] || 'X' }));
}

/** โควต้าที่แอดมินแก้เป็นรายวัน → { 'RN': { วันที่: {off, short} } } */
function overrides_(key) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QuotaOverride');
  const out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 5).getValues().forEach(r => {
    if (String(r[0]).trim() !== key) return;
    const g = String(r[2]).trim();
    (out[g] = out[g] || {})[Number(r[1])] = { off: Number(r[3]) || 0, short: Number(r[4]) || 0 };
  });
  return out;
}

/** ที่ว่างของแต่ละวัน (ไม่นับคำขอของตัวเอง) + วันที่ห้ามหยุดเพราะซ้อนกับคู่ที่กำหนด */
function limits_(u, y, m, key) {
  const g = grp_(u.role), C = caps_()[g] || {}, rows = allReq_(key), n = daysIn_(y, m), ov = overrides_(key)[g] || {};
  const offLeft = {}, shortLeft = {}, blocked = [];
  const pair = EXCL.find(p => p.indexOf(u.code) >= 0);
  const partner = pair ? pair.filter(c => c !== u.code)[0] : null;
  for (let d = 1; d <= n; d++) {
    const c0 = C[new Date(y, m - 1, d).getDay()] || { off: 0, short: 0 };
    const c = ov[d] || c0;   // โควต้าที่แอดมินแก้รายวัน มาก่อนค่าปกติ
    let o = 0, sh = 0;
    rows.forEach(r => {
      if (r.day !== d || r.code === u.code || grp_(r.role) !== g) return;
      if (r.type === 'S') sh++; else o++;
    });
    offLeft[d] = c.off - o;
    shortLeft[d] = c.short - sh;
    if (partner && rows.some(r => r.code === partner && r.day === d && r.type !== 'S')) blocked.push(d);
  }
  return { offLeft: offLeft, shortLeft: shortLeft, blocked: blocked };
}

/** คำขอของคนนี้ในเดือนนั้น → { วันที่: 'X'|'S'|'V' } (อ่านเฉพาะของตัวเอง) */
function getReq_(code, key) {
  const s = sh_('Requests');
  const last = s.getLastRow();
  const out = {};
  if (last < 2) return out;
  s.getRange(2, 1, last - 1, 7).getValues().forEach(r => {
    if (String(r[1]).trim() === code && String(r[4]).trim() === key) {
      out[Number(r[5])] = TYPE_KEY[String(r[6]).trim()] || 'X';
    }
  });
  return out;
}

/** คำขอของตัวเองทุกเดือน → { 'yyyy-mm': { วันที่: 'X'|'S'|'V' } } (ใช้ดูปฏิทินย้อนหลัง/ล่วงหน้า) */
function allOwn_(code) {
  const s = sh_('Requests'), last = s.getLastRow(), out = {};
  if (last < 2) return out;
  s.getRange(2, 1, last - 1, 7).getValues().forEach(r => {
    if (String(r[1]).trim() !== code) return;
    const k = String(r[4]).trim();
    (out[k] = out[k] || {})[Number(r[5])] = TYPE_KEY[String(r[6]).trim()] || 'X';
  });
  return out;
}

function state_(u) {
  const m = months_(), c = cfg_();
  return {
    user: u,
    open: !!u.admin || (m.t.d >= c.openDay && m.t.d <= c.closeDay),   // แอดมินขอเวรของตัวเองได้ตลอด ไม่ติดวันที่ 5-15
    openDay: c.openDay, closeDay: c.closeDay, quota: c.quota,
    all: allOwn_(u.code),
    missing: missingHours_(u),
    cur:  { key: m.cur,  y: m.t.y, m: m.t.m, req: getReq_(u.code, m.cur), hours: hoursFor_(u, m.cur) },
    next: { key: m.next, y: m.ny,  m: m.nm,  req: getReq_(u.code, m.next), lim: limits_(u, m.ny, m.nm, m.next), hours: hoursFor_(u, m.next) }
  };
}

/* ---------- ฟังก์ชันที่หน้าเว็บเรียก ---------- */
function login(code, pin) {
  const u = auth_(code, pin);
  if (!u) return { ok: false, msg: 'ไม่พบรหัสพนักงาน หรือ PIN ไม่ถูกต้อง' };
  return { ok: true, state: state_(u) };
}

/** เดือนเป้าหมายของการขอเวร: แอดมินเลือกเดือนไหนก็ได้ (yyyy-mm), คนอื่นได้เฉพาะเดือนหน้า */
function target_(u, m, key) {
  key = String(key || '').trim();
  if (u.admin && /^\d{4}-(0[1-9]|1[0-2])$/.test(key)) {
    return { key: key, y: Number(key.slice(0, 4)), m: Number(key.slice(5)) };
  }
  return { key: m.next, y: m.ny, m: m.nm };
}

/** ข้อมูลเดือนใดก็ได้สำหรับแอดมิน (คำขอของตัวเอง + ที่ว่างรายวัน) */
function monthInfo_(u, tg) {
  return { key: tg.key, y: tg.y, m: tg.m, req: getReq_(u.code, tg.key), lim: limits_(u, tg.y, tg.m, tg.key), hours: hoursFor_(u, tg.key) };
}

function monthInfo(code, pin, key) {
  const u = auth_(code, pin);
  if (!u) return { ok: false, msg: 'ยืนยันตัวตนไม่ผ่าน' };
  if (!u.admin) return { ok: false, msg: 'ไม่มีสิทธิ์แอดมิน' };
  return { ok: true, info: monthInfo_(u, target_(u, months_(), key)) };
}

/** บันทึกคำขอของเดือนหน้า (แอดมินเลือกเดือนอื่นได้) แทนที่ของเดิมทั้งเดือน พร้อมตรวจโควต้า */
function save(code, pin, items, key) {
  const u = auth_(code, pin);
  if (!u) return { ok: false, msg: 'ยืนยันตัวตนไม่ผ่าน' };

  const m = months_(), c = cfg_(), tg = target_(u, m, key);
  if (!u.admin && !(m.t.d >= c.openDay && m.t.d <= c.closeDay)) {
    return { ok: false, msg: 'ปิดรับคำขอแล้ว กรุณาแจ้งแอดมิน', state: state_(u) };
  }

  const n = daysIn_(tg.y, tg.m), rows = [];
  let used = 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lim = limits_(u, tg.y, tg.m, tg.key);
    const fail = msg => ({ ok: false, msg: msg, state: state_(u), info: u.admin ? monthInfo_(u, tg) : undefined });
    for (const k in items) {
      const d = Number(k), t = items[k];
      if (!(d >= 1 && d <= n) || !TYPE_TH[t]) return { ok: false, msg: 'ข้อมูลไม่ถูกต้อง' };
      if ((u.fixed || []).indexOf(new Date(tg.y, tg.m - 1, d).getDay()) >= 0) {
        continue;   // วันหยุดประจำ: ไม่บันทึกและไม่นับโควต้า (ไม่แจ้งอะไร)
      }
      if (t === 'S') {
        if (lim.shortLeft[d] < 1) return fail('เวรสั้นวันที่ ' + d + ' เต็มแล้ว');
      } else {
        used++;
        if (lim.offLeft[d] < 1) return fail('วันหยุดวันที่ ' + d + ' เต็มแล้ว');
        if (lim.blocked.indexOf(d) >= 0) return fail('วันที่ ' + d + ' ไม่สามารถหยุดได้');
      }
      rows.push([new Date(), u.code, u.name, u.role, tg.key, d, TYPE_TH[t]]);
    }
    if (used > c.quota) return fail('เกินโควต้าเวรหยุด ' + c.quota + ' วัน (ลาพักร้อนหักจากโควต้านี้)');
    rows.sort((x, y) => x[5] - y[5]);

    const s = sh_('Requests');
    const last = s.getLastRow();
    if (last > 1) {
      const vals = s.getRange(2, 1, last - 1, 7).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][1]).trim() === u.code && String(vals[i][4]).trim() === tg.key) s.deleteRow(i + 2);
      }
    }
    if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, state: state_(u), info: u.admin ? monthInfo_(u, tg) : undefined };
}

/** ข้อมูลสำหรับหน้าแอดมิน: คำขอของทุกคน ทั้งเดือนปัจจุบันและเดือนหน้า */
function ymOf_(key) {
  key = String(key || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(key)) return null;
  return { y: Number(key.slice(0, 4)), mm: Number(key.slice(5)) };
}

function adminData(code, pin, extra) {
  const u = auth_(code, pin);
  if (!u || !u.admin) return { ok: false, msg: 'ไม่มีสิทธิ์แอดมิน' };
  const m = months_();
  const mk = key => { const t = ymOf_(key); return { key: key, y: t.y, m: t.mm, n: daysIn_(t.y, t.mm), rows: allReq_(key), hours: reqHoursMonth_(key) }; };
  const staff = sh_('Staff').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({ code: String(r[0]).trim(), name: String(r[1]).trim(), role: String(r[2] || 'PN/NA').trim().toUpperCase() }));
  // แอดมินดู/แก้ได้ทุกเดือน: เดือนปัจจุบัน + เดือนหน้า + เดือนอื่นที่เลือกเพิ่ม
  const keys = [m.cur, m.next];
  (Array.isArray(extra) ? extra : []).forEach(k => { k = String(k || '').trim(); if (ymOf_(k) && keys.indexOf(k) < 0) keys.push(k); });
  keys.sort();
  const ov = {}; keys.forEach(k => ov[k] = overrides_(k));
  return { ok: true, caps: caps_(), ov: ov, staff: staff, months: keys.map(mk) };
}

/** แอดมินแก้โควต้ารายวัน (off = null → คืนค่าปกติ) และแจ้งเตือนถ้ามีคนขอไว้เกินโควต้าใหม่ */
function setQuota(code, pin, key, group, day, off, short) {
  const u = auth_(code, pin);
  if (!u || !u.admin) return { ok: false, msg: 'ไม่มีสิทธิ์แอดมิน' };
  const t = ymOf_(key);
  if (!t) return { ok: false, msg: 'เดือนไม่ถูกต้อง' };
  key = String(key).trim();
  const y = t.y, mm = t.mm;
  day = Number(day);
  if (['RN', 'PN/NA'].indexOf(group) < 0 || !(day >= 1 && day <= daysIn_(y, mm))) return { ok: false, msg: 'ข้อมูลไม่ถูกต้อง' };
  const reset = off === null || off === undefined;
  if (!reset) {
    off = Number(off); short = Number(short);
    if (!(off >= 0 && off <= 50 && short >= 0 && short <= 50) || off % 1 || short % 1) {
      return { ok: false, msg: 'โควต้าต้องเป็นจำนวนเต็ม 0-50' };
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh_('QuotaOverride');
    const last = s.getLastRow();
    if (last > 1) {
      const vals = s.getRange(2, 1, last - 1, 3).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][0]).trim() === key && Number(vals[i][1]) === day && String(vals[i][2]).trim() === group) s.deleteRow(i + 2);
      }
    }
    if (!reset) s.getRange(s.getLastRow() + 1, 1, 1, 7).setValues([[key, day, group, off, short, u.code, new Date()]]);
  } finally {
    lock.releaseLock();
  }

  // ตรวจว่ามีคนขอไว้เกินโควต้าใหม่หรือไม่
  const def = (caps_()[group] || {})[new Date(y, mm - 1, day).getDay()] || { off: 0, short: 0 };
  const eff = (overrides_(key)[group] || {})[day] || def;
  const rows = allReq_(key).filter(r => r.day === day && grp_(r.role) === group);
  const offR = rows.filter(r => r.type !== 'S'), shR = rows.filter(r => r.type === 'S');
  const over = [];
  if (offR.length > eff.off) over.push({ label: 'หยุด/ลาพักร้อน', used: offR.length, cap: eff.off, names: offR.map(r => r.name) });
  if (shR.length > eff.short) over.push({ label: 'เวรสั้น', used: shR.length, cap: eff.short, names: shR.map(r => r.name) });
  return { ok: true, over: over, data: adminData(code, pin, [key]) };
}

/** รายการที่เกินโควต้าของวันนั้น (ใช้แจ้งเตือนแอดมิน) */
function overFor_(key, group, day, y, mm) {
  const def = (caps_()[group] || {})[new Date(y, mm - 1, day).getDay()] || { off: 0, short: 0 };
  const eff = (overrides_(key)[group] || {})[day] || def;
  const rows = allReq_(key).filter(r => r.day === day && grp_(r.role) === group);
  const offR = rows.filter(r => r.type !== 'S'), shR = rows.filter(r => r.type === 'S');
  const over = [];
  if (offR.length > eff.off) over.push({ label: 'เวรหยุด/ลาพักร้อน', used: offR.length, cap: eff.off, names: offR.map(r => r.name) });
  if (shR.length > eff.short) over.push({ label: 'เวรสั้น', used: shR.length, cap: eff.short, names: shR.map(r => r.name) });
  return over;
}

/** แอดมินเพิ่ม/แก้/ลบเวรแทนพนักงาน (type = null → ลบ) แก้ได้แม้เกินโควต้า แต่จะแจ้งเตือน */
function adminSetRequest(code, pin, key, targetCode, day, type) {
  const u = auth_(code, pin);
  if (!u || !u.admin) return { ok: false, msg: 'ไม่มีสิทธิ์แอดมิน' };
  const t = ymOf_(key);
  if (!t) return { ok: false, msg: 'เดือนไม่ถูกต้อง' };
  key = String(key).trim();
  const y = t.y, mm = t.mm;
  day = Number(day);
  targetCode = String(targetCode || '').trim();
  const del = type === null || type === undefined;
  if (!(day >= 1 && day <= daysIn_(y, mm)) || (!del && !TYPE_TH[type])) return { ok: false, msg: 'ข้อมูลไม่ถูกต้อง' };
  const st = sh_('Staff').getDataRange().getValues().slice(1).find(r => String(r[0]).trim() === targetCode);
  if (!st) return { ok: false, msg: 'ไม่พบพนักงานในแท็บ Staff' };
  const name = String(st[1]).trim(), role = String(st[2] || 'PN/NA').trim().toUpperCase();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh_('Requests');
    const last = s.getLastRow();
    if (last > 1) {
      const vals = s.getRange(2, 1, last - 1, 7).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][1]).trim() === targetCode && String(vals[i][4]).trim() === key && Number(vals[i][5]) === day) s.deleteRow(i + 2);
      }
    }
    if (!del) s.getRange(s.getLastRow() + 1, 1, 1, 7).setValues([[new Date(), targetCode, name, role, key, day, TYPE_TH[type]]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, over: overFor_(key, grp_(role), day, y, mm), data: adminData(code, pin, [key]) };
}

/* ===================== จัดเวรอัตโนมัติ ===================== */
const MAXRUN = 7;   // ทำงานติดกันเกินกี่วันจึงแจ้งเตือน
// รหัสที่มีเวรประจำ (นับรวมในแผนกทุกวัน ถ้าไม่ขอหยุด): จ-พฤ 10-16, ศ 12-16, ส 8-20, อา 8-17
// ถ้าขอหยุด ให้รหัส 'backup' รับเวรนั้นแทนเสมอ
const FIXED = { code: '53133', backup: '60064', byDow: { 0: '8-17', 1: '10-16', 2: '10-16', 3: '10-16', 4: '10-16', 5: '12-16', 6: '8-20' } };
const LABEL_RE = /^\d{1,2}(\.\d+)?-\d{1,2}(\.\d+)?$/;

function hoursDef_(lb) { const p = String(lb).split('-').map(Number); const d = (p[1] || 0) - (p[0] || 0); return d >= 8 ? d - 1 : d; }
function endH_(lb) { return Number(String(lb).split('-')[1]) || 0; }
function hrsOf_(slots, lb) { return slots[lb] ? slots[lb].hours : hoursDef_(lb); }

function slotsData_() {
  const s = sh_('Slots'), last = s.getLastRow(), out = {};
  if (last < 2) return out;
  s.getRange(2, 1, last - 1, 10).getValues().forEach(r => {
    const g = String(r[0]).trim(), lb = String(r[1]).trim();
    if (!g || !lb) return;
    (out[g] = out[g] || {})[lb] = { hours: Number(r[2]) || hoursDef_(lb), cnt: r.slice(3, 10).map(x => Number(x) || 0) };
  });
  return out;
}

/** ชม.บังคับที่แอดมินกรอกรายเดือน → { NA, PN, RN } (ยังไม่กรอก = null) */
function reqHoursMonth_(key) {
  const out = { NA: null, PN: null, RN: null };
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RequiredHours');
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 3).getValues().forEach(r => {
    const p = String(r[1]).trim().toUpperCase();
    if (String(r[0]).trim() === key && p in out && r[2] !== '') out[p] = Number(r[2]);
  });
  return out;
}

/** ชม.บังคับสำหรับจัดเวร: ค่ารายเดือน ถ้ายังไม่กรอกใช้ค่าเริ่มต้นจากแท็บ Config */
function reqHours_(key) {
  const v = sh_('Config').getRange('B6:B8').getValues(), mo = reqHoursMonth_(key);
  return { NA: mo.NA || Number(v[0][0]) || 192, PN: mo.PN || Number(v[1][0]) || 184, RN: mo.RN || Number(v[2][0]) || 184 };
}

function roleKey_(role) { return role === 'RN' ? 'RN' : role === 'NA' ? 'NA' : 'PN'; }

/** ชม.บังคับของเดือนนั้น: แอดมินเห็นครบทุกตำแหน่ง ผู้ใช้ทั่วไปเห็นเฉพาะของตำแหน่งตัวเอง */
function hoursFor_(u, key) {
  const h = reqHoursMonth_(key);
  if (u.admin) return h;
  const k = roleKey_(u.role), o = { NA: null, PN: null, RN: null };
  o[k] = h[k];
  return o;
}

/** แอดมินกรอกชม.บังคับของเดือน (ล่วงหน้าได้) แยกตามตำแหน่ง */
function setReqHours(code, pin, key, vals) {
  const u = auth_(code, pin);
  if (!u || !u.admin) return { ok: false, msg: 'ไม่มีสิทธิ์แอดมิน' };
  key = String(key || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(key)) return { ok: false, msg: 'เดือนไม่ถูกต้อง' };
  const roles = ['NA', 'PN', 'RN'], clean = {};
  for (const p of roles) {
    const v = Number(vals && vals[p]);
    if (!(v > 0 && v <= 400)) return { ok: false, msg: 'ชม.บังคับ ' + p + ' ต้องเป็นตัวเลข 1-400' };
    clean[p] = v;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh_('RequiredHours');
    const last = s.getLastRow();
    if (last > 1) {
      const rows = s.getRange(2, 1, last - 1, 2).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]).trim() === key) s.deleteRow(i + 2);
      }
    }
    const add = roles.map(p => [key, p, clean[p], u.code, new Date()]);
    s.getRange(s.getLastRow() + 1, 1, add.length, 5).setValues(add);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, key: key, hours: reqHoursMonth_(key), missing: missingHours_(u) };
}

/** แอดมินยังไม่กรอกชม.บังคับของเดือนหน้า และถึงวันเปิดให้ขอเวรแล้ว → รายชื่อตำแหน่งที่ขาด */
function missingHours_(u) {
  if (!u.admin) return [];
  const m = months_(), c = cfg_();
  if (m.t.d < c.openDay) return [];
  const h = reqHoursMonth_(m.next);
  return ['NA', 'PN', 'RN'].filter(p => !h[p]);
}

/** จัดเวรให้ครบแพทเทิน เกลี่ยวันทำงาน/ชั่วโมงให้ใกล้กัน คนที่ขอหยุดมากได้ทำงานน้อยลงตามจริง */
function buildRoster_(P) {
  const n = daysIn_(P.y, P.m), alerts = [];
  const dow = d => new Date(P.y, P.m - 1, d).getDay();
  const hrs = lb => hrsOf_(P.slots, lb);
  const byCode = {};
  const grid = {};
  P.staff.forEach(s => {
    byCode[s.code] = s; grid[s.code] = {};
    const rq = P.req[s.code] || {};
    for (const d in rq) { if (rq[d] === 'X') grid[s.code][d] = 'XR'; else if (rq[d] === 'V') grid[s.code][d] = 'V'; }
  });
  const free = (c, d) => grid[c] && grid[c][d] === undefined;

  // ช่องเวรที่ต้องจัดของแต่ละวัน
  const left = {};
  for (let d = 1; d <= n; d++) {
    left[d] = [];
    for (const lb in P.slots) for (let i = 0; i < P.slots[lb].cnt[dow(d)]; i++) left[d].push(lb);
  }
  const take = (d, lb) => { const i = left[d].indexOf(lb); if (i < 0) return false; left[d].splice(i, 1); return true; };

  // พาสทาม: ลงตามแท็บ PartTime
  const warnedPT = {};
  P.part.forEach(e => {
    const s = byCode[e.code];
    if (!s) {
      if (!warnedPT[e.code]) {
        warnedPT[e.code] = 1;
        alerts.push('แท็บ PartTime มีรหัส ' + e.code + ' ที่ไม่อยู่ในแท็บ Staff ของกลุ่ม ' + P.group + ' (ข้ามไป) – เพิ่มชื่อในแท็บ Staff และใส่ Y ช่อง F');
      }
      return;
    }
    if (e.day < 1 || e.day > n) return;
    if (!take(e.day, e.shift)) alerts.push('พาสทาม ' + s.name + ' วันที่ ' + e.day + ' เวลา ' + e.shift + ' ไม่มีในแพทเทิน (นับเป็นคนเกิน)');
    grid[e.code][e.day] = e.shift;
  });

  // เวรประจำ + คนสำรอง
  const fx = byCode[FIXED.code], bk = byCode[FIXED.backup];
  if (fx && !fx.part) {
    for (let d = 1; d <= n; d++) {
      const lb = FIXED.byDow[dow(d)];
      if (!lb) { if (free(fx.code, d)) grid[fx.code][d] = '-'; continue; }
      if (free(fx.code, d)) {
        if (take(d, lb)) grid[fx.code][d] = lb; else { grid[fx.code][d] = '-'; alerts.push('วันที่ ' + d + ' ไม่มีช่องเวร ' + lb + ' ในแพทเทิน'); }
      } else if (left[d].indexOf(lb) >= 0) {
        if (bk && !bk.part && free(bk.code, d) && take(d, lb)) grid[bk.code][d] = lb;
        else alerts.push('วันที่ ' + d + ' ' + fx.name + ' หยุด และไม่มีคนสำรองรับเวร ' + lb);
      }
    }
  }

  // คนที่ระบบจัดให้ + เป้าหมายจำนวนวันทำงาน (เกลี่ยให้เท่ากัน คนที่ขอหยุดเยอะได้น้อยลงเอง)
  const pool = P.staff.filter(s => !s.part && !(fx && s.code === FIXED.code));
  const st = {};
  pool.forEach(s => st[s.code] = { streak: 0, dstr: 0, hours: 0, nw: 0 });
  let total = 0;
  for (let d = 1; d <= n; d++) total += left[d].length;
  const freeCnt = {}, base = {}, suf = {};
  pool.forEach(s => {
    let f = 0, b = 0;
    suf[s.code] = new Array(n + 2).fill(0);
    for (let d = n; d >= 1; d--) {
      const v = grid[s.code][d];
      if (v === undefined) f++; else if (LABEL_RE.test(v)) b++;
      suf[s.code][d] = f;
    }
    freeCnt[s.code] = f; base[s.code] = b;
  });
  const wOf = L => { const w = {}; pool.forEach(s => w[s.code] = Math.max(0, Math.min(freeCnt[s.code], L - base[s.code]))); return w; };
  const sum = w => pool.reduce((a, s) => a + w[s.code], 0);
  let L = 0;
  while (L <= n + 1 && sum(wOf(L)) < total) L++;
  let w;
  if (L > n + 1) {
    w = wOf(n + 1);
    alerts.push('พนักงานไม่พอสำหรับแพทเทินเดือนนี้ (ขาดประมาณ ' + (total - sum(w)) + ' เวร)');
  } else {
    w = wOf(Math.max(L - 1, 0));
    const full = wOf(L);
    const rem = total - sum(w);
    pool.filter(s => full[s.code] > w[s.code])
      .sort((a, b) => (base[a.code] + w[a.code]) - (base[b.code] + w[b.code]))
      .slice(0, Math.max(rem, 0)).forEach(s => w[s.code]++);
  }

  const dpen = i => i.dstr >= 2 ? 40 : i.dstr === 1 ? 8 : 0;
  for (let d = 1; d <= n; d++) {
    const N = left[d].length;
    const scored = pool.filter(s => free(s.code, d)).map(s => {
      const i = st[s.code], need = w[s.code] - i.nw, r = suf[s.code][d];
      return { s: s, i: i, sreq: (P.req[s.code] || {})[d] === 'S', must: need > 0 && need >= r,
               ratio: (r > 0 ? need / r : 0) - 0.12 * i.streak };
    });
    const must = scored.filter(x => x.sreq || x.must).sort((a, b) => (b.sreq - a.sreq) || (a.i.streak - b.i.streak));
    const rest = scored.filter(x => !(x.sreq || x.must)).sort((a, b) => b.ratio - a.ratio || a.i.hours - b.i.hours);
    const chosen = must.slice(0, N);
    rest.forEach(x => { if (chosen.length < N && x.i.streak < MAXRUN) chosen.push(x); });
    rest.forEach(x => { if (chosen.length < N && chosen.indexOf(x) < 0) chosen.push(x); });

    // แบ่งเวลาเวร: คนขอเวรสั้นได้เวรที่เลิกเร็วที่สุด / เวรยาวให้คนที่ชั่วโมงสะสมน้อยสุด และเลี่ยง D ติดกัน
    const labels = left[d].slice(), todo = chosen.slice();
    const give = (x, lb) => { grid[x.s.code][d] = lb; labels.splice(labels.indexOf(lb), 1); todo.splice(todo.indexOf(x), 1); };
    todo.filter(x => x.sreq).forEach(x => {
      if (!labels.length) return;
      give(x, labels.slice().sort((a, b) => endH_(a) - endH_(b) || hrs(a) - hrs(b))[0]);
    });
    labels.slice().sort((a, b) => hrs(b) - hrs(a)).forEach(lb => {
      if (!todo.length) return;
      const isD = endH_(lb) >= 20;
      todo.sort((a, b) => (a.i.hours + (isD ? dpen(a.i) : 0)) - (b.i.hours + (isD ? dpen(b.i) : 0)));
      give(todo[0], lb);
    });

    pool.forEach(s => {
      const v = grid[s.code][d], i = st[s.code];
      if (v !== undefined && LABEL_RE.test(v)) {
        i.streak++; i.hours += hrs(v); i.dstr = endH_(v) >= 20 ? i.dstr + 1 : 0;
        if (chosen.some(x => x.s.code === s.code)) i.nw++;
      } else {
        i.streak = 0; i.dstr = 0;
        if (v === undefined) grid[s.code][d] = 'X';
      }
    });
  }
  P.staff.forEach(s => { for (let d = 1; d <= n; d++) if (grid[s.code][d] === undefined) grid[s.code][d] = 'X'; });
  return { grid: grid, alerts: alerts };
}

/** สรุปวันทำงาน/ชั่วโมง/OT และตรวจทำงานติดกันเกิน MAXRUN วัน */
function analyze_(grid, P) {
  const n = daysIn_(P.y, P.m), stats = [], alerts = [], dayHave = [], dayNeed = [];
  // ตรวจทุกวันเทียบแพทเทิน: เกิน / ขาด แยกตามเวลาทำงาน
  for (let d = 1; d <= n; d++) {
    const dw = new Date(P.y, P.m - 1, d).getDay();
    const have = {}, need = {};
    let th = 0, tn = 0;
    for (const lb in P.slots) { need[lb] = P.slots[lb].cnt[dw]; tn += need[lb]; }
    P.staff.forEach(s => {
      const v = String(((grid[s.code] || {})[d]) === undefined ? '' : grid[s.code][d]).trim();
      if (LABEL_RE.test(v)) { have[v] = (have[v] || 0) + 1; th++; }
    });
    const over = [], under = [];
    Object.keys(need).concat(Object.keys(have).filter(k => !(k in need))).forEach(lb => {
      const h = have[lb] || 0, nd = need[lb] || 0;
      if (h > nd) over.push(lb + ' (ได้ ' + h + '/ต้องการ ' + nd + ')');
      else if (h < nd) under.push(lb + ' (ได้ ' + h + '/ต้องการ ' + nd + ')');
    });
    if (over.length) alerts.push('วันที่ ' + d + ' เกินแพทเทิน: ' + over.join(', '));
    if (under.length) alerts.push('วันที่ ' + d + ' ขาดคน: ' + under.join(', '));
    dayHave.push(th); dayNeed.push(tn);
  }
  P.staff.forEach(s => {
    const g = grid[s.code] || {};
    let work = 0, off = 0, hours = 0, run = 0, start = 0;
    const endRun = d => {
      if (run > MAXRUN) alerts.push(s.name + ' ทำงานติดกัน ' + run + ' วัน (วันที่ ' + start + '–' + (d - 1) + ')');
      run = 0;
    };
    for (let d = 1; d <= n; d++) {
      const v = String(g[d] === undefined ? '' : g[d]).trim();
      if (LABEL_RE.test(v)) { if (run === 0) start = d; run++; work++; hours += hrsOf_(P.slots, v); }
      else { endRun(d); if (v === 'X' || v === 'XR' || v === 'V') off++; }
    }
    endRun(n + 1);
    const required = s.part ? 0 : (P.reqH[s.role] || P.reqH.PN);
    stats.push({ code: s.code, name: s.name, role: s.role, work: work, off: off, hours: hours, required: required, ot: hours - required });
  });
  return { stats: stats, alerts: alerts, dayHave: dayHave, dayNeed: dayNeed };
}

function rosterSheetName_(group, key) { return 'เวร ' + (group === 'RN' ? 'RN' : 'PN-NA') + ' ' + key; }

function rosterInput_(key, y, mm, group) {
  const slotsAll = slotsData_();
  if (!slotsAll[group]) throw new Error('ไม่พบแพทเทินของกลุ่ม ' + group + ' ในแท็บ Slots – ให้รันฟังก์ชัน setup ก่อน');
  const staff = sh_('Staff').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({ code: String(r[0]).trim(), name: String(r[1]).trim(), role: String(r[2] || 'PN/NA').trim().toUpperCase(),
                 part: String(r[5] || '').trim().toUpperCase() === 'Y',
                 fixed: fixedOff_(String(r[0]).trim(), r[6]) }))
    .filter(s => grp_(s.role) === group);
  const req = {};
  allReq_(key).forEach(r => { (req[r.code] = req[r.code] || {})[r.day] = r.type; });
  // วันหยุดประจำ → ลงเป็นวันหยุดในตารางเลย (ไม่นับโควต้าขอเวร) ถ้าไม่ได้ขอลา/เวรสั้นไว้เอง
  staff.forEach(s => {
    for (let d = 1; d <= daysIn_(y, mm); d++) {
      if (s.fixed.indexOf(new Date(y, mm - 1, d).getDay()) >= 0 && !(req[s.code] && req[s.code][d])) {
        (req[s.code] = req[s.code] || {})[d] = 'X';
      }
    }
  });
  const part = [];
  const ps = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartTime');
  if (ps && ps.getLastRow() > 1) {
    ps.getRange(2, 1, ps.getLastRow() - 1, 4).getValues().forEach(r => {
      if (String(r[0]).trim() === key) part.push({ code: String(r[1]).trim(), day: Number(r[2]), shift: String(r[3]).trim() });
    });
  }
  return { y: y, m: mm, key: key, group: group, staff: staff, req: req, slots: slotsAll[group], part: part, reqH: reqHours_(key) };
}

function writeRoster_(P, grid, A, alerts, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  const n = daysIn_(P.y, P.m);
  const head = ['รหัส', 'ชื่อ-สกุล', 'ตำแหน่ง'], wk = ['', '', ''];
  for (let d = 1; d <= n; d++) { head.push(d); wk.push(DOWN[new Date(P.y, P.m - 1, d).getDay()]); }
  head.push('วันทำงาน', 'วันหยุด', 'ชม.ทำงาน', 'ชม.บังคับ', 'OT'); wk.push('', '', '', '', '');
  const rows = [], colors = [];
  P.staff.forEach((s, k) => {
    const row = [s.code, s.name, s.role], col = ['', '', ''];
    for (let d = 1; d <= n; d++) {
      const v = grid[s.code][d];
      row.push(v === 'XR' ? 'X' : v);
      const asked = (P.req[s.code] || {})[d];
      // แดง = ขอหยุด / ลาพักร้อน, เขียว = เวรสั้นที่ขอ, ดำ = ที่ระบบจัดให้
      col.push(v === 'XR' || v === 'V' ? '#d92d20' : (LABEL_RE.test(String(v)) && asked === 'S') ? '#188038' : '#000000');
    }
    const t = A.stats[k];
    row.push(t.work, t.off, t.hours, t.required, t.ot); col.push('', '', '', '', '');
    rows.push(row); colors.push(col);
  });
  const w = head.length;
  sh.getRange(1, 1).setValue('ตารางเวร ' + P.group + ' เดือน ' + P.key + '  (สีแดง = ขอหยุด/ลาพักร้อน, สีเขียว = เวรสั้นที่ขอ, สีดำ = ที่ระบบจัดให้, - = ไม่จัด)');
  sh.getRange(2, 1, 1, w).setValues([wk]);
  sh.getRange(3, 1, 1, w).setValues([head]).setFontWeight('bold').setBackground('#e3f3ef');
  sh.getRange(4, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(4, 4, rows.length, n).setNumberFormat('@');
  const body = sh.getRange(4, 1, rows.length, w);
  body.setValues(rows);
  body.setFontColors(colors).setHorizontalAlignment('center');
  sh.getRange(4, 1, rows.length, 3).setHorizontalAlignment('left');
  const sr = 4 + rows.length;
  const hv = ['', 'รวมคนที่ขึ้นเวร', ''], nd = ['', 'ตามแพทเทิน', ''], cs = ['', '', ''];
  for (let d = 1; d <= n; d++) {
    hv.push(A.dayHave[d - 1]); nd.push(A.dayNeed[d - 1]);
    cs.push(A.dayHave[d - 1] === A.dayNeed[d - 1] ? '#000000' : '#d92d20');
  }
  while (hv.length < w) { hv.push(''); nd.push(''); cs.push(''); }
  const sg = sh.getRange(sr, 1, 2, w);
  sg.setValues([hv, nd]);
  sg.setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(sr, 1, 1, w).setFontColors([cs]);
  sh.getRange(sr, 1, 2, 3).setHorizontalAlignment('left');
  const r0 = sr + 3;
  const msgs = [['แจ้งเตือน (' + alerts.length + ')']].concat(alerts.length ? alerts.map(a => ['⚠ ' + a]) : [['ไม่มี']]);
  sh.getRange(r0, 1, msgs.length, 1).setValues(msgs);
  sh.getRange(r0, 1).setFontWeight('bold');
  sh.setFrozenRows(3); sh.setFrozenColumns(2);
}

function rosterGuard_(code, pin, key, group) {
  const u = auth_(code, pin);
  if (!u || !u.admin) return { err: 'ไม่มีสิทธิ์แอดมิน' };
  const t = ymOf_(key);
  if (!t) return { err: 'เดือนไม่ถูกต้อง' };
  if (['RN', 'PN/NA'].indexOf(group) < 0) return { err: 'กลุ่มไม่ถูกต้อง' };
  return { y: t.y, mm: t.mm };
}

/** แอดมินกด "จัดเวร": สร้างตารางร่างลงแท็บใหม่ในชีต (เขียนทับของเดิมของเดือน/กลุ่มนี้) */
function generateRoster(code, pin, key, group) {
  const g = rosterGuard_(code, pin, key, group);
  if (g.err) return { ok: false, msg: g.err };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const P = rosterInput_(key, g.y, g.mm, group);
    const R = buildRoster_(P);
    const A = analyze_(R.grid, P);
    const alerts = R.alerts.concat(A.alerts);
    const name = rosterSheetName_(group, key);
    writeRoster_(P, R.grid, A, alerts, name);
    return { ok: true, sheet: name, stats: A.stats, alerts: alerts };
  } finally {
    lock.releaseLock();
  }
}

/** แอดมินกด "ตรวจเวรจากชีต": อ่านตารางที่แก้เองแล้ว คำนวณชั่วโมง/OT และเตือนทำงานติดกันเกิน 7 วัน */
function checkRoster(code, pin, key, group) {
  const g = rosterGuard_(code, pin, key, group);
  if (g.err) return { ok: false, msg: g.err };
  const P = rosterInput_(key, g.y, g.mm, group);
  const name = rosterSheetName_(group, key);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) return { ok: false, msg: 'ยังไม่มีแท็บเวรของเดือนนี้ – กด "จัดเวรเดือนนี้" ก่อน' };
  const n = daysIn_(g.y, g.mm);
  const vals = sh.getRange(4, 1, Math.max(sh.getLastRow() - 3, 1), 3 + n).getValues();
  const grid = {};
  vals.forEach(r => {
    const c = String(r[0]).trim();
    if (!c) return;
    grid[c] = {};
    for (let d = 1; d <= n; d++) grid[c][d] = String(r[2 + d]).trim();
  });
  P.staff = P.staff.filter(s => grid[s.code]);
  const A = analyze_(grid, P);
  return { ok: true, checked: true, sheet: name, stats: A.stats, alerts: A.alerts };
}

/** ส่งข้อความถึงแอดมิน (ไปลงแท็บ Notices) */
function notifyAdmin(code, pin, text) {
  const u = auth_(code, pin);
  if (!u) return { ok: false, msg: 'ยืนยันตัวตนไม่ผ่าน' };
  const m = months_();
  sh_('Notices').appendRow([new Date(), u.code, u.name, m.next, String(text || '').slice(0, 300), 'ใหม่']);
  return { ok: true };
}
