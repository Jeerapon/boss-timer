/**
 * Boss Timer v2 – Google Apps Script backend
 * -------------------------------------------
 * วิธีอัปเดตจากเวอร์ชันเดิม:
 *  1) ลบโค้ดเดิมทั้งหมดใน Apps Script แล้ววางไฟล์นี้ทั้งไฟล์ > Ctrl+S
 *  2) เลือกฟังก์ชัน setup > กด "เรียกใช้" (จะสร้างชีต Bosses / Events และแปลงชีตเดิมให้อัตโนมัติ
 *     ข้อมูลเดิมไม่หาย)
 *  3) การทำให้ใช้งานได้ > จัดการการทำให้ใช้งานได้ > ✏️ > เวอร์ชัน: "เวอร์ชันใหม่" > ทำให้ใช้งานได้
 *     (URL เดิมใช้ต่อได้เลย)
 */

// รหัสลับ (ไม่บังคับ) ถ้าใส่ ต้องใส่ค่าเดียวกันในหน้า "ตั้งค่า" ของโปรแกรมด้วย
// แนะนำให้ตั้ง เพราะเวอร์ชันนี้มีระบบแอดมินที่แก้ไข/ลบข้อมูลได้
const SECRET_KEY = '';

// บอส/กิจกรรมเดียวกัน ที่เวลาห่างกันไม่เกินกี่นาที ให้ถือว่าเป็น "รอบเดียวกัน"
const DUP_WINDOW_MIN = 10;

const SH = {
  members: 'Members',
  bosses: 'Bosses',
  events: 'Events',
  att: 'Attendance',
  sum: 'Summary'
};

const HEAD = {
  members: ['Name', 'Status', 'Joined Date', 'Note'],
  bosses: ['Boss Name', 'Points'],
  events: ['Event Name', 'Points'],
  att: ['Session ID', 'Recorded At', 'Time', 'Boss / Event', 'Points',
        'Member', 'Contribution', 'Recorded By', 'Type']
};

// บอสตั้งต้น (ใส่ให้เฉพาะตอนสร้างชีต Bosses ครั้งแรก แก้ไขได้ในหน้าแอดมิน)
const DEFAULT_BOSSES = [
  ['Kelsus', 2], ['Tromba', 1], ['INV_Gahareth', 2], ['Hisilrome', 2], ['Selu', 2],
  ['Chertuba', 1], ['INV_Matura', 1], ['INV_Hisilrome', 2], ['Pan Narod', 1]
];

/* ======================= SETUP ======================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ต้องสร้างสคริปต์นี้จากเมนู ส่วนขยาย > Apps Script ภายใน Google Sheet');

  // ---- Members ----
  let m = ss.getSheetByName(SH.members);
  if (!m) {
    m = ss.insertSheet(SH.members);
    writeHead_(m, HEAD.members);
  } else if (String(m.getRange(1, 2).getValue()).trim() === 'Note') {
    // แปลงจากเวอร์ชัน 1 (Name, Note) -> (Name, Status, Joined Date, Note)
    m.insertColumnsAfter(1, 2);
    writeHead_(m, HEAD.members);
    const n = m.getLastRow() - 1;
    if (n > 0) m.getRange(2, 2, n, 1).setValue('Active');
  }
  m.getRange('C:C').setNumberFormat('yyyy-mm-dd');
  m.getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Inactive'], true).build());

  // ---- Bosses ----
  let b = ss.getSheetByName(SH.bosses);
  if (!b) {
    b = ss.insertSheet(SH.bosses);
    writeHead_(b, HEAD.bosses);
    b.getRange(2, 1, DEFAULT_BOSSES.length, 2).setValues(DEFAULT_BOSSES);
  }

  // ---- Events ----
  let e = ss.getSheetByName(SH.events);
  if (!e) {
    e = ss.insertSheet(SH.events);
    writeHead_(e, HEAD.events);
  }

  // ---- Attendance ----
  let a = ss.getSheetByName(SH.att);
  if (!a) {
    a = ss.insertSheet(SH.att);
    writeHead_(a, HEAD.att);
  } else if (String(a.getRange(1, 9).getValue()).trim() !== 'Type') {
    // แปลงจากเวอร์ชัน 1: เพิ่มคอลัมน์ Type ให้ข้อมูลเดิมเป็น Boss
    writeHead_(a, HEAD.att);
    const n = a.getLastRow() - 1;
    if (n > 0) a.getRange(2, 9, n, 1).setValue('Boss');
  }
  a.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm');
  a.getRange('G:G').setNumberFormat('#,##0');

  // ---- Summary ----
  let s = ss.getSheetByName(SH.sum);
  if (!s) s = ss.insertSheet(SH.sum);
  s.getRange('A1').setFormula(
    "=QUERY(" + SH.att + "!A:I," +
    "\"select F, sum(E), count(F), sum(G) where F is not null group by F " +
    "order by sum(E) desc label F 'Member', sum(E) 'Total Points', " +
    "count(F) 'Attended', sum(G) 'Total Contribution'\",1)");
}

function writeHead_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/* ======================= HTTP ======================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return json_({ ok: false, error: 'รหัสลับไม่ถูกต้อง' });
  try {
    switch (p.action) {
      case 'bootstrap':
        return json_({ ok: true, members: getMembers_(), bosses: getCatalog_('bosses'), events: getCatalog_('events') });
      case 'members':
        return json_({ ok: true, members: getMembers_() });
      case 'history':
        return json_({ ok: true, history: getHistory_(Number(p.limit) || 50, p.type) });
      case 'recap':
        return json_(Object.assign({ ok: true }, getRecap_(p.type, p.from, p.to)));
      default:
        return json_({ ok: true, message: 'Boss Timer API v2 พร้อมใช้งาน' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'ข้อมูลที่ส่งมาไม่ใช่ JSON' }); }
  if (!checkKey_(body.key)) return json_({ ok: false, error: 'รหัสลับไม่ถูกต้อง' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'attendance':        return json_(saveAttendance_(body));
      case 'addToSession':      return json_(addToSession_(body));
      case 'removeFromSession': return json_(removeFromSession_(body));
      case 'deleteSession':     return json_(deleteSession_(body));
      case 'saveMember':        return json_(saveMember_(body));
      case 'deleteMember':      return json_(deleteMember_(body));
      case 'addMembers':        return json_(addMembers_(body.names || []));
      case 'saveCatalog':       return json_(saveCatalog_(body));
      case 'deleteCatalog':     return json_(deleteCatalog_(body));
      default: return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ======================= MEMBERS ======================= */

function getMembers_() {
  return rows_('members')
    .filter(function (r) { return String(r.v[0]).trim(); })
    .map(function (r) {
      return {
        name: String(r.v[0]).trim(),
        status: String(r.v[1]).trim() === 'Inactive' ? 'Inactive' : 'Active',
        joined: fmtDate_(r.v[2]),
        note: String(r.v[3] || '')
      };
    })
    .sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
}

function saveMember_(b) {
  const name = String(b.name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อสมาชิก');
  const status = b.status === 'Inactive' ? 'Inactive' : 'Active';
  const joined = parseDate_(b.joined) || new Date();
  const values = [[name, status, joined, String(b.note || '')]];
  const sheet = sheet_('members');
  const dup = findRow_('members', name);

  if (b.original) {
    const row = findRow_('members', b.original);
    if (!row) throw new Error('ไม่พบสมาชิก ' + b.original);
    if (dup && dup.row !== row.row) throw new Error('มีชื่อ ' + name + ' อยู่แล้ว');
    sheet.getRange(row.row, 1, 1, 4).setValues(values);
    if (String(b.original).trim() !== name) renameInAttendance_(b.original, name);
  } else {
    if (dup) throw new Error('มีชื่อ ' + name + ' อยู่แล้ว');
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues(values);
  }
  return { ok: true, members: getMembers_() };
}

/** เปลี่ยนชื่อในประวัติทั้งหมดด้วย เมื่อแก้ชื่อสมาชิก */
function renameInAttendance_(from, to) {
  const s = sheet_('att');
  const last = s.getLastRow();
  if (last < 2) return;
  const rng = s.getRange(2, 6, last - 1, 1);
  const vals = rng.getValues();
  const f = String(from).trim().toLowerCase();
  let changed = false;
  vals.forEach(function (r) {
    if (String(r[0]).trim().toLowerCase() === f) { r[0] = to; changed = true; }
  });
  if (changed) rng.setValues(vals);
}

function deleteMember_(b) {
  const row = findRow_('members', b.name);
  if (!row) throw new Error('ไม่พบสมาชิก ' + b.name);
  sheet_('members').deleteRow(row.row);
  return { ok: true, members: getMembers_() };
}

function addMembers_(names) {
  const sheet = sheet_('members');
  const existing = getMembers_().map(function (m) { return m.name.toLowerCase(); });
  const today = new Date();
  const toAdd = [];
  names.forEach(function (n) {
    n = String(n || '').trim();
    if (n && existing.indexOf(n.toLowerCase()) === -1) {
      existing.push(n.toLowerCase());
      toAdd.push([n, 'Active', today, '']);
    }
  });
  if (toAdd.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 4).setValues(toAdd);
  return { ok: true, added: toAdd.length, members: getMembers_() };
}

/* ======================= BOSSES / EVENTS ======================= */

function getCatalog_(kind) {
  return rows_(kind)
    .filter(function (r) { return String(r.v[0]).trim(); })
    .map(function (r) { return { name: String(r.v[0]).trim(), points: Number(r.v[1]) || 0 }; });
}

function saveCatalog_(b) {
  const kind = b.kind === 'events' ? 'events' : 'bosses';
  const label = kind === 'events' ? 'กิจกรรม' : 'บอส';
  const name = String(b.name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อ' + label);
  const values = [[name, Number(b.points) || 0]];
  const sheet = sheet_(kind);
  const dup = findRow_(kind, name);

  if (b.original) {
    const row = findRow_(kind, b.original);
    if (!row) throw new Error('ไม่พบ' + label + ' ' + b.original);
    if (dup && dup.row !== row.row) throw new Error('มี' + label + 'ชื่อ ' + name + ' อยู่แล้ว');
    sheet.getRange(row.row, 1, 1, 2).setValues(values);
  } else {
    if (dup) throw new Error('มี' + label + 'ชื่อ ' + name + ' อยู่แล้ว');
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues(values);
  }
  return { ok: true, kind: kind, list: getCatalog_(kind) };
}

function deleteCatalog_(b) {
  const kind = b.kind === 'events' ? 'events' : 'bosses';
  const row = findRow_(kind, b.name);
  if (!row) throw new Error('ไม่พบ ' + b.name);
  sheet_(kind).deleteRow(row.row);
  return { ok: true, kind: kind, list: getCatalog_(kind) };
}

/* ======================= ATTENDANCE ======================= */

function saveAttendance_(b) {
  const list = (b.participants || []).filter(function (p) { return p && p.name; });
  const name = String(b.name || b.boss || '').trim();
  if (!name) return { ok: false, error: 'ยังไม่ได้เลือกบอส/กิจกรรม' };
  if (!list.length) return { ok: false, error: 'ยังไม่ได้เลือกสมาชิก' };

  const sessionId = b.sessionId || Utilities.getUuid().slice(0, 8);
  if (sessionRows_(sessionId).length) {
    return { ok: false, error: 'รอบนี้ถูกบันทึกไปแล้ว (Session ' + sessionId + ')' };
  }

  const now = new Date();
  const time = b.killTime ? new Date(b.killTime) : now;
  const type = b.type === 'event' ? 'Event' : 'Boss';
  const mode = b.dupMode || '';   // '' = ตรวจซ้ำก่อน, 'merge' | 'replace' | 'new'

  // ---- ตรวจรอบซ้ำ: ชื่อเดียวกัน + ประเภทเดียวกัน + เวลาห่างกันไม่เกิน DUP_WINDOW_MIN นาที ----
  if (!mode) {
    const dup = findDuplicate_(type, name, time);
    if (dup) return { ok: false, duplicate: dup, error: 'พบรอบที่บันทึกไว้แล้วในช่วงเวลาใกล้กัน' };
  }

  if (mode === 'merge') return mergeIntoSession_(b.targetSessionId, list, b.recordedBy);

  let replaced = 0;
  if (mode === 'replace') {
    const old = sessionRows_(b.targetSessionId);
    if (!old.length) throw new Error('ไม่พบรอบเดิมที่จะแทนที่');
    deleteRows_(old);
    replaced = old.length;
  }

  const rows = list.map(function (p) {
    return [sessionId, now, time, name, Number(b.points) || 0,
            String(p.name), Number(p.contribution) || '', String(b.recordedBy || ''), type];
  });
  appendAtt_(rows);
  return { ok: true, sessionId: sessionId, saved: rows.length, replaced: replaced };
}

/** หารอบที่บันทึกไว้แล้ว ของบอส/กิจกรรมเดียวกันในช่วงเวลาใกล้กัน */
function findDuplicate_(type, name, time) {
  const win = DUP_WINDOW_MIN * 60 * 1000;
  const n = String(name).toLowerCase();
  const sessions = {};
  rows_('att').forEach(function (r) {
    const v = r.v;
    if (!v[0] || String(v[3]).toLowerCase() !== n) return;
    if (String(v[8] || 'Boss') !== type) return;
    const t = v[2] instanceof Date ? v[2].getTime() : new Date(v[2]).getTime();
    if (isNaN(t) || Math.abs(t - time.getTime()) > win) return;
    const id = String(v[0]);
    if (!sessions[id]) sessions[id] = { sessionId: id, time: toIso_(v[2]), name: String(v[3]), points: Number(v[4]) || 0,
                                        recordedBy: String(v[7]).replace(/ \((เพิ่มภายหลัง|รวมรอบ)\)$/, ''), members: [], diff: Math.abs(t - time.getTime()) };
    sessions[id].members.push(String(v[5]));
  });
  const list = Object.keys(sessions).map(function (k) { return sessions[k]; })
    .sort(function (a, b) { return a.diff - b.diff; });
  return list.length ? list[0] : null;
}

/** รวมรายชื่อเข้ารอบเดิม: เพิ่มเฉพาะคนที่ยังไม่มี / อัปเดต Contribution ถ้าค่าใหม่มากกว่า */
function mergeIntoSession_(targetId, list, recordedBy) {
  const rows = sessionRows_(targetId);
  if (!rows.length) throw new Error('ไม่พบรอบเดิมที่จะรวม');
  const s = sheet_('att');
  const t = rows[0].v;
  const byName = {};
  rows.forEach(function (r) { byName[String(r.v[5]).toLowerCase()] = r; });

  let added = 0, updated = 0;
  const out = [];
  list.forEach(function (p) {
    const key = String(p.name).toLowerCase();
    const c = Number(p.contribution) || 0;
    const ex = byName[key];
    if (ex) {
      if (c > (Number(ex.v[6]) || 0)) { s.getRange(ex.row, 7).setValue(c); updated++; }
    } else {
      byName[key] = true;
      out.push([t[0], new Date(), t[2], t[3], t[4], String(p.name), c || '',
                String(recordedBy || '') + ' (รวมรอบ)', t[8] || 'Boss']);
      added++;
    }
  });
  if (out.length) appendAtt_(out);
  return { ok: true, merged: true, sessionId: targetId, added: added, updated: updated };
}

/** เพิ่มรายชื่อที่ตกหล่นเข้าไปในรอบที่บันทึกแล้ว */
function addToSession_(b) {
  const rows = sessionRows_(b.sessionId);
  if (!rows.length) throw new Error('ไม่พบรอบที่ต้องการ');
  const t = rows[0].v;
  const have = rows.map(function (r) { return String(r.v[5]).toLowerCase(); });
  const recorder = String(b.recordedBy || '') + ' (เพิ่มภายหลัง)';
  const out = [];
  (b.names || []).forEach(function (n) {
    n = String(n || '').trim();
    if (n && have.indexOf(n.toLowerCase()) === -1) {
      have.push(n.toLowerCase());
      out.push([t[0], new Date(), t[2], t[3], t[4], n, '', recorder, t[8] || 'Boss']);
    }
  });
  if (out.length) appendAtt_(out);
  return { ok: true, added: out.length };
}

function removeFromSession_(b) {
  const target = String(b.name || '').toLowerCase();
  const rows = sessionRows_(b.sessionId).filter(function (r) {
    return String(r.v[5]).toLowerCase() === target;
  });
  if (!rows.length) throw new Error('ไม่พบ ' + b.name + ' ในรอบนี้');
  deleteRows_(rows);
  return { ok: true, removed: rows.length };
}

function deleteSession_(b) {
  const rows = sessionRows_(b.sessionId);
  if (!rows.length) throw new Error('ไม่พบรอบที่ต้องการ');
  deleteRows_(rows);
  return { ok: true, removed: rows.length };
}

function getHistory_(limit, type) {
  const sessions = {}, order = [];
  rows_('att').forEach(function (r) {
    const v = r.v, id = String(v[0]);
    if (!id) return;
    const t = String(v[8] || 'Boss');
    if (type && type !== 'all' && t.toLowerCase() !== type) return;
    if (!sessions[id]) {
      sessions[id] = {
        sessionId: id, recordedAt: toIso_(v[1]), time: toIso_(v[2]), name: String(v[3]),
        points: Number(v[4]) || 0, recordedBy: String(v[7]).replace(/ \((เพิ่มภายหลัง|รวมรอบ)\)$/, ''),
        type: t, members: []
      };
      order.push(id);
    }
    sessions[id].members.push({ name: String(v[5]), contribution: Number(v[6]) || 0 });
  });
  return order.reverse().slice(0, limit).map(function (id) { return sessions[id]; });
}

/**
 * สรุปคะแนนตามช่วงเวลา
 * from / to = 'yyyy-MM-dd' (ตามเขตเวลาของชีต) เว้นว่าง = ไม่จำกัด
 * คืนค่า recap (รายคน), days (รายวัน), sessions (รายรอบ)
 */
function getRecap_(type, from, to) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const status = {};
  getMembers_().forEach(function (m) { status[m.name.toLowerCase()] = m.status; });

  const map = {}, days = {}, sessions = {};
  rows_('att').forEach(function (r) {
    const v = r.v;
    if (!v[5]) return;
    const t = String(v[8] || 'Boss').toLowerCase() === 'event' ? 'event' : 'boss';
    if (type && type !== 'all' && t !== type) return;
    const tm = v[2] instanceof Date ? v[2] : new Date(v[2]);
    if (isNaN(tm.getTime())) return;
    const day = Utilities.formatDate(tm, tz, 'yyyy-MM-dd');
    if (from && day < from) return;
    if (to && day > to) return;

    const pts = Number(v[4]) || 0, dmg = Number(v[6]) || 0, id = String(v[0]);

    // รายคน
    const key = String(v[5]).toLowerCase();
    if (!map[key]) map[key] = { name: String(v[5]), points: 0, count: 0, boss: 0, event: 0, contribution: 0, _d: {} };
    const x = map[key];
    x.points += pts; x.count += 1; x[t] += 1; x.contribution += dmg; x._d[day] = 1;

    // รายวัน
    if (!days[day]) days[day] = { date: day, attend: 0, bossAttend: 0, eventAttend: 0, points: 0, _s: {} };
    const d = days[day];
    d.attend += 1; d[t + 'Attend'] += 1; d.points += pts; d._s[id] = t;

    // รายรอบ
    if (!sessions[id]) sessions[id] = { sessionId: id, time: toIso_(tm), day: day, name: String(v[3]), type: t, points: pts, count: 0, contribution: 0 };
    sessions[id].count += 1; sessions[id].contribution += dmg;
  });

  const recap = Object.keys(map).map(function (k) {
    const x = map[k];
    x.days = Object.keys(x._d).length; delete x._d;
    x.status = status[k] || 'Unknown';
    return x;
  }).sort(function (a, b) { return b.points - a.points || b.count - a.count || b.contribution - a.contribution; });

  const dayList = Object.keys(days).sort().map(function (k) {
    const d = days[k], ids = Object.keys(d._s);
    d.sessions = ids.length;
    d.bossSessions = ids.filter(function (i) { return d._s[i] === 'boss'; }).length;
    d.eventSessions = d.sessions - d.bossSessions;
    delete d._s;
    return d;
  });

  const sessionList = Object.keys(sessions).map(function (k) { return sessions[k]; })
    .sort(function (a, b) { return a.time < b.time ? 1 : -1; }).slice(0, 300);

  return { recap: recap, days: dayList, sessions: sessionList };
}

/* ======================= HELPERS ======================= */

function sheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SH[key]);
  if (!s) { setup(); s = ss.getSheetByName(SH[key]); }
  return s;
}

function rows_(key) {
  const s = sheet_(key);
  const last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, HEAD[key].length).getValues()
    .map(function (v, i) { return { row: i + 2, v: v }; });
}

function findRow_(key, name) {
  const n = String(name || '').trim().toLowerCase();
  if (!n) return null;
  const all = rows_(key);
  for (let i = 0; i < all.length; i++) {
    if (String(all[i].v[0]).trim().toLowerCase() === n) return all[i];
  }
  return null;
}

function sessionRows_(id) {
  if (!id) return [];
  return rows_('att').filter(function (r) { return String(r.v[0]) === String(id); });
}

function appendAtt_(rows) {
  const s = sheet_('att');
  s.getRange(s.getLastRow() + 1, 1, rows.length, HEAD.att.length).setValues(rows);
}

function deleteRows_(rows) {
  const s = sheet_('att');
  rows.map(function (r) { return r.row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (r) { s.deleteRow(r); });
}

function parseDate_(v) {
  if (!v) return null;
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function fmtDate_(v) {
  if (!(v instanceof Date)) return String(v || '');
  return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

function checkKey_(key) { return !SECRET_KEY || key === SECRET_KEY; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
