/**
 * Boss Timer v2 – Google Apps Script backend
 * -------------------------------------------
 * How to update:
 *  1) Delete all old code in Apps Script, paste this whole file > Ctrl+S
 *  2) Select function "setup" > Run (creates/updates sheets automatically, existing data is kept)
 *  3) Deploy > Manage deployments > ✏️ > Version: "New version" > Deploy
 *     (the Web app URL stays the same)
 */

// Secret key (optional). If set, enter the same key in the app's Settings page.
// Recommended, because the Admin tools can edit/delete data.
const SECRET_KEY = '';

// Same boss/event within this many minutes = treated as the same run (duplicate check)
const DUP_WINDOW_MIN = 10;

const SH = {
  members: 'Members',
  bosses: 'Bosses',
  events: 'Events',
  att: 'Attendance',
  sum: 'Summary',
  audit: 'AuditLog'
};

const HEAD = {
  members: ['Name', 'Status', 'Joined Date', 'Note', 'Alias'],
  bosses: ['Boss Name', 'Points'],
  events: ['Event Name', 'Points'],
  att: ['Session ID', 'Recorded At', 'Time', 'Boss / Event', 'Points',
        'Member', 'Contribution', 'Recorded By', 'Type'],
  audit: ['Time', 'Action', 'By', 'Device', 'Category', 'Target', 'Session ID', 'Detail']
};

// Starter bosses (only added when the Bosses sheet is first created; edit in Admin)
const DEFAULT_BOSSES = [
  ['Kelsus', 2], ['Tromba', 1], ['INV_Gahareth', 2], ['Hisilrome', 2], ['Selu', 2],
  ['Chertuba', 1], ['INV_Matura', 1], ['INV_Hisilrome', 2], ['Pan Narod', 1]
];

/* ======================= SETUP ======================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('This script must be created from inside a Google Sheet: Extensions > Apps Script');

  // ---- Members ----
  let m = ss.getSheetByName(SH.members);
  if (!m) {
    m = ss.insertSheet(SH.members);
    writeHead_(m, HEAD.members);
  } else if (String(m.getRange(1, 2).getValue()).trim() === 'Note') {
    // Upgrade from v1 (Name, Note) -> (Name, Status, Joined Date, Note)
    m.insertColumnsAfter(1, 2);
    writeHead_(m, HEAD.members);
    const n = m.getLastRow() - 1;
    if (n > 0) m.getRange(2, 2, n, 1).setValue('Active');
  }
  if (String(m.getRange(1, 5).getValue()).trim() !== 'Alias') writeHead_(m, HEAD.members);
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
    // Upgrade from v1: add Type column, old rows = Boss
    writeHead_(a, HEAD.att);
    const n = a.getLastRow() - 1;
    if (n > 0) a.getRange(2, 9, n, 1).setValue('Boss');
  }
  a.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm');
  a.getRange('G:G').setNumberFormat('#,##0');

  // ---- AuditLog (who added/removed/edited what) ----
  let lg = ss.getSheetByName(SH.audit);
  if (!lg) {
    lg = ss.insertSheet(SH.audit);
    writeHead_(lg, HEAD.audit);
    lg.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    lg.setColumnWidth(8, 520);
  }

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
  if (!checkKey_(p.key)) return json_({ ok: false, error: 'Wrong secret key' });
  try {
    switch (p.action) {
      case 'bootstrap':
        return json_({ ok: true, members: getMembers_(), bosses: getCatalog_('bosses'), events: getCatalog_('events') });
      case 'members':
        return json_({ ok: true, members: getMembers_() });
      case 'history':
        return json_({ ok: true, history: getHistory_(Number(p.limit) || 50, p.type) });
      case 'audit':
        return json_({ ok: true, log: getAudit_(p.sessionId, Number(p.limit) || 300, p.category) });
      case 'recap':
        return json_(Object.assign({ ok: true }, getRecap_(p.type, p.from, p.to)));
      default:
        return json_({ ok: true, message: 'Boss Timer API v2 is ready' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Request is not valid JSON' }); }
  if (!checkKey_(body.key)) return json_({ ok: false, error: 'Wrong secret key' });

  // OCR runs without the lock (many people can scan at once)
  if (body.action === 'ocr') {
    try { return json_(googleOcr_(body)); }
    catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const before = auditBefore_(body);   // snapshot before changes, for the audit log
    let res;
    switch (body.action) {
      case 'attendance':        res = saveAttendance_(body); break;
      case 'addToSession':      res = addToSession_(body); break;
      case 'removeFromSession': res = removeFromSession_(body); break;
      case 'deleteSession':     res = deleteSession_(body); break;
      case 'saveMember':        res = saveMember_(body); break;
      case 'deleteMember':      res = deleteMember_(body); break;
      case 'addAlias':          res = addAlias_(body); break;
      case 'addMembers':        res = addMembers_(body.names || []); break;
      case 'saveCatalog':       res = saveCatalog_(body); break;
      case 'deleteCatalog':     res = deleteCatalog_(body); break;
      default: return json_({ ok: false, error: 'Unknown action: ' + body.action });
    }
    if (res && res.ok) { try { audit_(body, before, res); } catch (e) { console.error('audit failed', e); } }
    return json_(res);
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
        note: String(r.v[3] || ''),
        alias: String(r.v[4] || '')
      };
    })
    .sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
}

function saveMember_(b) {
  const name = String(b.name || '').trim();
  if (!name) throw new Error('Please enter a member name');
  const status = b.status === 'Inactive' ? 'Inactive' : 'Active';
  const joined = parseDate_(b.joined) || new Date();
  const values = [[name, status, joined, String(b.note || ''), normAlias_(b.alias)]];
  const sheet = sheet_('members');
  const dup = findRow_('members', name);

  if (b.original) {
    const row = findRow_('members', b.original);
    if (!row) throw new Error('Member not found: ' + b.original);
    if (dup && dup.row !== row.row) throw new Error(name + ' already exists');
    sheet.getRange(row.row, 1, 1, 5).setValues(values);
    if (String(b.original).trim() !== name) renameInAttendance_(b.original, name);
  } else {
    if (dup) throw new Error(name + ' already exists');
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 5).setValues(values);
  }
  return { ok: true, members: getMembers_() };
}

/** Rename the member in all past runs too */
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
  if (!row) throw new Error('Member not found: ' + b.name);
  sheet_('members').deleteRow(row.row);
  return { ok: true, members: getMembers_() };
}

/** Save an OCR misread as an alias of a member (teaches the scanner) */
function addAlias_(b) {
  const row = findRow_('members', b.name);
  if (!row) throw new Error('Member not found: ' + b.name);
  const cur = normAlias_(row.v[4]).split(', ').filter(String);
  const add = String(b.alias || '').trim();
  if (!add) throw new Error('Nothing to add');
  if (cur.map(function (x) { return x.toLowerCase(); }).indexOf(add.toLowerCase()) === -1) cur.push(add);
  sheet_('members').getRange(row.row, 5).setValue(cur.join(', '));
  return { ok: true, members: getMembers_() };
}

function normAlias_(v) {
  return String(v || '').split(/[,\n]/).map(function (x) { return x.trim(); })
    .filter(String).join(', ');
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
      toAdd.push([n, 'Active', today, '', '']);
    }
  });
  if (toAdd.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 5).setValues(toAdd);
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
  const label = kind === 'events' ? 'Event' : 'Boss';
  const name = String(b.name || '').trim();
  if (!name) throw new Error('Please enter a ' + label.toLowerCase() + ' name');
  const values = [[name, Number(b.points) || 0]];
  const sheet = sheet_(kind);
  const dup = findRow_(kind, name);

  if (b.original) {
    const row = findRow_(kind, b.original);
    if (!row) throw new Error(label + ' not found: ' + b.original);
    if (dup && dup.row !== row.row) throw new Error(label + ' "' + name + '" already exists');
    sheet.getRange(row.row, 1, 1, 2).setValues(values);
  } else {
    if (dup) throw new Error(label + ' "' + name + '" already exists');
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues(values);
  }
  return { ok: true, kind: kind, list: getCatalog_(kind) };
}

function deleteCatalog_(b) {
  const kind = b.kind === 'events' ? 'events' : 'bosses';
  const row = findRow_(kind, b.name);
  if (!row) throw new Error('Not found: ' + b.name);
  sheet_(kind).deleteRow(row.row);
  return { ok: true, kind: kind, list: getCatalog_(kind) };
}

/* ======================= ATTENDANCE ======================= */

function saveAttendance_(b) {
  const list = (b.participants || []).filter(function (p) { return p && p.name; });
  const name = String(b.name || b.boss || '').trim();
  if (!name) return { ok: false, error: 'Pick a boss/event first' };
  if (!list.length) return { ok: false, error: 'No members selected' };

  const sessionId = b.sessionId || Utilities.getUuid().slice(0, 8);
  if (sessionRows_(sessionId).length) {
    return { ok: false, error: 'This run was already saved (Session ' + sessionId + ')' };
  }

  const now = new Date();
  const time = b.killTime ? new Date(b.killTime) : now;
  const type = b.type === 'event' ? 'Event' : 'Boss';
  const mode = b.dupMode || '';   // '' = check duplicates first, 'merge' | 'replace' | 'new'

  // ---- Duplicate check: same name + same type + within DUP_WINDOW_MIN minutes ----
  if (!mode) {
    const dup = findDuplicate_(type, name, time);
    if (dup) return { ok: false, duplicate: dup, error: 'A run was already saved around this time' };
  }

  if (mode === 'merge') return mergeIntoSession_(b.targetSessionId, list, b.recordedBy);

  let replaced = 0;
  if (mode === 'replace') {
    const old = sessionRows_(b.targetSessionId);
    if (!old.length) throw new Error('Original run to replace not found');
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

/** Find an existing run of the same boss/event close in time */
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
                                        recordedBy: String(v[7]).replace(/ \((เพิ่มภายหลัง|รวมรอบ|added later|merged)\)$/, ''), members: [], diff: Math.abs(t - time.getTime()) };
    sessions[id].members.push(String(v[5]));
  });
  const list = Object.keys(sessions).map(function (k) { return sessions[k]; })
    .sort(function (a, b) { return a.diff - b.diff; });
  return list.length ? list[0] : null;
}

/** Merge into existing run: add only new people / update Contribution if higher */
function mergeIntoSession_(targetId, list, recordedBy) {
  const rows = sessionRows_(targetId);
  if (!rows.length) throw new Error('Original run to merge into not found');
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
                String(recordedBy || '') + ' (merged)', t[8] || 'Boss']);
      added++;
    }
  });
  if (out.length) appendAtt_(out);
  return { ok: true, merged: true, sessionId: targetId, added: added, updated: updated };
}

/** Add missed members to a saved run */
function addToSession_(b) {
  const rows = sessionRows_(b.sessionId);
  if (!rows.length) throw new Error('Run not found');
  const t = rows[0].v;
  const have = rows.map(function (r) { return String(r.v[5]).toLowerCase(); });
  const recorder = String(b.recordedBy || '') + ' (added later)';
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
  if (!rows.length) throw new Error(b.name + ' is not in this run');
  deleteRows_(rows);
  return { ok: true, removed: rows.length };
}

function deleteSession_(b) {
  const rows = sessionRows_(b.sessionId);
  if (!rows.length) throw new Error('Run not found');
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
        points: Number(v[4]) || 0, recordedBy: String(v[7]).replace(/ \((เพิ่มภายหลัง|รวมรอบ|added later|merged)\)$/, ''),
        type: t, members: []
      };
      order.push(id);
    }
    const by = String(v[7]);
    const late = / \((เพิ่มภายหลัง|รวมรอบ|added later|merged)\)$/.test(by);
    sessions[id].members.push({ name: String(v[5]), contribution: Number(v[6]) || 0,
      late: late, by: late ? by.replace(/ \((เพิ่มภายหลัง|รวมรอบ|added later|merged)\)$/, '') : '', at: late ? toIso_(v[1]) : '' });
  });
  const out = order.reverse().slice(0, limit).map(function (id) { return sessions[id]; });
  // How many times each run was edited after saving (from AuditLog; also matches old Thai labels)
  const edits = {};
  rows_('audit').forEach(function (r) {
    const sid = String(r.v[6]);
    if (sid && ['บันทึกรอบใหม่', 'New run'].indexOf(String(r.v[1])) === -1) edits[sid] = (edits[sid] || 0) + 1;
  });
  out.forEach(function (x) { x.edits = edits[x.sessionId] || 0; });
  return out;
}

/**
 * Leaderboard for a date range
 * from / to = 'yyyy-MM-dd' (sheet time zone), empty = no limit
 * Returns recap (per member), days (per day), sessions (per run)
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

    // per member
    const key = String(v[5]).toLowerCase();
    if (!map[key]) map[key] = { name: String(v[5]), points: 0, count: 0, boss: 0, event: 0, contribution: 0, _d: {} };
    const x = map[key];
    x.points += pts; x.count += 1; x[t] += 1; x.contribution += dmg; x._d[day] = 1;

    // per day
    if (!days[day]) days[day] = { date: day, attend: 0, bossAttend: 0, eventAttend: 0, points: 0, _s: {} };
    const d = days[day];
    d.attend += 1; d[t + 'Attend'] += 1; d.points += pts; d._s[id] = t;

    // per run
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

/* ======================= AUDIT LOG ======================= */

function sessionInfo_(id) {
  const rows = sessionRows_(id);
  if (!rows.length) return null;
  const v = rows[0].v;
  return {
    name: String(v[3]), time: toIso_(v[2]), type: String(v[8] || 'Boss'), points: Number(v[4]) || 0,
    members: rows.map(function (r) { return { name: String(r.v[5]), contribution: Number(r.v[6]) || 0 }; })
  };
}

/** Snapshot data before a change, for the audit log (e.g. removed names, old values) */
function auditBefore_(b) {
  switch (b.action) {
    case 'addToSession': case 'removeFromSession': case 'deleteSession':
      return { session: sessionInfo_(b.sessionId) };
    case 'attendance':
      return b.targetSessionId ? { session: sessionInfo_(b.targetSessionId) } : {};
    case 'saveMember': case 'deleteMember': case 'addAlias': {
      const key = String(b.original || b.name || '').toLowerCase();
      return { member: getMembers_().filter(function (m) { return m.name.toLowerCase() === key; })[0] || null };
    }
    case 'addMembers':
      return { names: getMembers_().map(function (m) { return m.name.toLowerCase(); }) };
    case 'saveCatalog': case 'deleteCatalog': {
      const kind = b.kind === 'events' ? 'events' : 'bosses';
      const key = String(b.original || b.name || '').toLowerCase();
      return { item: getCatalog_(kind).filter(function (x) { return x.name.toLowerCase() === key; })[0] || null };
    }
  }
  return {};
}

function audit_(b, bf, r) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const when = function (iso) { return iso ? Utilities.formatDate(new Date(iso), tz, 'dd/MM HH:mm') : ''; };
  const list = function (arr) { return arr.map(function (m) { return m.contribution ? m.name + ' (' + m.contribution + ')' : m.name; }).join(', '); };
  const S = bf.session;
  let action = '', cat = '', target = '', sid = '', detail = '';

  switch (b.action) {
    case 'attendance': {
      cat = 'Run'; target = String(b.name || b.boss || '');
      const parts = (b.participants || []).filter(function (p) { return p && p.name; })
        .map(function (p) { return { name: String(p.name), contribution: Number(p.contribution) || 0 }; });
      if (r.merged) {
        action = 'Merged into run'; sid = r.sessionId;
        const have = (S ? S.members : []).map(function (m) { return m.name.toLowerCase(); });
        const added = parts.filter(function (p) { return have.indexOf(p.name.toLowerCase()) === -1; });
        detail = 'Added ' + r.added + ': ' + (list(added) || '-') + ' · DMG updated for ' + r.updated;
      } else if (r.replaced) {
        action = 'Replaced run'; sid = r.sessionId;
        detail = 'Old run ' + (S ? when(S.time) + ' ' + S.members.length + ' players: ' + list(S.members) : '') +
                 ' → New ' + parts.length + ' players: ' + list(parts);
      } else {
        action = 'New run'; sid = r.sessionId;
        detail = when(b.killTime) + ' · ' + (Number(b.points) || 0) + ' pts · ' + parts.length + ' players: ' + list(parts);
      }
      break;
    }
    case 'addToSession': {
      action = 'Added to run'; cat = 'Run'; sid = b.sessionId; target = S ? S.name : '';
      const have = (S ? S.members : []).map(function (m) { return m.name.toLowerCase(); });
      const added = (b.names || []).filter(function (n) { return n && have.indexOf(String(n).toLowerCase()) === -1; });
      detail = (S ? when(S.time) + ' · ' : '') + 'Added ' + added.length + ': ' + added.join(', ');
      break;
    }
    case 'removeFromSession': {
      action = 'Removed from run'; cat = 'Run'; sid = b.sessionId; target = S ? S.name : '';
      const m = S ? S.members.filter(function (x) { return x.name.toLowerCase() === String(b.name).toLowerCase(); })[0] : null;
      detail = (S ? when(S.time) + ' · ' : '') + 'Removed: ' + b.name + (m && m.contribution ? ' (DMG ' + m.contribution + ')' : '');
      break;
    }
    case 'deleteSession':
      action = 'Deleted run'; cat = 'Run'; sid = b.sessionId; target = S ? S.name : '';
      detail = S ? when(S.time) + ' · ' + S.points + ' pts · ' + S.members.length + ' players: ' + list(S.members) : '';
      break;
    case 'saveMember': {
      cat = 'Member'; target = String(b.name);
      const o = bf.member;
      if (b.original && o) {
        action = 'Edited member';
        const ch = [];
        if (o.name !== String(b.name).trim()) ch.push('Name: ' + o.name + ' → ' + b.name);
        if (o.status !== (b.status === 'Inactive' ? 'Inactive' : 'Active')) ch.push('Status: ' + o.status + ' → ' + b.status);
        if ((o.joined || '') !== (b.joined || o.joined)) ch.push('Joined: ' + o.joined + ' → ' + b.joined);
        if ((o.note || '') !== String(b.note || '')) ch.push('Note: ' + (o.note || '-') + ' → ' + (b.note || '-'));
        if (normAlias_(o.alias) !== normAlias_(b.alias)) ch.push('Alias: ' + (o.alias || '-') + ' → ' + (normAlias_(b.alias) || '-'));
        detail = ch.join(' · ') || 'No changes';
      } else {
        action = 'Added member';
        detail = 'Status ' + (b.status || 'Active') + ' · Joined ' + (b.joined || '') + (b.note ? ' · Note ' + b.note : '');
      }
      break;
    }
    case 'deleteMember': {
      action = 'Deleted member'; cat = 'Member'; target = String(b.name);
      const o = bf.member;
      detail = o ? 'Status ' + o.status + ' · Joined ' + o.joined + (o.note ? ' · Note ' + o.note : '') + (o.alias ? ' · Alias ' + o.alias : '') : '';
      break;
    }
    case 'addAlias':
      action = 'Added alias'; cat = 'Member'; target = String(b.name);
      detail = '"' + b.alias + '" → ' + b.name;
      break;
    case 'addMembers': {
      action = 'Added members (bulk)'; cat = 'Member';
      const had = bf.names || [];
      const added = (b.names || []).map(function (n) { return String(n).trim(); })
        .filter(function (n) { return n && had.indexOf(n.toLowerCase()) === -1; });
      if (!added.length) return;
      target = added.length + ' players'; detail = added.join(', ');
      break;
    }
    case 'saveCatalog': case 'deleteCatalog': {
      const label = b.kind === 'events' ? 'event' : 'boss';
      const Label = b.kind === 'events' ? 'Event' : 'Boss';
      cat = Label; target = String(b.name);
      const o = bf.item;
      if (b.action === 'deleteCatalog') { action = 'Deleted ' + label; detail = o ? o.points + ' pts' : ''; }
      else if (b.original && o) {
        action = 'Edited ' + label;
        const ch = [];
        if (o.name !== String(b.name).trim()) ch.push('Name: ' + o.name + ' → ' + b.name);
        if (o.points !== (Number(b.points) || 0)) ch.push('Points: ' + o.points + ' → ' + b.points);
        detail = ch.join(' · ') || 'No changes';
      } else { action = 'Added ' + label; detail = (Number(b.points) || 0) + ' pts'; }
      break;
    }
    default: return;
  }

  const s = sheet_('audit');
  s.getRange(s.getLastRow() + 1, 1, 1, HEAD.audit.length).setValues([[
    new Date(), action, String(b.recordedBy || '(no name)'), String(b.device || ''),
    cat, target, sid, detail
  ]]);
}

// Old Thai labels from earlier versions -> English (so old log rows display in English)
const OLD_LABELS = {
  'บันทึกรอบใหม่': 'New run', 'รวมเข้ารอบเดิม': 'Merged into run', 'แทนที่รอบ': 'Replaced run',
  'เพิ่มชื่อในรอบ': 'Added to run', 'เอาชื่อออกจากรอบ': 'Removed from run', 'ลบรอบ': 'Deleted run',
  'เพิ่มสมาชิก': 'Added member', 'แก้ไขสมาชิก': 'Edited member', 'ลบสมาชิก': 'Deleted member',
  'เพิ่ม Alias': 'Added alias', 'เพิ่มสมาชิก (หลายคน)': 'Added members (bulk)',
  'เพิ่มบอส': 'Added boss', 'แก้ไขบอส': 'Edited boss', 'ลบบอส': 'Deleted boss',
  'เพิ่มกิจกรรม': 'Added event', 'แก้ไขกิจกรรม': 'Edited event', 'ลบกิจกรรม': 'Deleted event',
  'รอบ': 'Run', 'สมาชิก': 'Member', 'บอส': 'Boss', 'กิจกรรม': 'Event'
};
const en_ = function (x) { x = String(x); return OLD_LABELS[x] || x; };

function getAudit_(sessionId, limit, category) {
  return rows_('audit')
    .filter(function (r) { return r.v[1] && (!sessionId || String(r.v[6]) === String(sessionId)) && (!category || en_(r.v[4]) === category); })
    .reverse().slice(0, limit)
    .map(function (r) {
      const v = r.v;
      return { time: toIso_(v[0]), action: en_(v[1]), by: String(v[2]), device: String(v[3]),
               category: en_(v[4]), target: String(v[5]), sessionId: String(v[6]), detail: String(v[7]) };
    });
}

/* ======================= GOOGLE OCR ======================= */
/**
 * Read text from an image with Google's free OCR (via Google Drive)
 * Requires the "Drive API" service: Apps Script left menu > Services > + > Drive API
 */
function googleOcr_(b) {
  if (typeof Drive === 'undefined') {
    throw new Error('Drive API is not enabled in Apps Script (Services > + > Drive API > Add), then Deploy a new version');
  }
  const b64 = String(b.image || '').replace(/^data:image\/\w+;base64,/, '');
  if (!b64) throw new Error('No image to read');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), b.mime || 'image/png', 'bt-ocr.png');
  const lang = String(b.lang || 'en');
  let id = null;
  try {
    if (Drive.Files.create) {          // Drive API v3
      const f = Drive.Files.create({ name: 'bt-ocr-temp', mimeType: 'application/vnd.google-apps.document' },
                                   blob, { ocrLanguage: lang });
      id = f.id;
    } else {                           // Drive API v2
      const f = Drive.Files.insert({ title: 'bt-ocr-temp', mimeType: blob.getContentType() },
                                   blob, { ocr: true, ocrLanguage: lang });
      id = f.id;
    }
    const text = DocumentApp.openById(id).getBody().getText();
    return { ok: true, text: text };
  } finally {
    if (id) { try { Drive.Files.remove(id); } catch (e) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) {} } }
  }
}

/** Test OCR from the Apps Script editor: select this function and Run to grant Drive/Docs permission */
function testGoogleOcr() {
  const png = Utilities.newBlob(Utilities.base64Decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='), 'image/png');
  const r = googleOcr_({ image: Utilities.base64Encode(png.getBytes()) });
  Logger.log('Google OCR is ready: ' + JSON.stringify(r));
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
