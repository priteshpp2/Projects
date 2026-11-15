const crypto = require('crypto');
const express = require('express');
const ExcelJS = require('exceljs');

const MILESTONES = [1, 3, 7, 14, 30, 50, 75, 100, 150, 200, 250, 300, 365];
const ADMIN_USER = process.env.ADMIN_USER || 'Pritesh';
const DEFAULT_PASSWORD = process.env.DEFAULT_PASSWORD || 'abc@123';
const AUTH_SHEET = '_auth';
const TOKEN_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const DONE_VALUES = new Set(['yes', 'y', '1', 'true', 'x', 'done', '✓', '✔', 'ok']);

// ---------- date helpers ----------
function todayStr(tzOffsetMin) {
  // tzOffsetMin = minutes to ADD to UTC to get the user's local time (e.g. IST = 330)
  const now = new Date();
  const local = new Date(now.getTime() + (tzOffsetMin ?? -now.getTimezoneOffset()) * 60000);
  return toKey(local.getUTCFullYear(), local.getUTCMonth() + 1, local.getUTCDate());
}
function toKey(y, m, d) {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}
function addDays(key, n) {
  const [y, m, d] = key.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + n));
  return toKey(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
}
function normalizeDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return toKey(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  if (typeof v === 'object' && v.result instanceof Date) return normalizeDate(v.result);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return toKey(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/); // dd/mm/yyyy
  if (m) return toKey(+m[3], +m[2], +m[1]);
  const parsed = new Date(s);
  if (!isNaN(parsed)) return toKey(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
  return null;
}
function isDone(v) {
  if (v == null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'object' && 'result' in v) return isDone(v.result);
  return DONE_VALUES.has(String(v).trim().toLowerCase());
}
function cellText(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if ('result' in v) return cellText(v.result);
    if (v.text) return String(v.text);
  }
  return String(v).trim();
}

// ---------- workbook I/O ----------
async function loadWorkbook(storage) {
  const wb = new ExcelJS.Workbook();
  const buf = await storage.load();
  if (buf && buf.length) await wb.xlsx.load(buf);
  return wb;
}
async function saveWorkbook(storage, wb) {
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  await storage.save(buf);
}

// ---------- auth (hidden sheet: Username | PasswordHash | Salt) ----------
const isAdminName = name => String(name).toLowerCase() === ADMIN_USER.toLowerCase();
const sameUser = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
const listUsers = wb => wb.worksheets.filter(ws => !ws.name.startsWith('_')).map(ws => ws.name);

function hashPassword(password, salt) {
  return crypto.scryptSync(String(password), salt, 32).toString('hex');
}

function readAuth(wb) {
  const map = new Map();
  const ws = wb.getWorksheet(AUTH_SHEET);
  if (!ws) return map;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const username = cellText(row.getCell(1).value);
    const hash = cellText(row.getCell(2).value);
    const salt = cellText(row.getCell(3).value);
    if (username && hash && salt) map.set(username.toLowerCase(), { username, hash, salt });
  }
  return map;
}

function writeAuth(wb, map) {
  const existing = wb.getWorksheet(AUTH_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(AUTH_SHEET, { state: 'hidden' });
  ws.addRow(['Username', 'PasswordHash', 'Salt']);
  for (const e of map.values()) ws.addRow([e.username, e.hash, e.salt]);
}

function setPassword(auth, username, password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const entry = { username, hash: hashPassword(password, salt), salt };
  auth.set(username.toLowerCase(), entry);
  return entry;
}

function checkPassword(auth, username, password) {
  const entry = auth.get(username.toLowerCase());
  if (!entry) return String(password) === DEFAULT_PASSWORD;
  const a = Buffer.from(hashPassword(password, entry.salt), 'hex');
  const b = Buffer.from(entry.hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

const b64url = s => Buffer.from(s).toString('base64url');
const sign = (payload, key) => crypto.createHmac('sha256', key).update(payload).digest('base64url');

function issueToken(entry) {
  const payload = b64url(JSON.stringify({ u: entry.username, exp: Date.now() + TOKEN_TTL_MS }));
  return `${payload}.${sign(payload, entry.hash)}`;
}

function verifyToken(auth, token) {
  if (!token || !token.includes('.')) return null;
  const [payload, sig] = token.split('.');
  let data;
  try {
    data = JSON.parse(Buffer.from(payload, 'base64url').toString());
  } catch {
    return null;
  }
  if (!data || !data.u || !(data.exp > Date.now())) return null;
  const entry = auth.get(String(data.u).toLowerCase());
  if (!entry) return null;
  const expected = sign(payload, entry.hash);
  const a = Buffer.from(sig || ''), b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  return { username: entry.username, isAdmin: isAdminName(entry.username) };
}

function parseSheet(ws) {
  const habits = [];
  const completions = {};
  if (!ws) return { habits, completions };
  const header = ws.getRow(1);
  const colToHabit = {};
  for (let c = 2; c <= ws.columnCount; c++) {
    const name = cellText(header.getCell(c).value);
    if (name) {
      colToHabit[c] = habits.length;
      habits.push(name);
    }
  }
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const date = normalizeDate(row.getCell(1).value);
    if (!date) continue;
    if (!completions[date]) completions[date] = new Set();
    for (const c of Object.keys(colToHabit)) {
      if (isDone(row.getCell(+c).value)) completions[date].add(colToHabit[c]);
    }
  }
  return { habits, completions };
}

function writeSheet(wb, username, model) {
  const existing = wb.getWorksheet(username);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(username);

  ws.addRow(['Date', ...model.habits]);
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6C5CE7' } };
  ws.getRow(1).alignment = { horizontal: 'center' };
  ws.getColumn(1).width = 14;
  for (let i = 0; i < model.habits.length; i++) {
    ws.getColumn(i + 2).width = Math.max(14, model.habits[i].length + 4);
  }

  for (const date of Object.keys(model.completions).sort()) {
    const set = model.completions[date];
    const values = [date];
    for (let i = 0; i < model.habits.length; i++) values.push(set.has(i) ? 'Yes' : '');
    const row = ws.addRow(values);
    row.alignment = { horizontal: 'center' };
    for (let i = 0; i < model.habits.length; i++) {
      if (set.has(i)) {
        row.getCell(i + 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8F7C5' } };
      }
    }
  }
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  return ws;
}

// ---------- streak math ----------
function computeStats(habitIndex, completions, today) {
  const doneDates = Object.keys(completions).filter(d => completions[d].has(habitIndex)).sort();
  const doneSet = new Set(doneDates);

  let current = 0;
  let cursor = doneSet.has(today) ? today : addDays(today, -1);
  while (doneSet.has(cursor)) {
    current++;
    cursor = addDays(cursor, -1);
  }

  let longest = 0, run = 0, prev = null;
  for (const d of doneDates) {
    run = prev && addDays(prev, 1) === d ? run + 1 : 1;
    if (run > longest) longest = run;
    prev = d;
  }

  const history = [];
  for (let i = 29; i >= 0; i--) {
    const d = addDays(today, -i);
    history.push({ date: d, done: doneSet.has(d) });
  }

  const badges = MILESTONES.filter(m => longest >= m);
  const nextMilestone = MILESTONES.find(m => m > current) || null;
  const prevMilestone = [...MILESTONES].reverse().find(m => m <= current) || 0;

  return {
    currentStreak: current,
    longestStreak: longest,
    totalCompletions: doneDates.length,
    doneToday: doneSet.has(today),
    badges,
    nextMilestone,
    prevMilestone,
    history,
  };
}

function buildUserView(username, model, today) {
  const habits = model.habits.map((name, i) => ({
    id: i,
    name,
    ...computeStats(i, model.completions, today),
  }));
  return { username, today, milestones: MILESTONES, habits };
}

// ---------- express app ----------
function httpError(message, status) {
  const e = new Error(message);
  e.status = status;
  return e;
}

function cleanUsername(raw) {
  const name = String(raw || '').trim().replace(/[\[\]\*\?\/\\:]/g, '').replace(/^_+/, '').slice(0, 31);
  if (!name) throw httpError('Username is required', 400);
  return name;
}

function cleanPassword(raw) {
  const pw = String(raw || '');
  if (pw.length < 4) throw httpError('Password must be at least 4 characters', 400);
  if (pw.length > 72) throw httpError('Password is too long', 400);
  return pw;
}

function createApp(storage) {
  const app = express();
  app.use(express.json());
  app.use(express.raw({ type: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/octet-stream'], limit: '10mb' }));

  const wrap = fn => (req, res) => fn(req, res).catch(err => {
    if (!err.status || err.status >= 500) console.error(err);
    res.status(err.status || 500).json({ error: err.message || 'Server error' });
  });

  const getToken = req => {
    const h = req.get('authorization') || '';
    return h.startsWith('Bearer ') ? h.slice(7) : (req.query.token || null);
  };

  // who is calling? throws 401 if not logged in
  const actorOf = (req, wb) => {
    const actor = verifyToken(readAuth(wb), getToken(req));
    if (!actor) throw httpError('Please log in', 401);
    return actor;
  };

  // may `actor` touch `target`'s data?
  const assertAccess = (actor, target) => {
    if (!actor.isAdmin && !sameUser(actor.username, target)) throw httpError('You can only access your own data', 403);
  };

  const assertAdmin = actor => {
    if (!actor.isAdmin) throw httpError('Admin only', 403);
  };

  // client sends its UTC offset (minutes) so "today" matches the user's clock
  const today = req => {
    const tz = Number(req.query.tz ?? (req.body && req.body.tz));
    return todayStr(Number.isFinite(tz) ? tz : undefined);
  };

  app.get('/api/users', wrap(async (req, res) => {
    const wb = await loadWorkbook(storage);
    res.json({ users: listUsers(wb), admin: ADMIN_USER });
  }));

  // register a new user (requires a password)
  app.post('/api/users', wrap(async (req, res) => {
    const username = cleanUsername(req.body.username);
    const password = cleanPassword(req.body.password);
    const wb = await loadWorkbook(storage);
    if (listUsers(wb).some(u => sameUser(u, username))) throw httpError('That name is already taken — log in instead', 409);
    const auth = readAuth(wb);
    const entry = setPassword(auth, username, password);
    writeSheet(wb, username, { habits: [], completions: {} });
    writeAuth(wb, auth);
    await saveWorkbook(storage, wb);
    res.json({ token: issueToken(entry), username, isAdmin: isAdminName(username) });
  }));

  app.post('/api/login', wrap(async (req, res) => {
    const username = cleanUsername(req.body.username);
    const password = String(req.body.password || '');
    const wb = await loadWorkbook(storage);
    const canonical = listUsers(wb).find(u => sameUser(u, username));
    if (!canonical) throw httpError('User not found', 404);
    const auth = readAuth(wb);
    if (!checkPassword(auth, canonical, password)) throw httpError('Incorrect password', 401);
    let entry = auth.get(canonical.toLowerCase());
    if (!entry) {
      // first login of a pre-existing user: persist the default password as a real hash
      entry = setPassword(auth, canonical, password);
      writeAuth(wb, auth);
      await saveWorkbook(storage, wb);
    }
    res.json({ token: issueToken(entry), username: canonical, isAdmin: isAdminName(canonical) });
  }));

  app.get('/api/me', wrap(async (req, res) => {
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    res.json({ ...actor, users: actor.isAdmin ? listUsers(wb) : undefined });
  }));

  app.post('/api/users/:username/password', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const newPassword = cleanPassword(req.body.newPassword);
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    assertAccess(actor, username);
    const canonical = listUsers(wb).find(u => sameUser(u, username));
    if (!canonical) throw httpError('User not found', 404);
    const auth = readAuth(wb);
    if (sameUser(actor.username, canonical) && !checkPassword(auth, canonical, req.body.currentPassword)) {
      throw httpError('Current password is incorrect', 400);
    }
    const entry = setPassword(auth, canonical, newPassword);
    writeAuth(wb, auth);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, token: sameUser(actor.username, canonical) ? issueToken(entry) : undefined });
  }));

  app.get('/api/users/:username', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const ws = wb.getWorksheet(username);
    if (!ws) throw httpError('User not found', 404);
    res.json(buildUserView(username, parseSheet(ws), today(req)));
  }));

  app.post('/api/users/:username/habits', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const name = String(req.body.name || '').trim().slice(0, 60);
    if (!name) throw httpError('Habit name is required', 400);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (model.habits.some(h => h.toLowerCase() === name.toLowerCase())) throw httpError('That habit already exists', 409);
    model.habits.push(name);
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    res.json(buildUserView(username, model, today(req)));
  }));

  app.delete('/api/users/:username/habits/:id', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    model.habits.splice(id, 1);
    const remapped = {};
    for (const [date, set] of Object.entries(model.completions)) {
      const ns = new Set();
      for (const i of set) {
        if (i < id) ns.add(i);
        else if (i > id) ns.add(i - 1);
      }
      if (ns.size) remapped[date] = ns;
    }
    model.completions = remapped;
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    res.json(buildUserView(username, model, today(req)));
  }));

  app.post('/api/users/:username/habits/:id/toggle', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const t = today(req);
    const date = normalizeDate(req.body.date) || t;
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    if (!model.completions[date]) model.completions[date] = new Set();
    const set = model.completions[date];
    const before = computeStats(id, model.completions, t);
    if (set.has(id)) set.delete(id); else set.add(id);
    if (!set.size) delete model.completions[date];
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    const view = buildUserView(username, model, t);
    const newBadges = view.habits[id].badges.filter(b => !before.badges.includes(b));
    res.json({ ...view, changed: { id, newBadges } });
  }));

  app.post('/api/users/:username/complete-all', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const t = today(req);
    const date = normalizeDate(req.body.date) || t;
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    const beforeView = buildUserView(username, model, t);
    if (!model.completions[date]) model.completions[date] = new Set();
    for (let i = 0; i < model.habits.length; i++) model.completions[date].add(i);
    if (!model.completions[date].size) delete model.completions[date];
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    const view = buildUserView(username, model, t);
    const newBadges = [];
    view.habits.forEach((h, i) => {
      const prev = beforeView.habits[i].badges;
      h.badges.filter(b => !prev.includes(b)).forEach(b => newBadges.push({ id: i, name: h.name, badge: b }));
    });
    res.json({ ...view, newBadges });
  }));

  // download the workbook
  app.get('/api/export', wrap(async (req, res) => {
    const buf = await storage.load();
    if (!buf) throw httpError('No data yet', 404);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    assertAdmin(actorOf(req, wb));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="streaks.xlsx"');
    res.setHeader('Cache-Control', 'no-store');
    res.send(buf);
  }));

  // upload a workbook edited offline (replaces all data)
  app.post('/api/import', wrap(async (req, res) => {
    if (!Buffer.isBuffer(req.body) || !req.body.length) throw httpError('Send the .xlsx file as the request body', 400);
    const current = await loadWorkbook(storage);
    assertAdmin(actorOf(req, current));
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(req.body);
    } catch {
      throw httpError('That file is not a valid .xlsx workbook', 400);
    }
    if (!wb.getWorksheet(AUTH_SHEET)) writeAuth(wb, readAuth(current));
    await saveWorkbook(storage, wb);
    res.json({ users: listUsers(wb) });
  }));

  return app;
}

module.exports = { createApp, MILESTONES };
