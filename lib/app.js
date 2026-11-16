const crypto = require('crypto');
const express = require('express');
const ExcelJS = require('exceljs');

const MILESTONES = [1, 3, 7, 14, 30, 50, 75, 100, 150, 200, 250, 300, 365];
const ADMIN_USER = process.env.ADMIN_USER || 'Pritesh';
const DEFAULT_PASSWORD = process.env.DEFAULT_PASSWORD || 'abc@123';
const AUTH_SHEET = '_auth';
const GROUPS_SHEET = '_groups';
const RESETS_SHEET = '_resets';
const SCHEDULES_SHEET = '_schedules';
const COLORS_SHEET = '_colors';
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
function weekday(key) { // 0 = Sunday ... 6 = Saturday
  const [y, m, d] = key.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay();
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

// ---------- auth (hidden sheet: Username | PasswordHash | Salt | ResetRequested) ----------
const isAdminName = name => String(name).toLowerCase() === ADMIN_USER.toLowerCase();
const sameUser = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
const listUsers = wb => wb.worksheets.filter(ws => !ws.name.startsWith('_')).map(ws => ws.name);
const findUser = (wb, name) => listUsers(wb).find(u => sameUser(u, name)) || null;
const userSheet = (wb, name) => wb.worksheets.find(ws => !ws.name.startsWith('_') && sameUser(ws.name, name));

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
    const resetRequested = cellText(row.getCell(4).value);
    if (username && hash && salt) map.set(username.toLowerCase(), { username, hash, salt, resetRequested });
  }
  return map;
}

function writeAuth(wb, map) {
  const existing = wb.getWorksheet(AUTH_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(AUTH_SHEET, { state: 'hidden' });
  ws.addRow(['Username', 'PasswordHash', 'Salt', 'ResetRequested']);
  for (const e of map.values()) ws.addRow([e.username, e.hash, e.salt, e.resetRequested || '']);
}

function setPassword(auth, username, password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const entry = { username, hash: hashPassword(password, salt), salt, resetRequested: '' };
  auth.set(username.toLowerCase(), entry);
  return entry;
}

const resetRequests = auth => [...auth.values()].filter(e => e.resetRequested).map(e => ({ username: e.username, at: e.resetRequested }));

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

// ---------- streak resets (hidden sheet: Username | Habit | From) ----------
// "Break streak": completions before `From` no longer count toward the current streak (history/badges stay).
const resetKey = (username, habit) => `${String(username).toLowerCase()}|${String(habit).trim().toLowerCase()}`;
function readResets(wb) {
  const map = new Map();
  const ws = wb.getWorksheet(RESETS_SHEET);
  if (!ws) return map;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const user = cellText(row.getCell(1).value), habit = cellText(row.getCell(2).value), from = normalizeDate(row.getCell(3).value);
    if (user && habit && from) map.set(resetKey(user, habit), { user, habit, from });
  }
  return map;
}
function writeResets(wb, map) {
  const existing = wb.getWorksheet(RESETS_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(RESETS_SHEET, { state: 'hidden' });
  ws.addRow(['Username', 'Habit', 'From']);
  for (const r of map.values()) ws.addRow([r.user, r.habit, r.from]);
}
const resetFor = (resets, username, habit) => (resets && resets.get(resetKey(username, habit)) || {}).from || null;

// ---------- schedules (hidden sheet: Username | Habit | Days) ----------
// Days = comma-separated weekday numbers (0=Sun..6=Sat). No entry = every day. Unscheduled days never break a streak.
function cleanDays(raw) {
  if (raw == null || raw === '') return null;
  const arr = Array.isArray(raw) ? raw : String(raw).split(',');
  const days = [...new Set(arr.map(Number).filter(n => Number.isInteger(n) && n >= 0 && n <= 6))].sort();
  if (!days.length) throw httpError('Pick at least one day', 400);
  return days.length === 7 ? null : days;
}
function readSchedules(wb) {
  const map = new Map();
  const ws = wb.getWorksheet(SCHEDULES_SHEET);
  if (!ws) return map;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const user = cellText(row.getCell(1).value), habit = cellText(row.getCell(2).value);
    let days = null;
    try { days = cleanDays(cellText(row.getCell(3).value)); } catch { days = null; }
    if (user && habit && days) map.set(resetKey(user, habit), { user, habit, days });
  }
  return map;
}
function writeSchedules(wb, map) {
  const existing = wb.getWorksheet(SCHEDULES_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(SCHEDULES_SHEET, { state: 'hidden' });
  ws.addRow(['Username', 'Habit', 'Days']);
  for (const s of map.values()) ws.addRow([s.user, s.habit, s.days.join(',')]);
}
const daysFor = (schedules, username, habit) => (schedules && schedules.get(resetKey(username, habit)) || {}).days || null;
const isScheduled = (days, key) => !days || days.includes(weekday(key));

// ---------- colors (hidden sheet: Username | Habit | Color) ----------
// Color is a hex like #6c5ce7 chosen by the user; no entry = auto palette on the client.
function cleanColor(raw) {
  if (raw == null || raw === '') return null;
  const s = String(raw).trim().toLowerCase();
  if (!/^#[0-9a-f]{6}$/.test(s)) throw httpError('Color must be a hex like #6c5ce7', 400);
  return s;
}
function readColors(wb) {
  const map = new Map();
  const ws = wb.getWorksheet(COLORS_SHEET);
  if (!ws) return map;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const user = cellText(row.getCell(1).value), habit = cellText(row.getCell(2).value);
    let color = null;
    try { color = cleanColor(cellText(row.getCell(3).value)); } catch { color = null; }
    if (user && habit && color) map.set(resetKey(user, habit), { user, habit, color });
  }
  return map;
}
function writeColors(wb, map) {
  const existing = wb.getWorksheet(COLORS_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(COLORS_SHEET, { state: 'hidden' });
  ws.addRow(['Username', 'Habit', 'Color']);
  for (const c of map.values()) ws.addRow([c.user, c.habit, c.color]);
}
const colorFor = (colors, username, habit) => (colors && colors.get(resetKey(username, habit)) || {}).color || null;

// move every per-habit metadata row (reset / schedule / color) from an old habit name to a new one
function rekeyHabitMeta(wb, username, oldName, newName) {
  for (const [read, write, field] of [[readResets, writeResets, 'from'], [readSchedules, writeSchedules, 'days'], [readColors, writeColors, 'color']]) {
    const map = read(wb);
    const entry = map.get(resetKey(username, oldName));
    if (!entry) continue;
    map.delete(resetKey(username, oldName));
    map.set(resetKey(username, newName), { user: username, habit: newName, [field]: entry[field] });
    write(wb, map);
  }
}

// ---------- streak math ----------
function computeStats(habitIndex, completions, today, { resetFrom = null, days = null } = {}) {
  const doneDates = Object.keys(completions).filter(d => completions[d].has(habitIndex)).sort();
  const doneSet = new Set(doneDates);

  // current streak: walk back from today; a done day counts, an unscheduled day is skipped, a missed scheduled day ends it
  let current = 0;
  let cursor = doneSet.has(today) ? today : addDays(today, -1);
  for (let guard = 0; guard < 4000 && (!resetFrom || cursor >= resetFrom); guard++) {
    if (doneSet.has(cursor)) current++;
    else if (isScheduled(days, cursor)) break;
    cursor = addDays(cursor, -1);
  }

  let longest = 0, run = 0;
  if (doneDates.length) {
    for (let d = doneDates[0]; d <= today; d = addDays(d, 1)) {
      if (doneSet.has(d)) { run++; if (run > longest) longest = run; }
      else if (isScheduled(days, d)) run = 0;
    }
  }

  const history = [];
  for (let i = 29; i >= 0; i--) {
    const d = addDays(today, -i);
    history.push({ date: d, done: doneSet.has(d), scheduled: isScheduled(days, d) });
  }

  const badges = MILESTONES.filter(m => longest >= m);
  const nextMilestone = MILESTONES.find(m => m > current) || null;
  const prevMilestone = [...MILESTONES].reverse().find(m => m <= current) || 0;

  return {
    currentStreak: current,
    longestStreak: longest,
    totalCompletions: doneDates.length,
    doneToday: doneSet.has(today),
    scheduledToday: isScheduled(days, today),
    days,
    badges,
    nextMilestone,
    prevMilestone,
    history,
  };
}

function buildUserView(username, model, today, resets = null, schedules = null, colors = null) {
  const habits = model.habits.map((name, i) => ({
    id: i,
    name,
    color: colorFor(colors, username, name),
    ...computeStats(i, model.completions, today, { resetFrom: resetFor(resets, username, name), days: daysFor(schedules, username, name) }),
  }));
  return { username, today, milestones: MILESTONES, habits };
}

// ---------- groups (hidden sheet: Id | Name | Habit | Owner | Members) ----------
// A group is one shared habit name across members; every member keeps their own streak on their own sheet.
function readGroups(wb) {
  const ws = wb.getWorksheet(GROUPS_SHEET);
  const groups = [];
  if (!ws) return groups;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const id = cellText(row.getCell(1).value);
    const name = cellText(row.getCell(2).value);
    const habit = cellText(row.getCell(3).value);
    const owner = cellText(row.getCell(4).value);
    const members = cellText(row.getCell(5).value).split('|').map(s => s.trim()).filter(Boolean);
    if (id && name && habit && owner) groups.push({ id, name, habit, owner, members });
  }
  return groups;
}

function writeGroups(wb, groups) {
  const existing = wb.getWorksheet(GROUPS_SHEET);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(GROUPS_SHEET, { state: 'hidden' });
  ws.addRow(['Id', 'Name', 'Habit', 'Owner', 'Members']);
  for (const g of groups) ws.addRow([g.id, g.name, g.habit, g.owner, g.members.join('|')]);
}

const sameHabit = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

// make sure `username` has the group's habit on their sheet
function ensureHabit(wb, username, habit) {
  const model = parseSheet(userSheet(wb, username));
  if (model.habits.some(h => sameHabit(h, habit))) return false;
  model.habits.push(habit);
  writeSheet(wb, username, model);
  return true;
}

function groupBoard(wb, group, today, resets = readResets(wb), schedules = readSchedules(wb)) {
  const rows = [];
  for (const m of group.members) {
    const ws = userSheet(wb, m);
    if (!ws) continue;
    const model = parseSheet(ws);
    const idx = model.habits.findIndex(h => sameHabit(h, group.habit));
    const stats = idx >= 0 ? computeStats(idx, model.completions, today, { resetFrom: resetFor(resets, ws.name, group.habit), days: daysFor(schedules, ws.name, group.habit) }) : { currentStreak: 0, longestStreak: 0, totalCompletions: 0, doneToday: false };
    rows.push({ username: ws.name, currentStreak: stats.currentStreak, longestStreak: stats.longestStreak, totalCompletions: stats.totalCompletions, doneToday: stats.doneToday, hasHabit: idx >= 0 });
  }
  rows.sort((a, b) => b.currentStreak - a.currentStreak || b.longestStreak - a.longestStreak || b.totalCompletions - a.totalCompletions || a.username.localeCompare(b.username));
  rows.forEach((r, i) => { r.rank = i > 0 && rows[i - 1].currentStreak === r.currentStreak ? rows[i - 1].rank : i + 1; });
  return { id: group.id, name: group.name, habit: group.habit, owner: group.owner, members: group.members, board: rows };
}

function groupsFor(wb, username, today, resets, schedules) {
  return readGroups(wb).filter(g => g.members.some(m => sameUser(m, username))).map(g => groupBoard(wb, g, today, resets, schedules));
}

// user view + the groups they belong to, with each habit tagged by its group rank
function fullView(wb, username, model, today) {
  const resets = readResets(wb), schedules = readSchedules(wb);
  const view = buildUserView(username, model, today, resets, schedules, readColors(wb));
  view.groups = groupsFor(wb, username, today, resets, schedules);
  for (const h of view.habits) {
    const g = view.groups.find(g => sameHabit(g.habit, h.name));
    if (!g) continue;
    const row = g.board.find(r => sameUser(r.username, username));
    h.group = { id: g.id, name: g.name, rank: row ? row.rank : null, size: g.board.length };
  }
  return view;
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
    res.json({
      ...actor,
      users: actor.isAdmin ? listUsers(wb) : undefined,
      resetRequests: actor.isAdmin ? resetRequests(readAuth(wb)) : undefined,
    });
  }));

  // user forgot their password: flag it for the admin (public, no login needed)
  app.post('/api/users/:username/reset-request', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    const canonical = listUsers(wb).find(u => sameUser(u, username));
    if (!canonical) throw httpError('User not found', 404);
    const auth = readAuth(wb);
    const entry = auth.get(canonical.toLowerCase()) || setPassword(auth, canonical, DEFAULT_PASSWORD);
    entry.resetRequested = new Date().toISOString();
    writeAuth(wb, auth);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, message: `Request sent — ${ADMIN_USER} will reset your password` });
  }));

  // admin: reset a user's password to the default (and clear the request)
  app.post('/api/users/:username/reset', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    assertAdmin(actorOf(req, wb));
    const canonical = listUsers(wb).find(u => sameUser(u, username));
    if (!canonical) throw httpError('User not found', 404);
    const auth = readAuth(wb);
    setPassword(auth, canonical, DEFAULT_PASSWORD);
    writeAuth(wb, auth);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, username: canonical, resetRequests: resetRequests(auth) });
  }));

  // admin: dismiss a reset request without changing the password
  app.delete('/api/users/:username/reset-request', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    assertAdmin(actorOf(req, wb));
    const auth = readAuth(wb);
    const entry = auth.get(username.toLowerCase());
    if (entry && entry.resetRequested) {
      entry.resetRequested = '';
      writeAuth(wb, auth);
      await saveWorkbook(storage, wb);
    }
    res.json({ ok: true, resetRequests: resetRequests(auth) });
  }));

  // admin: delete a user and all their data
  app.delete('/api/users/:username', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    assertAdmin(actor);
    if (sameUser(actor.username, username)) throw httpError('You cannot delete your own account', 400);
    const ws = wb.worksheets.find(w => !w.name.startsWith('_') && sameUser(w.name, username));
    if (!ws) throw httpError('User not found', 404);
    wb.removeWorksheet(ws.id);
    const auth = readAuth(wb);
    auth.delete(username.toLowerCase());
    writeAuth(wb, auth);
    // drop them from groups; hand ownership to the next member or dissolve empty groups
    const groups = readGroups(wb).map(g => ({ ...g, members: g.members.filter(m => !sameUser(m, username)) }))
      .filter(g => g.members.length)
      .map(g => (sameUser(g.owner, username) ? { ...g, owner: g.members[0] } : g));
    writeGroups(wb, groups);
    const resets = readResets(wb);
    for (const [k, v] of resets) if (sameUser(v.user, username)) resets.delete(k);
    writeResets(wb, resets);
    const schedules = readSchedules(wb);
    for (const [k, v] of schedules) if (sameUser(v.user, username)) schedules.delete(k);
    writeSchedules(wb, schedules);
    const colors = readColors(wb);
    for (const [k, v] of colors) if (sameUser(v.user, username)) colors.delete(k);
    writeColors(wb, colors);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, users: listUsers(wb) });
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
    const ws = userSheet(wb, username);
    if (!ws) throw httpError('User not found', 404);
    res.json(fullView(wb, ws.name, parseSheet(ws), today(req)));
  }));

  app.post('/api/users/:username/habits', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const name = String(req.body.name || '').trim().slice(0, 60);
    if (!name) throw httpError('Habit name is required', 400);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (model.habits.some(h => h.toLowerCase() === name.toLowerCase())) throw httpError('That habit already exists', 409);
    const days = cleanDays(req.body.days);
    const color = cleanColor(req.body.color);
    model.habits.push(name);
    writeSheet(wb, username, model);
    const schedules = readSchedules(wb);
    if (days) schedules.set(resetKey(username, name), { user: username, habit: name, days }); else schedules.delete(resetKey(username, name));
    writeSchedules(wb, schedules);
    const colors = readColors(wb);
    if (color) colors.set(resetKey(username, name), { user: username, habit: name, color }); else colors.delete(resetKey(username, name));
    writeColors(wb, colors);
    await saveWorkbook(storage, wb);
    res.json(fullView(wb, username, model, today(req)));
  }));

  // edit a habit: any of { name, days, color }. Group-shared habits cannot be renamed (the name links members).
  app.put('/api/users/:username/habits/:id', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    let habit = model.habits[id];
    if (req.body.name !== undefined) {
      const name = String(req.body.name || '').trim().slice(0, 60);
      if (!name) throw httpError('Habit name is required', 400);
      if (name !== habit) {
        if (model.habits.some((h, i) => i !== id && sameHabit(h, name))) throw httpError('That habit already exists', 409);
        const inGroup = readGroups(wb).some(g => sameHabit(g.habit, habit) && g.members.some(m => sameUser(m, username)));
        if (inGroup) throw httpError('This habit is part of a group challenge and cannot be renamed', 409);
        rekeyHabitMeta(wb, username, habit, name);
        model.habits[id] = name;
        writeSheet(wb, username, model);
        habit = name;
      }
    }
    if (req.body.days !== undefined) {
      const days = cleanDays(req.body.days);
      const schedules = readSchedules(wb);
      if (days) schedules.set(resetKey(username, habit), { user: username, habit, days }); else schedules.delete(resetKey(username, habit));
      writeSchedules(wb, schedules);
    }
    if (req.body.color !== undefined) {
      const color = cleanColor(req.body.color);
      const colors = readColors(wb);
      if (color) colors.set(resetKey(username, habit), { user: username, habit, color }); else colors.delete(resetKey(username, habit));
      writeColors(wb, colors);
    }
    await saveWorkbook(storage, wb);
    res.json(fullView(wb, username, model, today(req)));
  }));

  // change which weekdays a habit is expected on (null/all 7 = daily)
  app.put('/api/users/:username/habits/:id/schedule', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const days = cleanDays(req.body.days);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    const schedules = readSchedules(wb);
    const key = resetKey(username, model.habits[id]);
    if (days) schedules.set(key, { user: username, habit: model.habits[id], days }); else schedules.delete(key);
    writeSchedules(wb, schedules);
    await saveWorkbook(storage, wb);
    res.json(fullView(wb, username, model, today(req)));
  }));

  app.delete('/api/users/:username/habits/:id', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    const resets = readResets(wb);
    if (resets.delete(resetKey(username, model.habits[id]))) writeResets(wb, resets);
    const schedules = readSchedules(wb);
    if (schedules.delete(resetKey(username, model.habits[id]))) writeSchedules(wb, schedules);
    const colors = readColors(wb);
    if (colors.delete(resetKey(username, model.habits[id]))) writeColors(wb, colors);
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
    res.json(fullView(wb, username, model, today(req)));
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
    const before = computeStats(id, model.completions, t, { resetFrom: resetFor(readResets(wb), username, model.habits[id]), days: daysFor(readSchedules(wb), username, model.habits[id]) });
    if (set.has(id)) set.delete(id); else set.add(id);
    if (!set.size) delete model.completions[date];
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    const view = fullView(wb, username, model, t);
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
    const schedules = readSchedules(wb);
    const beforeView = buildUserView(username, model, t, readResets(wb), schedules);
    if (!model.completions[date]) model.completions[date] = new Set();
    // only habits scheduled on that day; unscheduled ones stay untouched
    for (let i = 0; i < model.habits.length; i++) if (isScheduled(daysFor(schedules, username, model.habits[i]), date)) model.completions[date].add(i);
    if (!model.completions[date].size) delete model.completions[date];
    writeSheet(wb, username, model);
    await saveWorkbook(storage, wb);
    const view = fullView(wb, username, model, t);
    const newBadges = [];
    view.habits.forEach((h, i) => {
      const prev = beforeView.habits[i].badges;
      h.badges.filter(b => !prev.includes(b)).forEach(b => newBadges.push({ id: i, name: h.name, badge: b }));
    });
    res.json({ ...view, newBadges });
  }));

  // break the current streak: it restarts from today (today's tick is cleared); history, totals and badges are kept
  app.post('/api/users/:username/habits/:id/break', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const id = Number(req.params.id);
    const t = today(req);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const model = parseSheet(wb.getWorksheet(username));
    if (!(id >= 0 && id < model.habits.length)) throw httpError('Habit not found', 404);
    if (model.completions[t]) {
      model.completions[t].delete(id);
      if (!model.completions[t].size) delete model.completions[t];
      writeSheet(wb, username, model);
    }
    const resets = readResets(wb);
    resets.set(resetKey(username, model.habits[id]), { user: username, habit: model.habits[id], from: t });
    writeResets(wb, resets);
    await saveWorkbook(storage, wb);
    res.json(fullView(wb, username, model, t));
  }));

  // ---------- groups ----------
  const cleanGroupName = raw => {
    const name = String(raw || '').trim().slice(0, 40);
    if (!name) throw httpError('Group name is required', 400);
    return name;
  };
  const cleanHabitName = raw => {
    const name = String(raw || '').trim().slice(0, 60);
    if (!name) throw httpError('Habit name is required', 400);
    return name;
  };
  const findGroup = (groups, id) => {
    const g = groups.find(g => g.id === String(id));
    if (!g) throw httpError('Group not found', 404);
    return g;
  };
  const assertGroupOwner = (actor, g) => {
    if (!actor.isAdmin && !sameUser(actor.username, g.owner)) throw httpError('Only the group owner can do that', 403);
  };

  // create a group challenge: `username` is the owner; every member gets the habit on their sheet
  app.post('/api/users/:username/groups', wrap(async (req, res) => {
    const username = cleanUsername(req.params.username);
    const name = cleanGroupName(req.body.name);
    const habit = cleanHabitName(req.body.habit);
    const wb = await loadWorkbook(storage);
    assertAccess(actorOf(req, wb), username);
    const owner = findUser(wb, username);
    if (!owner) throw httpError('User not found', 404);
    const wanted = Array.isArray(req.body.members) ? req.body.members : [];
    const members = [owner];
    for (const m of wanted) {
      const canonical = findUser(wb, cleanUsername(m));
      if (!canonical) throw httpError(`User "${m}" not found`, 404);
      if (!members.some(x => sameUser(x, canonical))) members.push(canonical);
    }
    const groups = readGroups(wb);
    if (groups.some(g => sameHabit(g.name, name))) throw httpError('A group with that name already exists', 409);
    const group = { id: crypto.randomBytes(4).toString('hex'), name, habit, owner, members };
    for (const m of members) ensureHabit(wb, m, habit);
    groups.push(group);
    writeGroups(wb, groups);
    await saveWorkbook(storage, wb);
    res.json(fullView(wb, owner, parseSheet(userSheet(wb, owner)), today(req)));
  }));

  // owner/admin adds a member
  app.post('/api/groups/:id/members', wrap(async (req, res) => {
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    const groups = readGroups(wb);
    const g = findGroup(groups, req.params.id);
    assertGroupOwner(actor, g);
    const canonical = findUser(wb, cleanUsername(req.body.username));
    if (!canonical) throw httpError('User not found', 404);
    if (g.members.some(m => sameUser(m, canonical))) throw httpError(`${canonical} is already in the group`, 409);
    g.members.push(canonical);
    ensureHabit(wb, canonical, g.habit);
    writeGroups(wb, groups);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, group: groupBoard(wb, g, today(req)) });
  }));

  // owner/admin removes a member, or a member leaves (their habit + history stay on their sheet)
  app.delete('/api/groups/:id/members/:username', wrap(async (req, res) => {
    const target = cleanUsername(req.params.username);
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    let groups = readGroups(wb);
    const g = findGroup(groups, req.params.id);
    if (!sameUser(actor.username, target)) assertGroupOwner(actor, g);
    if (!g.members.some(m => sameUser(m, target))) throw httpError('Not a member of this group', 404);
    g.members = g.members.filter(m => !sameUser(m, target));
    if (!g.members.length) groups = groups.filter(x => x !== g);
    else if (sameUser(g.owner, target)) g.owner = g.members[0];
    writeGroups(wb, groups);
    await saveWorkbook(storage, wb);
    res.json({ ok: true, group: g.members.length ? groupBoard(wb, g, today(req)) : null });
  }));

  // owner/admin deletes the group (members keep their habits)
  app.delete('/api/groups/:id', wrap(async (req, res) => {
    const wb = await loadWorkbook(storage);
    const actor = actorOf(req, wb);
    const groups = readGroups(wb);
    const g = findGroup(groups, req.params.id);
    assertGroupOwner(actor, g);
    writeGroups(wb, groups.filter(x => x !== g));
    await saveWorkbook(storage, wb);
    res.json({ ok: true });
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
    if (!wb.getWorksheet(GROUPS_SHEET)) writeGroups(wb, readGroups(current));
    if (!wb.getWorksheet(RESETS_SHEET)) writeResets(wb, readResets(current));
    if (!wb.getWorksheet(SCHEDULES_SHEET)) writeSchedules(wb, readSchedules(current));
    if (!wb.getWorksheet(COLORS_SHEET)) writeColors(wb, readColors(current));
    await saveWorkbook(storage, wb);
    res.json({ users: listUsers(wb) });
  }));

  return app;
}

module.exports = { createApp, MILESTONES };
