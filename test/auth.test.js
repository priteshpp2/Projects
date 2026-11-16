// Run: node test/auth.test.js
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ExcelJS = require('exceljs');
const { createApp } = require('../lib/app');
const { fileStorage } = require('../lib/storage');

async function main() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'streak-test-'));
  const storage = fileStorage(dir);

  // seed a legacy workbook: two users, no _auth sheet (simulates existing data)
  const wb = new ExcelJS.Workbook();
  for (const name of ['Pritesh', 'Amrita']) {
    const ws = wb.addWorksheet(name);
    ws.addRow(['Date', 'Exercise']);
    ws.addRow(['2026-11-15', 'Yes']);
  }
  await storage.save(Buffer.from(await wb.xlsx.writeBuffer()));

  const server = createApp(storage).listen(0);
  const base = `http://127.0.0.1:${server.address().port}`;
  const call = async (method, url, { body, token, raw } = {}) => {
    const res = await fetch(base + url, {
      method,
      headers: { ...(raw ? { 'Content-Type': 'application/octet-stream' } : { 'Content-Type': 'application/json' }), ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: raw ? body : body ? JSON.stringify(body) : undefined,
    });
    const ct = res.headers.get('content-type') || '';
    return { status: res.status, body: ct.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()) };
  };

  try {
    let r;
    r = await call('GET', '/api/users');
    assert.deepStrictEqual(r.body.users, ['Pritesh', 'Amrita'], 'lists users (no _auth leak)');

    r = await call('GET', '/api/users/Pritesh');
    assert.strictEqual(r.status, 401, 'no token -> 401');

    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'wrong' } });
    assert.strictEqual(r.status, 401, 'wrong password rejected');

    r = await call('POST', '/api/login', { body: { username: 'amrita', password: 'abc@123' } });
    assert.strictEqual(r.status, 200, 'default password works for legacy user');
    assert.strictEqual(r.body.username, 'Amrita');
    assert.strictEqual(r.body.isAdmin, false);
    const amrita = r.body.token;

    r = await call('POST', '/api/login', { body: { username: 'Pritesh', password: 'abc@123' } });
    assert.strictEqual(r.body.isAdmin, true, 'Pritesh is admin');
    const admin = r.body.token;

    r = await call('GET', '/api/users/Amrita', { token: amrita });
    assert.strictEqual(r.status, 200, 'own data OK');
    r = await call('GET', '/api/users/Pritesh', { token: amrita });
    assert.strictEqual(r.status, 403, 'other user blocked');
    r = await call('POST', '/api/users/Pritesh/habits', { token: amrita, body: { name: 'Hack' } });
    assert.strictEqual(r.status, 403, 'cannot write others data');
    r = await call('GET', '/api/users/Amrita', { token: admin });
    assert.strictEqual(r.status, 200, 'admin reads anyone');

    r = await call('GET', '/api/export', { token: amrita });
    assert.strictEqual(r.status, 403, 'export admin-only');
    r = await call('GET', '/api/export', { token: admin });
    assert.strictEqual(r.status, 200, 'admin export');
    const exported = r.body;
    assert.strictEqual(exported.slice(0, 2).toString(), 'PK');

    r = await call('POST', '/api/import', { token: amrita, raw: true, body: exported });
    assert.strictEqual(r.status, 403, 'import admin-only');

    r = await call('POST', '/api/users/Amrita/password', { token: amrita, body: { currentPassword: 'nope', newPassword: 'newpass1' } });
    assert.strictEqual(r.status, 400, 'wrong current password');
    r = await call('POST', '/api/users/Amrita/password', { token: amrita, body: { currentPassword: 'abc@123', newPassword: 'newpass1' } });
    assert.strictEqual(r.status, 200, 'password changed');
    const amrita2 = r.body.token;
    r = await call('GET', '/api/users/Amrita', { token: amrita });
    assert.strictEqual(r.status, 401, 'old token invalid after password change');
    r = await call('GET', '/api/users/Amrita', { token: amrita2 });
    assert.strictEqual(r.status, 200, 'new token valid');
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'abc@123' } });
    assert.strictEqual(r.status, 401, 'default password no longer works');

    r = await call('POST', '/api/users/Amrita/password', { token: admin, body: { newPassword: 'reset123' } });
    assert.strictEqual(r.status, 200, 'admin resets without current password');
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'reset123' } });
    assert.strictEqual(r.status, 200);
    const amrita3 = r.body.token;

    r = await call('POST', '/api/users', { body: { username: 'Amrita', password: 'xxxx' } });
    assert.strictEqual(r.status, 409, 'duplicate register blocked');
    r = await call('POST', '/api/users', { body: { username: 'Rohan', password: '123' } });
    assert.strictEqual(r.status, 400, 'short password rejected');
    r = await call('POST', '/api/users', { body: { username: '_auth', password: 'abcd' } });
    assert.strictEqual(r.status, 200, 'leading underscore stripped');
    assert.strictEqual(r.body.username, 'auth');

    // import of a workbook with no _auth sheet keeps existing passwords
    r = await call('POST', '/api/import', { token: admin, raw: true, body: Buffer.from(await wb.xlsx.writeBuffer()) });
    assert.strictEqual(r.status, 200);
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'reset123' } });
    assert.strictEqual(r.status, 200, 'passwords preserved across import');

    r = await call('GET', '/api/users');
    assert.ok(!r.body.users.includes('_auth'), '_auth hidden from user list');

    // forgot password -> admin sees request -> resets to default
    r = await call('POST', '/api/users/Nobody/reset-request', { body: {} });
    assert.strictEqual(r.status, 404, 'reset request for unknown user');
    r = await call('POST', '/api/users/amrita/reset-request', { body: {} });
    assert.strictEqual(r.status, 200, 'reset request accepted without login');
    r = await call('GET', '/api/me', { token: amrita3 });
    assert.strictEqual(r.body.resetRequests, undefined, 'non-admin does not see requests');
    r = await call('GET', '/api/me', { token: admin });
    assert.deepStrictEqual(r.body.resetRequests.map(x => x.username), ['Amrita'], 'admin sees pending request');
    r = await call('POST', '/api/users/Amrita/reset', { token: amrita3, body: {} });
    assert.strictEqual(r.status, 403, 'only admin can reset');
    r = await call('POST', '/api/users/Amrita/reset', { token: admin, body: {} });
    assert.strictEqual(r.status, 200);
    assert.deepStrictEqual(r.body.resetRequests, [], 'request cleared after reset');
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'abc@123' } });
    assert.strictEqual(r.status, 200, 'default password works after admin reset');
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'reset123' } });
    assert.strictEqual(r.status, 401, 'old password dead after reset');

    r = await call('POST', '/api/users/Amrita/reset-request', { body: {} });
    r = await call('DELETE', '/api/users/Amrita/reset-request', { token: admin });
    assert.deepStrictEqual(r.body.resetRequests, [], 'dismiss clears request');

    // admin deletes a user
    r = await call('POST', '/api/users', { body: { username: 'Rohan', password: 'abcd' } });
    assert.strictEqual(r.status, 200);
    r = await call('DELETE', '/api/users/Rohan', { token: amrita3 });
    assert.strictEqual(r.status, 401, 'token invalidated by admin reset');
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'abc@123' } });
    r = await call('DELETE', '/api/users/Rohan', { token: r.body.token });
    assert.strictEqual(r.status, 403, 'non-admin cannot delete users');
    r = await call('DELETE', '/api/users/Pritesh', { token: admin });
    assert.strictEqual(r.status, 400, 'admin cannot delete self');
    r = await call('DELETE', '/api/users/rohan', { token: admin });
    assert.strictEqual(r.status, 200);
    assert.deepStrictEqual(r.body.users, ['Pritesh', 'Amrita'], 'user removed');
    r = await call('POST', '/api/login', { body: { username: 'Rohan', password: 'abcd' } });
    assert.strictEqual(r.status, 404, 'deleted user cannot log in');

    r = await call('GET', '/api/export', { token: admin });
    const check = new ExcelJS.Workbook();
    await check.xlsx.load(r.body);
    assert.deepStrictEqual(check.worksheets.map(w => w.name).sort(), ['Amrita', 'Pritesh', '_auth', '_groups', '_resets'], 'workbook consistent');

    // ---- break streak ----
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'abc@123' } });
    const amritaB = r.body.token;
    r = await call('POST', '/api/users/Amrita/habits', { token: amritaB, body: { name: 'Yoga' } });
    const exId = r.body.habits.findIndex(h => h.name === 'Yoga');
    const todayKey = r.body.today;
    const dayBefore = d => { const x = new Date(d + 'T00:00:00Z'); x.setUTCDate(x.getUTCDate() - 1); return x.toISOString().slice(0, 10); };
    await call('POST', `/api/users/Amrita/habits/${exId}/toggle`, { token: amritaB, body: { date: dayBefore(todayKey) } });
    r = await call('POST', `/api/users/Amrita/habits/${exId}/toggle`, { token: amritaB, body: {} });
    assert.strictEqual(r.body.habits[exId].currentStreak, 2, 'streak of 2 before break');
    const totalBefore = r.body.habits[exId].totalCompletions;
    r = await call('POST', `/api/users/Amrita/habits/${exId}/break`, { token: amritaB, body: {} });
    assert.strictEqual(r.status, 200, 'break ok');
    assert.strictEqual(r.body.habits[exId].currentStreak, 0, 'streak reset to 0');
    assert.strictEqual(r.body.habits[exId].doneToday, false, 'today cleared');
    assert.strictEqual(r.body.habits[exId].totalCompletions, totalBefore - 1, 'history kept (only today removed)');
    assert.strictEqual(r.body.habits[exId].longestStreak, 1, 'past history still counts toward longest');
    r = await call('POST', `/api/users/Amrita/habits/${exId}/toggle`, { token: amritaB, body: {} });
    assert.strictEqual(r.body.habits[exId].currentStreak, 1, 'fresh streak starts at 1 after break');
    r = await call('POST', `/api/users/Amrita/habits/${exId}/break`, { token: amrita });
    assert.strictEqual(r.status, 401, 'revoked token cannot break');

    // ---- group challenges ----
    r = await call('POST', '/api/login', { body: { username: 'Amrita', password: 'abc@123' } });
    const amrita4 = r.body.token;
    r = await call('POST', '/api/users', { body: { username: 'Rohan', password: 'abcd' } });
    const rohan = r.body.token;

    r = await call('POST', '/api/users/Amrita/groups', { token: amrita4, body: { name: 'Runners', habit: 'Morning Run', members: ['rohan'] } });
    assert.strictEqual(r.status, 200, 'group created');
    assert.strictEqual(r.body.groups.length, 1);
    const gid = r.body.groups[0].id;
    assert.deepStrictEqual(r.body.groups[0].members, ['Amrita', 'Rohan'], 'owner + members');
    assert.ok(r.body.habits.some(h => h.name === 'Morning Run' && h.group && h.group.id === gid), 'habit added to owner and tagged');
    r = await call('GET', '/api/users/Rohan', { token: rohan });
    assert.ok(r.body.habits.some(h => h.name === 'Morning Run'), 'habit added to member sheet');
    assert.strictEqual(r.body.groups.length, 1, 'member sees the group');

    r = await call('POST', '/api/users/Amrita/groups', { token: amrita4, body: { name: 'runners', habit: 'x' } });
    assert.strictEqual(r.status, 409, 'duplicate group name blocked');
    r = await call('POST', '/api/users/Amrita/groups', { token: amrita4, body: { name: 'Ghosts', habit: 'x', members: ['Nobody'] } });
    assert.strictEqual(r.status, 404, 'unknown member rejected');

    // Rohan completes today -> leads the board
    r = await call('GET', '/api/users/Rohan', { token: rohan });
    const runId = r.body.habits.find(h => h.name === 'Morning Run').id;
    r = await call('POST', `/api/users/Rohan/habits/${runId}/toggle`, { token: rohan, body: {} });
    assert.strictEqual(r.status, 200);
    assert.deepStrictEqual(r.body.groups[0].board.map(x => [x.username, x.rank, x.currentStreak]), [['Rohan', 1, 1], ['Amrita', 2, 0]], 'leaderboard ranks by streak');
    assert.strictEqual(r.body.habits[runId].group.rank, 1, 'habit tagged with rank');

    r = await call('POST', `/api/groups/${gid}/members`, { token: rohan, body: { username: 'Pritesh' } });
    assert.strictEqual(r.status, 403, 'non-owner cannot add members');
    r = await call('POST', `/api/groups/${gid}/members`, { token: admin, body: { username: 'Pritesh' } });
    assert.strictEqual(r.status, 200, 'admin can add members');
    assert.strictEqual(r.body.group.board.length, 3);
    r = await call('POST', `/api/groups/${gid}/members`, { token: amrita4, body: { username: 'pritesh' } });
    assert.strictEqual(r.status, 409, 'duplicate member blocked');

    r = await call('DELETE', `/api/groups/${gid}/members/Amrita`, { token: rohan });
    assert.strictEqual(r.status, 403, 'member cannot kick owner');
    r = await call('DELETE', `/api/groups/${gid}/members/Rohan`, { token: rohan });
    assert.strictEqual(r.status, 200, 'member can leave');
    assert.deepStrictEqual(r.body.group.members, ['Amrita', 'Pritesh']);
    r = await call('GET', '/api/users/Rohan', { token: rohan });
    assert.strictEqual(r.body.groups.length, 0, 'left group gone');
    assert.ok(r.body.habits.some(h => h.name === 'Morning Run'), 'habit kept after leaving');

    r = await call('DELETE', `/api/groups/${gid}/members/Amrita`, { token: amrita4 });
    assert.strictEqual(r.status, 200, 'owner leaves');
    assert.strictEqual(r.body.group.owner, 'Pritesh', 'ownership transferred');

    r = await call('DELETE', `/api/groups/${gid}`, { token: amrita4 });
    assert.strictEqual(r.status, 403, 'non-owner cannot delete group');
    r = await call('DELETE', `/api/groups/${gid}`, { token: admin });
    assert.strictEqual(r.status, 200, 'owner/admin deletes group');
    r = await call('GET', '/api/users/Pritesh', { token: admin });
    assert.strictEqual(r.body.groups.length, 0, 'group removed');

    // deleting a user cleans groups
    r = await call('POST', '/api/users/Rohan/groups', { token: rohan, body: { name: 'Solo', habit: 'Read', members: ['Amrita'] } });
    const gid2 = r.body.groups[0].id;
    r = await call('DELETE', '/api/users/Rohan', { token: admin });
    assert.strictEqual(r.status, 200);
    r = await call('GET', '/api/users/Amrita', { token: amrita4 });
    assert.strictEqual(r.body.groups.length, 1);
    assert.strictEqual(r.body.groups[0].id, gid2);
    assert.strictEqual(r.body.groups[0].owner, 'Amrita', 'ownership moved after owner deleted');
    assert.deepStrictEqual(r.body.groups[0].members, ['Amrita']);

    console.log('All auth tests passed');
  } finally {
    server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

main().catch(e => { console.error(e); process.exit(1); });
