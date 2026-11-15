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

    console.log('All auth tests passed');
  } finally {
    server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

main().catch(e => { console.error(e); process.exit(1); });
