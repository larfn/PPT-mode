const { test } = require('node:test');
const assert = require('node:assert');
const { createApp, shouldWarnPortFallback } = require('../src/index.js');
const { VERSION } = require('../src/version.js');

test('GET /api/health returns ok and the running version', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  try {
    await new Promise((resolve) => server.once('listening', resolve));
    const port = server.address().port;
    const res = await fetch(`http://127.0.0.1:${port}/api/health`);
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.deepEqual(body, { ok: true, version: VERSION });
  } finally {
    await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  }
});

test('explicit non-default port does not report the default port as occupied', () => {
  assert.equal(shouldWarnPortFallback(3791, 3791), false);
  assert.equal(shouldWarnPortFallback(3792, 3791), true);
});
