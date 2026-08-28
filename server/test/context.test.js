// 当前演示文稿上下文接口测试：路由参数 / 结构清洗 / 错误路径（PowerPoint 未运行）
const { test } = require('node:test');
const assert = require('node:assert');
const { createApp } = require('../src/index.js');
const { setPsRunner, sanitize } = require('../src/pptContext.js');

let lastArgs = null; // mock 记录收到的参数

function withServer(fn) {
  return async () => {
    const app = createApp();
    const server = app.listen(0, '127.0.0.1');
    await new Promise((resolve) => server.once('listening', resolve));
    const base = 'http://127.0.0.1:' + server.address().port;
    try { await fn(base); } finally { server.close(); }
  };
}

test('sanitize: PS5.1 null arrays and null values are cleaned, nested depth capped', () => {
  assert.deepEqual(sanitize({ a: null, b: 1, c: [1, null, { x: null, y: 2 }] }), { b: 1, c: [1, { y: 2 }] });
  assert.deepEqual(sanitize([null, null]), []);
  assert.deepEqual(sanitize(null), null);
  assert.deepEqual(sanitize({ deep: { deep: { deep: {} } } }), { deep: { deep: { deep: {} } } });
});

test('GET /api/context/presentation returns structured summary via runner', withServer(async (base) => {
  setPsRunner(async (args) => {
    lastArgs = args;
    return {
      ok: true, kind: 'presentation', name: '季度汇报.pptx', saved: true,
      slideCount: 3, slideSize: { width: 13.33, height: 7.5 },
      currentSlide: { index: 2, slideId: 256, selectionType: 2 },
      slides: [
        { index: 1, slideId: 256, shapeCount: 2 },
        { index: 2, slideId: 257, shapeCount: 5 },
        { index: 3, slideId: 258, shapeCount: 0 }
      ]
    };
  });
  const res = await fetch(base + '/api/context/presentation');
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(body.ok, true);
  assert.equal(body.name, '季度汇报.pptx');
  assert.equal(body.slideCount, 3);
  assert.equal(body.currentSlide.slideId, 256);
  assert.equal(body.slides.length, 3);
  assert.equal(body.slides[2].shapeCount, 0);
  assert.ok(lastArgs.includes('-Action') && lastArgs.includes('presentation'), 'runner should receive -Action presentation');
}));

test('GET /api/context/current-slide returns slide structure', withServer(async (base) => {
  setPsRunner(async (args) => ({ ok: true, kind: 'current-slide', index: 1, slideId: 256, shapeCount: 1, shapes: [{ id: 11, name: '标题 1', type: 'textBox', text: '标题文字' }] }));
  const body = await (await fetch(base + '/api/context/current-slide')).json();
  assert.equal(body.kind, 'current-slide');
  assert.equal(body.shapes[0].name, '标题 1');
}));

test('GET /api/context/slide?index=2 forwards index; ?id=300 forwards id', withServer(async (base) => {
  const calls = [];
  setPsRunner(async (args) => {
    calls.push(args);
    return { ok: true, kind: 'slide', index: 2, slideId: 257, shapeCount: 1, shapes: [] };
  });
  const byIndex = await (await fetch(base + '/api/context/slide?index=2')).json();
  assert.equal(byIndex.ok, true);
  assert.ok(calls[0].includes('-Index') && calls[0].includes('2'));
  const byId = await (await fetch(base + '/api/context/slide?id=300')).json();
  assert.equal(byId.ok, true);
  assert.ok(calls[1].includes('-Id') && calls[1].includes('300'));
}));

test('GET /api/context/slide without index/id returns friendly error, runner not called', withServer(async (base) => {
  let called = false;
  setPsRunner(async () => { called = true; return { ok: true }; });
  const res = await fetch(base + '/api/context/slide');
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(body.ok, false);
  assert.ok(body.error.includes('index'), 'error should mention index: ' + body.error);
  assert.equal(called, false, 'runner must not be called without params');
}));

test('GET /api/context/inspect without params defaults to current slide (no -Index/-Id)', withServer(async (base) => {
  setPsRunner(async (args) => {
    lastArgs = args;
    return { ok: true, kind: 'inspect', index: 1, slideId: 256, shapeCount: 2, shapes: [{ id: 1, name: 'a', type: 'textBox', text: 'x', left: 1 }] };
  });
  const body = await (await fetch(base + '/api/context/inspect')).json();
  assert.equal(body.kind, 'inspect');
  assert.ok(!lastArgs.includes('-Index') && !lastArgs.includes('-Id'), 'inspect default should not pass index/id');
}));

test('GET /api/context/inspect?index=3 forwards index', withServer(async (base) => {
  setPsRunner(async (args) => {
    lastArgs = args;
    return { ok: true, kind: 'inspect', index: 3, slideId: 259, shapes: [] };
  });
  await (await fetch(base + '/api/context/inspect?index=3')).json();
  assert.ok(lastArgs.includes('-Index') && lastArgs.includes('3'));
}));

test('context routes never leak raw pptx base64 (shape payload is structured only)', withServer(async (base) => {
  setPsRunner(async () => ({ ok: true, name: 't.pptx', shapes: [{ id: 1, name: 'n', type: 'textBox', text: 'hi' }] }));
  const body = await (await fetch(base + '/api/context/presentation')).json();
  const raw = JSON.stringify(body);
  assert.ok(!raw.includes('base64') && !raw.includes('PK\u0003'), 'no binary/base64 should appear');
}));

test('real PowerShell path: PowerPoint absent or no document yields friendly structured error', withServer(async (base) => {
  setPsRunner(null); // 真实 spawn powershell.exe → GetActiveObject
  const res = await fetch(base + '/api/context/presentation');
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(typeof body.ok, 'boolean');
  if (body.ok === false) {
    assert.ok(typeof body.error === 'string' && body.error.length > 0, 'error message expected');
  } else {
    assert.equal(body.kind, 'presentation');
  }
}));
