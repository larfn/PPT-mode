
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const { setTemplateRoot, saveTemplate, deleteTemplate } = require('../src/templateStore.js');
const { setPendingRoot, pendingDir } = require('../src/pendingStore.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);
const pendTmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-pend-'));
setPendingRoot(pendTmp);

test('POST /api/ai/generate → pending queue → write → delete roundtrip', async () => {
  const saved = saveTemplate({
    name: 'AI模板',
    template: {
      schemaVersion: 1, name: 'AI模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'shp0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { font: '微软雅黑', size: 24 } },
        { id: 'shp1', type: 'text', role: 'manual_var', varName: '页码', bounds: { left: 12, top: 7, width: 1, height: 0.5 }, textStyle: { size: 12 } }
      ]
    },
    preview: ''
  });

  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    // 生成
    const gen = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ templateId: saved.id, texts: { shp0: 'AI 生成的内容' }, vars: { shp1: '05' } })
    });
    assert.equal(gen.status, 200);
    const genBody = await gen.json();
    assert.ok(genBody.ok && genBody.pendingId);
    // 列表（无 base64）
    const list = await (await fetch(base + '/api/ai/pending')).json();
    assert.equal(list.length, 1);
    assert.equal(list[0].templateName, 'AI模板');
    assert.equal(list[0].base64, undefined, 'list should not include base64');
    // 详情（含 base64，可解出 pptx）
    const detail = await (await fetch(base + '/api/ai/pending/' + genBody.pendingId)).json();
    assert.ok(detail.base64);
    const buf = Buffer.from(detail.base64, 'base64');
    assert.ok(buf.toString().includes('PK'), 'base64 should decode to a zip');
    // 标记写入
    const w = await fetch(base + '/api/ai/pending/' + genBody.pendingId + '/write', { method: 'POST', body: '{}', headers: { 'content-type': 'application/json' } });
    assert.equal(w.status, 200);
    assert.equal((await w.json()).ok, true);
    const after = await (await fetch(base + '/api/ai/pending')).json();
    assert.equal(after[0].written, true);
    // 删除
    const del = await fetch(base + '/api/ai/pending/' + genBody.pendingId, { method: 'DELETE' });
    assert.equal(del.status, 200);
    assert.equal((await (await fetch(base + '/api/ai/pending')).json()).length, 0);
  } finally {
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('POST /api/ai/generate rejects unknown template', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const res = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ templateId: '不存在', texts: {} })
    });
    assert.equal(res.status, 404);
  } finally {
    server.close();
  }
});

// —— 模板语义层：generate 时应用语义约束 ——

test('POST /api/ai/generate applies maxChars/maxLines from template semantics and reports warnings', async () => {
  const saved = saveTemplate({
    name: '约束模板',
    template: {
      schemaVersion: 1, name: '约束模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'title1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, semanticRole: 'title', maxChars: 5, maxLines: 1 },
        { id: 'body1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 2, width: 5, height: 1 }, maxChars: 0, minChars: 10 }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const gen = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        templateId: saved.id,
        texts: { title1: '0123456789', body1: 'short' },
        vars: {}
      })
    });
    assert.equal(gen.status, 200);
    const body = await gen.json();
    assert.ok(Array.isArray(body.warnings) && body.warnings.length === 2, 'expected 2 warnings, got ' + JSON.stringify(body.warnings));
    // 队列里的文本已按约束修正
    const detail = await (await fetch(base + '/api/ai/pending/' + body.pendingId)).json();
    assert.equal(detail.texts.title1, '01234', 'truncated to maxChars=5');
    assert.equal(detail.texts.body1, 'short', 'maxChars=0 means no limit; minChars only warns');
  } finally {
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('POST /api/ai/generate with old-style template (no semantics) passes texts through untouched', async () => {
  const saved = saveTemplate({
    name: '旧约束模板',
    template: {
      schemaVersion: 1, name: '旧约束模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'shp0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, prompt: '任意长度' }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const longText = '很长很长'.repeat(30);
    const gen = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ templateId: saved.id, texts: { shp0: longText }, vars: {} })
    });
    const body = await gen.json();
    assert.equal(body.warnings, undefined, 'old templates must not produce warnings');
    const detail = await (await fetch(base + '/api/ai/pending/' + body.pendingId)).json();
    assert.equal(detail.texts.shp0, longText, 'text passes through unchanged');
  } finally {
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— P1-D 补全：generate 支持图片位 images（dataURL 直通）——

test('POST /api/ai/generate accepts images (dataURL) and stores them in pending entry', async () => {
  const saved = saveTemplate({
    name: 'AI图片模板',
    template: {
      schemaVersion: 1, name: 'AI图片模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'txt1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } },
        { id: 'img1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 3, width: 4, height: 3 } }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const gen = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ templateId: saved.id, texts: { txt1: '带图' }, images: { img1: png } })
    });
    assert.equal(gen.status, 200);
    const body = await gen.json();
    const detail = await (await fetch(base + '/api/ai/pending/' + body.pendingId)).json();
    assert.equal(detail.images.img1, png);
    // 生成的 pptx 内应包含 png 资源
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(Buffer.from(detail.base64, 'base64'));
    assert.ok(Object.keys(zip.files).some((n) => /media\/.*\.png$/i.test(n)), 'pptx should embed the png');
  } finally {
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('POST /api/ai/generate rejects non-dataURL non-http image values', async () => {
  const saved = saveTemplate({
    name: 'AI图片模板2',
    template: {
      schemaVersion: 1, name: 'AI图片模板2',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'img1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 3, width: 4, height: 3 } }]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const gen = await fetch(base + '/api/ai/generate', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ templateId: saved.id, texts: {}, images: { img1: 'C:\\foo\\bar.jpg' } })
    });
    assert.equal(gen.status, 400);
    const body = await gen.json();
    assert.ok(body.error.includes('dataURL'), 'should explain accepted formats: ' + body.error);
  } finally {
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— P1-D 补全：套版整份生成入待写队列 ——

test('POST /api/ai/generate-deck builds multi-page base64 into pending queue', async () => {
  const t1 = saveTemplate({
    name: 'AI套版页1',
    template: {
      schemaVersion: 1, name: 'AI套版页1',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'a1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } }]
    },
    preview: ''
  });
  const t2 = saveTemplate({
    name: 'AI套版页2',
    template: {
      schemaVersion: 1, name: 'AI套版页2',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'b1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } }]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const gen = await fetch(base + '/api/ai/generate-deck', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        name: 'AI测试套版',
        pages: [
          { templateId: t1.id, texts: { a1: '第一页' } },
          { templateId: t2.id, texts: { b1: '第二页' } }
        ]
      })
    });
    assert.equal(gen.status, 200);
    const body = await gen.json();
    assert.ok(body.ok && body.pendingId && body.pageCount === 2);
    const detail = await (await fetch(base + '/api/ai/pending/' + body.pendingId)).json();
    assert.equal(detail.deck, true);
    assert.equal(detail.pageCount, 2);
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(Buffer.from(detail.base64, 'base64'));
    const slides = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
    assert.ok(slides.length >= 2, 'deck base64 should contain 2+ slides, got: ' + slides.length);
  } finally {
    server.close();
    deleteTemplate(t1.id);
    deleteTemplate(t2.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('POST /api/ai/generate-deck fails gracefully when one page template is missing', async () => {
  const t1 = saveTemplate({
    name: 'AI套版好页',
    template: {
      schemaVersion: 1, name: 'AI套版好页',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'a1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 } }]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const gen = await fetch(base + '/api/ai/generate-deck', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        pages: [
          { templateId: t1.id, texts: { a1: '好页' } },
          { templateId: '不存在的模板', texts: {} }
        ]
      })
    });
    const body = await gen.json();
    assert.ok(body.ok === true, 'one bad page should not fail the whole deck');
    assert.equal(body.pageResults.filter((p) => p.ok).length, 1);
    assert.equal(body.pageResults.filter((p) => !p.ok).length, 1);
  } finally {
    server.close();
    deleteTemplate(t1.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});
