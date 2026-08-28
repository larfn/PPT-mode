const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const { listTemplates, listFolders, saveTemplate, getTemplate, deleteTemplate, setTemplateRoot } = require('../src/templateStore.js');



// —— 模板语义层 ——

test('semantic: save/get roundtrip keeps semantic fields (new template)', () => {
  const saved = saveTemplate({
    name: '语义模板',
    template: {
      schemaVersion: 1, name: '语义模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        {
          id: 'shp0', type: 'text', role: 'ai_text',
          bounds: { left: 1, top: 1, width: 5, height: 1 },
          textStyle: { size: 32 },
          semanticRole: 'title', contentType: '公司名', required: true,
          maxChars: 20, maxLines: 1, minChars: 2, preferredLength: 12,
          generationInstruction: '用一句话概括'
        },
        { id: 'shp1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 3, width: 4, height: 3 } },
        { id: 'shp2', type: 'text', role: 'fixed', content: 'x', bounds: { left: 1, top: 6, width: 1, height: 0.5 } }
      ]
    },
    preview: ''
  });
  const got = getTemplate(saved.id);
  const t = got.template;
  const a = t.shapes[0];
  assert.equal(a.semanticRole, 'title');
  assert.equal(a.contentType, '公司名');
  assert.equal(a.required, true);
  assert.equal(a.maxChars, 20);
  assert.equal(a.maxLines, 1);
  assert.equal(a.minChars, 2);
  assert.equal(a.preferredLength, 12);
  assert.equal(a.generationInstruction, '用一句话概括');
  // 非 AI 文本位不强制语义字段
  assert.equal(t.shapes[1].semanticRole, undefined);
  assert.equal(t.shapes[2].semanticRole, undefined);
  deleteTemplate(saved.id);
});

test('semantic: old template without semantic fields loads with defaults (no crash, no fields)', () => {
  const saved = saveTemplate({
    name: '旧模板',
    template: {
      schemaVersion: 1, name: '旧模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'shp0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 }, prompt: '写标题' },
        { id: 'shp1', type: 'text', role: 'manual_var', varName: '页码', bounds: { left: 12, top: 7, width: 1, height: 0.5 } }
      ]
    },
    preview: ''
  });
  const t = getTemplate(saved.id).template;
  assert.equal(t.shapes.length, 2);
  for (const s of t.shapes) {
    assert.equal(s.semanticRole, undefined);
    assert.equal(s.maxChars, undefined);
    assert.equal(s.maxLines, undefined);
    assert.equal(s.minChars, undefined);
    assert.equal(s.preferredLength, undefined);
    assert.equal(s.required, undefined);
    assert.equal(s.contentType, undefined);
    assert.equal(s.generationInstruction, undefined);
  }
  assert.equal(t.shapes[0].prompt, '写标题', 'existing fields untouched');
  assert.equal(t.shapes[1].varName, '页码');
  deleteTemplate(saved.id);
});

test('semantic: save cleans invalid enum / negative numbers / non-boolean required', () => {
  const saved = saveTemplate({
    name: '脏模板',
    template: {
      schemaVersion: 1, name: '脏模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 's', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 },
          semanticRole: 'not-a-role', maxChars: -3, maxLines: 0, required: 1 }
      ]
    },
    preview: ''
  });
  const s = getTemplate(saved.id).template.shapes[0];
  assert.equal(s.semanticRole, 'other', 'invalid enum → 不指定');
  assert.equal(s.maxChars, undefined);
  assert.equal(s.maxLines, 0, 'maxLines=0 is a valid "no limit" value and must be preserved');
  assert.equal(s.required, undefined);
  deleteTemplate(saved.id);
});

test('semantic: POST /api/templates then GET returns semantic fields (JSON roundtrip)', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const post = await fetch(base + '/api/templates', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        name: '语义API模板',
        template: {
          schemaVersion: 1, name: '语义API模板',
          slideSize: { width: 13.33, height: 7.5 },
          shapes: [
            { id: 's', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 },
              semanticRole: 'body', maxChars: 80, generationInstruction: '给出明确的结论' }
          ]
        },
        preview: ''
      })
    });
    assert.equal(post.status, 200);
    const list = await (await fetch(base + '/api/templates')).json();
    const item = list.find((t) => t.name === '语义API模板');
    const got = await (await fetch(base + '/api/templates/' + encodeURIComponent(item.id))).json();
    const s = got.template.shapes[0];
    assert.equal(s.semanticRole, 'body');
    assert.equal(s.maxChars, 80);
    assert.equal(s.generationInstruction, '给出明确的结论');
    deleteTemplate(item.id);
  } finally {
    server.close();
  }
});

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);

test('save/list/get/delete template roundtrip', () => {
  const saved = saveTemplate({ name: '测试模板', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,AAAA' });
  const list = listTemplates();
  assert.equal(list.length, 1);
  assert.equal(list[0].name, '测试模板');
  const got = getTemplate(saved.id);
  assert.equal(got.name, '测试模板');
  deleteTemplate(saved.id);
  assert.equal(listTemplates().length, 0);
});

test('POST /api/templates then GET /api/templates', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const post = await fetch(`${base}/api/templates`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ name: 'API模板', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,BBBB' })
  });
  assert.equal(post.status, 200);
  const get = await fetch(`${base}/api/templates`);
  const list = await get.json();
  assert.equal(list.length, 1);
  assert.equal(list[0].name, 'API模板');
  // 清理，避免污染后续文件夹测试
  deleteTemplate(list[0].id);
  server.close();
});

test('saveTemplate rejects traversal name and leaves root unchanged', () => {
  const before = fs.readdirSync(tmp);
  assert.throws(() => saveTemplate({ name: '..', template: { schemaVersion: 1, shapes: [] } }), /invalid template id/);
  assert.deepEqual(fs.readdirSync(tmp), before);
});

test('deleteTemplate("..") is a no-op', () => {
  const parent = path.dirname(tmp);
  const before = fs.readdirSync(tmp);
  deleteTemplate('..');
  assert.equal(fs.existsSync(tmp), true);
  assert.equal(fs.existsSync(parent), true);
  assert.deepEqual(fs.readdirSync(tmp), before);
});

test('saveTemplate skips traversal asset keys', () => {
  const saved = saveTemplate({
    name: 'traversal-assets',
    template: { schemaVersion: 1, shapes: [] },
    assets: { '../escape.png': 'data:image/png;base64,AAAA' }
  });
  const parent = path.dirname(tmp);
  assert.equal(fs.existsSync(path.join(parent, 'escape.png')), false);
  assert.equal(fs.existsSync(path.join(parent, 'escape.png.png')), false);
  assert.deepEqual(fs.readdirSync(path.join(tmp, saved.id, 'assets')), []);
  deleteTemplate(saved.id);
});

test('folder: save/list/get/delete with folder, root templates still listed', () => {
  const a = saveTemplate({ name: '根模板', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,AAAA' });
  const b = saveTemplate({ name: '项目A模板', folder: '项目A', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,BBBB' });
  const c = saveTemplate({ name: '项目A封面', folder: '项目A', template: { schemaVersion: 1, shapes: [] } });
  const list = listTemplates();
  assert.equal(list.length, 3);
  const rootItem = list.find((t) => t.id === a.id);
  const folderItem = list.find((t) => t.id === b.id);
  assert.equal(rootItem.folder, '');
  assert.equal(folderItem.folder, '项目A');
  assert.ok(folderItem.preview.includes('folder=') && folderItem.preview.includes('id='), 'preview url should carry folder+id');

  const folders = listFolders();
  assert.deepEqual(folders, [{ name: '项目A', count: 2 }]);

  // P1-3 修复：不带 folder 也能全局找到子文件夹里的模板（folder 仅作辅助，避免 AI/MCP 因漏传 folder 而 not found）
  assert.equal(getTemplate(b.id).name, '项目A模板', 'getTemplate without folder finds folder template (global fallback)');
  assert.equal(getTemplate(b.id, '项目A').name, '项目A模板');
  assert.equal(getTemplate(a.id).name, '根模板');

  deleteTemplate(b.id, '项目A');
  assert.equal(listTemplates().length, 2);
  assert.equal(listFolders()[0].count, 1);
  deleteTemplate(c.id, '项目A');
  assert.deepEqual(listFolders(), []);
  deleteTemplate(a.id);
  assert.equal(listTemplates().length, 0);
});

test('folder: traversal and separator chars are neutralized', () => {
  const parent = path.dirname(tmp);
  // folder '..' 被清洗为空 → 模板安全保存到根目录，绝不穿越
  const rootSaved = saveTemplate({ name: 'x', folder: '..', template: { schemaVersion: 1, shapes: [] } });
  assert.equal(fs.existsSync(path.join(tmp, rootSaved.id)), true);
  assert.equal(fs.existsSync(path.join(parent, 'x')), false);
  deleteTemplate(rootSaved.id);
  // folder '../evil' 清洗为 '.._evil'（斜杠变下划线），仍位于 tmp 内
  const escaped = saveTemplate({ name: 'y', folder: '../evil', template: { schemaVersion: 1, shapes: [] } });
  assert.equal(fs.existsSync(path.join(tmp, '.._evil', escaped.id)), true);
  assert.equal(fs.existsSync(path.join(parent, 'evil', escaped.id)), false);
  deleteTemplate(escaped.id, '../evil');
  // 只删模板目录；空的分类文件夹保留（可能还有其他模板）
  assert.equal(fs.existsSync(path.join(tmp, '.._evil', 'y')), false);
});

test('POST /api/templates with folder, GET /api/templates/folders, preview.png with query', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  const post = await fetch(baseUrl + '/api/templates', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ name: '文件夹模板', folder: '财务', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,CCCC' })
  });
  assert.equal(post.status, 200);
  const folders = await (await fetch(baseUrl + '/api/templates/folders')).json();
  assert.deepEqual(folders, [{ name: '财务', count: 1 }]);
  const list = await (await fetch(baseUrl + '/api/templates')).json();
  const item = list.find((t) => t.name === '文件夹模板');
  assert.equal(item.folder, '财务');
  const get = await fetch(baseUrl + '/api/templates/' + encodeURIComponent(item.id) + '?folder=' + encodeURIComponent('财务'));
  assert.equal(get.status, 200);
  const got = await get.json();
  assert.equal(got.name, '文件夹模板');
  // preview.png 带 query
  const png = await fetch(baseUrl + '/api/templates/preview.png?folder=' + encodeURIComponent('财务') + '&id=' + encodeURIComponent(item.id));
  assert.equal(png.status, 200);
  const buf = Buffer.from(await png.arrayBuffer());
  assert.ok(buf.toString('base64').startsWith('CCCC'), 'preview bytes should match uploaded data');
  // 删除（带 folder）
  const del = await fetch(baseUrl + '/api/templates/' + encodeURIComponent(item.id) + '?folder=' + encodeURIComponent('财务'), { method: 'DELETE' });
  assert.equal(del.status, 200);
  assert.deepEqual(await (await fetch(baseUrl + '/api/templates/folders')).json(), []);
  server.close();
});

