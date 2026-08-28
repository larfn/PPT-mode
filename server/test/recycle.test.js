// P2-F 回收站：删除模板 = 移入 .回收站（可恢复），彻底删除/清空不可恢复
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const {
  setTemplateRoot, saveTemplate, getTemplate, listTemplates, listFolders, deleteTemplate,
  listRecycleBin, restoreTemplate, purgeTemplate, emptyRecycleBin, RECYCLE_NAME, recycleDir
} = require('../src/templateStore.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-rc-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);

// 每个用例开始前清空回收站 + 模板库，保证相互独立
function cleanAll() {
  try { emptyRecycleBin(); } catch { /* ignore */ }
  for (const t of listTemplates()) { try { deleteTemplate(t.id, t.folder); } catch { /* ignore */ } }
  try { emptyRecycleBin(); } catch { /* ignore */ }
}

test('deleteTemplate moves template into .recycle bin (not permanent), with versions intact', () => {
  cleanAll();
  const saved = saveTemplate({ name: '回收测试A', template: { schemaVersion: 1, shapes: [] }, preview: 'data:image/png;base64,AAAA' });
  const id = saved.id;
  deleteTemplate(id);
  assert.equal(getTemplate(id), null, 'template no longer in library');
  assert.equal(fs.existsSync(path.join(tmp, id)), false, 'original dir moved away');
  assert.equal(fs.existsSync(path.join(recycleDir(), id)), true, 'moved into recycle bin');
  const items = listRecycleBin();
  assert.equal(items.length, 1);
  assert.equal(items[0].id, id);
  assert.equal(items[0].name, '回收测试A');
  assert.ok(items[0].deletedAt, 'deletedAt recorded');
  // 版本目录随模板一起进回收站
  assert.equal(fs.existsSync(path.join(recycleDir(), id, 'versions')), true);
  // 恢复后模板与版本完整回到原位置
  const restored = restoreTemplate(items[0].entryId);
  assert.equal(restored.id, id);
  assert.equal(getTemplate(id).name, '回收测试A');
  assert.equal(fs.existsSync(path.join(recycleDir(), id)), false, 'recycle entry removed after restore');
});

test('folder template keeps folder structure in recycle bin and restores to original folder', () => {
  cleanAll();
  const saved = saveTemplate({ name: '回收测试B', folder: '项目B', template: { schemaVersion: 1, shapes: [] } });
  deleteTemplate(saved.id, '项目B');
  const items = listRecycleBin();
  assert.equal(items.length, 1);
  assert.equal(items[0].folder, '项目B');
  assert.equal(fs.existsSync(path.join(recycleDir(), '项目B', saved.id)), true);
  // 列表不显示回收站里的模板/文件夹
  assert.equal(listTemplates().length, 0);
  assert.deepEqual(listFolders(), []);
  const restored = restoreTemplate(items[0].entryId);
  assert.equal(restored.folder, '项目B');
  assert.equal(getTemplate(saved.id, '项目B').name, '回收测试B');
});

test('restore rejects when target name already exists', () => {
  cleanAll();
  saveTemplate({ name: '冲突模板', template: { schemaVersion: 1, shapes: [] } });
  deleteTemplate('冲突模板'); // 移入回收站
  saveTemplate({ name: '冲突模板', template: { schemaVersion: 1, shapes: [] } }); // 原位置新模板
  const items = listRecycleBin();
  assert.equal(items.length, 1);
  assert.throws(() => restoreTemplate(items[0].entryId), /已存在同名模板/);
});

test('purgeTemplate removes permanently; emptyRecycleBin clears all', () => {
  cleanAll();
  saveTemplate({ name: '清除A', template: { schemaVersion: 1, shapes: [] } });
  saveTemplate({ name: '清除B', folder: '分类X', template: { schemaVersion: 1, shapes: [] } });
  deleteTemplate('清除A');
  deleteTemplate('清除B', '分类X');
  assert.equal(listRecycleBin().length, 2);
  const items = listRecycleBin();
  const r = purgeTemplate(items[0].entryId);
  assert.equal(r.ok, true);
  assert.equal(listRecycleBin().length, 1);
  const e = emptyRecycleBin();
  assert.equal(e.ok, true);
  assert.ok(e.removed >= 1, 'clears remaining template (+ possibly empty folder)');
  assert.equal(listRecycleBin().length, 0);
});

test('findTemplateDir global scan skips recycle bin (deleted templates not found by id fallback)', () => {
  cleanAll();
  const saved = saveTemplate({ name: '全局扫描回收站', template: { schemaVersion: 1, shapes: [] } });
  deleteTemplate(saved.id);
  assert.equal(getTemplate(saved.id), null, 'global fallback must not find deleted template in recycle bin');
});

test('traversal-safe entryId is rejected', () => {
  cleanAll();
  // 穿越条目解析失败 → 报错（不会操作回收站之外的路径）
  assert.throws(() => restoreTemplate('../escape'), /回收站条目不存在/);
  const r = purgeTemplate('..\\..\\escape');
  assert.equal(r.ok, false);
  assert.equal(fs.existsSync(path.join(os.tmpdir(), 'escape')), false);
  // 空 entryId 同样被拒
  assert.throws(() => restoreTemplate(''), /回收站条目不存在/);
});

// —— HTTP 端点 ——
test('HTTP recycle endpoints: list / restore / purge / empty', async () => {
  cleanAll();
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const post = await fetch(base + '/api/templates', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'HTTP回收站', template: { schemaVersion: 1, shapes: [] } })
    });
    assert.equal(post.status, 200);
    const list = await (await fetch(base + '/api/templates')).json();
    const item = list.find((t) => t.name === 'HTTP回收站');
    const del = await fetch(base + '/api/templates/' + encodeURIComponent(item.id), { method: 'DELETE' });
    assert.equal(del.status, 200);

    let rc = await (await fetch(base + '/api/templates/recycle')).json();
    assert.equal(rc.items.length, 1);
    const entryId = rc.items[0].entryId;

    const rest = await fetch(base + '/api/templates/recycle/restore', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ entryId })
    });
    assert.equal(rest.status, 200);
    rc = await (await fetch(base + '/api/templates/recycle')).json();
    assert.equal(rc.items.length, 0);

    await fetch(base + '/api/templates/' + encodeURIComponent(item.id), { method: 'DELETE' });
    rc = await (await fetch(base + '/api/templates/recycle')).json();
    const purge = await fetch(base + '/api/templates/recycle/purge', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ entryId: rc.items[0].entryId })
    });
    assert.equal(purge.status, 200);
    rc = await (await fetch(base + '/api/templates/recycle')).json();
    assert.equal(rc.items.length, 0);

    const empty = await fetch(base + '/api/templates/recycle', { method: 'DELETE' });
    assert.equal(empty.status, 200);
  } finally {
    server.close();
  }
});