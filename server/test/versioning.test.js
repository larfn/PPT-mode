// 模板版本控制专项测试：v1/v2/v3 创建、restore、set-current、delete、迁移、损坏恢复、并发与边界
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const {
  setTemplateRoot, saveTemplate, getTemplate, deleteTemplate, listTemplates,
  listVersions, getVersion, restoreVersion, setCurrentVersion, deleteVersion
} = require('../src/templateStore.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ver-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);

const tpl = (name, tag) => ({
  schemaVersion: 1, name,
  slideSize: { width: 13.33, height: 7.5 },
  shapes: [
    { id: 's0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, semanticRole: 'title', prompt: tag + ' 提示词' }
  ]
});

test('保存同名模板创建 v1 → v2 → v3，绝不静默覆盖', () => {
  const a = saveTemplate({ name: '版本递增', template: tpl('版本递增', 'A'), preview: '' });
  assert.equal(a.version, 1); assert.equal(a.versionId, 'v1'); assert.equal(a.created, true);
  const b = saveTemplate({ name: '版本递增', template: tpl('版本递增', 'B'), preview: '' });
  assert.equal(b.version, 2); assert.equal(b.versionId, 'v2');
  const c = saveTemplate({ name: '版本递增', template: tpl('版本递增', 'C'), preview: '' });
  assert.equal(c.version, 3); assert.equal(c.versionId, 'v3');
  // v1 内容未被覆盖：v1 快照还是 A
  const v1 = getVersion('版本递增', '', 'v1');
  assert.equal(v1.version.shapes[0].prompt, 'A 提示词');
  assert.equal(v1.isCurrent, false);
  const v2 = getVersion('版本递增', '', 'v2');
  assert.equal(v2.version.shapes[0].prompt, 'B 提示词');
  // 当前 = v3
  const cur = getTemplate('版本递增').template;
  assert.equal(cur.version, 3);
  assert.equal(cur.versionId, 'v3');
  assert.equal(cur.shapes[0].prompt, 'C 提示词');
  deleteTemplate('版本递增');
});

test('listVersions 返回全部版本、顺序、字段与 isCurrent', () => {
  saveTemplate({ name: '版本列表', template: tpl('版本列表', '1') });
  saveTemplate({ name: '版本列表', template: tpl('版本列表', '2'), changeNote: '第二版说明' });
  const r = listVersions('版本列表');
  assert.equal(r.currentVersion, 2);
  assert.equal(r.currentVersionId, 'v2');
  assert.equal(r.versions.length, 2);
  assert.equal(r.versions[0].versionId, 'v1');
  assert.equal(r.versions[0].isCurrent, false);
  assert.equal(r.versions[1].versionId, 'v2');
  assert.equal(r.versions[1].isCurrent, true);
  assert.equal(r.versions[1].changeNote, '第二版说明');
  assert.ok(r.versions[0].createdAt && r.versions[1].updatedAt, '时间戳必须存在');
  deleteTemplate('版本列表');
});

test('restore v1：当前内容与版本元数据回到 v1；preview 同步不错配', () => {
  saveTemplate({ name: '恢复测试', template: tpl('恢复测试', 'A'), preview: 'data:image/png;base64,QUFBQQ==' });
  saveTemplate({ name: '恢复测试', template: tpl('恢复测试', 'B'), preview: 'data:image/png;base64,QkJCQg==' });
  const r = restoreVersion('恢复测试', '', 'v1');
  assert.equal(r.ok, true); assert.equal(r.version, 1);
  const cur = getTemplate('恢复测试').template;
  assert.equal(cur.version, 1); assert.equal(cur.versionId, 'v1');
  assert.equal(cur.shapes[0].prompt, 'A 提示词');
  // 当前预览恢复为 v1 的预览
  const curPreview = fs.readFileSync(path.join(tmp, '恢复测试', 'preview.png'));
  assert.equal(curPreview.toString('base64'), 'QUFBQQ==');
  // 版本文件未被破坏
  assert.equal(getVersion('恢复测试', '', 'v2').version.shapes[0].prompt, 'B 提示词');
  deleteTemplate('恢复测试');
});

test('set-current 与 restore 等价（内容+标记一致）', () => {
  saveTemplate({ name: '设当前', template: tpl('设当前', 'A') });
  saveTemplate({ name: '设当前', template: tpl('设当前', 'B') });
  const r = setCurrentVersion('设当前', '', 'v1');
  assert.equal(r.ok, true);
  const cur = getTemplate('设当前').template;
  assert.equal(cur.versionId, 'v1');
  assert.equal(cur.shapes[0].prompt, 'A 提示词');
  deleteTemplate('设当前');
});

test('删除版本：非当前可删；当前被拒；唯一版本被拒；版本与模板删除严格区分', () => {
  const a = saveTemplate({ name: '删除版本', template: tpl('删除版本', 'A') });
  saveTemplate({ name: '删除版本', template: tpl('删除版本', 'B') });
  saveTemplate({ name: '删除版本', template: tpl('删除版本', 'C') });
  // 删非当前 v1 → ok，目录保留
  assert.equal(deleteVersion('删除版本', '', 'v1').ok, true);
  assert.ok(fs.existsSync(path.join(tmp, '删除版本', 'template.json')), '模板目录必须保留');
  assert.equal(listVersions('删除版本').versions.length, 2);
  // 删当前 v3 → 拒绝
  assert.throws(() => deleteVersion('删除版本', '', 'v3'), /不能删除当前版本/);
  // 切到 v2 后删 v3 → ok
  restoreVersion('删除版本', '', 'v2');
  assert.equal(deleteVersion('删除版本', '', 'v3').ok, true);
  // 唯一版本 v2 → 拒绝
  assert.throws(() => deleteVersion('删除版本', '', 'v2'), /至少保留一个版本/);
  // 删除整个模板（与删版本分开的接口）
  deleteTemplate('删除版本');
  assert.equal(fs.existsSync(path.join(tmp, '删除版本')), false);
  assert.equal(listTemplates().length, 0);
  // 非法版本号一律拒绝
  saveTemplate({ name: '非法版本', template: tpl('非法版本', 'x') });
  for (const bad of ['v0', 'v-1', '..', 'abc', 'v1/..', '']) {
    assert.throws(() => deleteVersion('非法版本', '', bad), /非法版本号|版本不存在/);
    assert.throws(() => restoreVersion('非法版本', '', bad), /非法版本号|版本不存在/);
  }
  deleteTemplate('非法版本');
});

test('旧模板自动迁移：无版本结构 → v1/current；preview 副本；幂等；迁移后保存 → v2', () => {
  const dir = path.join(tmp, '旧模板迁移');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'template.json'), JSON.stringify({
    schemaVersion: 1, name: '旧模板迁移',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [{ id: 's0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, prompt: '旧提示词' }]
  }));
  fs.writeFileSync(path.join(dir, 'preview.jpg'), Buffer.from('PREVIEW1', 'utf8'));
  // 读取触发迁移
  const t = getTemplate('旧模板迁移').template;
  assert.equal(t.version, 1); assert.equal(t.versionId, 'v1');
  assert.equal(t.shapes[0].prompt, '旧提示词');
  // 版本快照 + 预览副本已生成
  assert.ok(fs.existsSync(path.join(dir, 'versions', 'v1.json')));
  assert.ok(fs.existsSync(path.join(dir, 'versions', 'v1.jpg')));
  // listVersions 显示 v1/current
  const lv = listVersions('旧模板迁移');
  assert.equal(lv.versions.length, 1);
  assert.equal(lv.versions[0].isCurrent, true);
  // 幂等：再次读取不重复迁移、不新增版本
  getTemplate('旧模板迁移');
  assert.equal(listVersions('旧模板迁移').versions.length, 1);
  // 迁移后保存 → v2（不是覆盖 v1）
  const s = saveTemplate({ name: '旧模板迁移', template: tpl('旧模板迁移', '新内容') });
  assert.equal(s.version, 2);
  assert.equal(getVersion('旧模板迁移', '', 'v1').version.shapes[0].prompt, '旧提示词');
  deleteTemplate('旧模板迁移');
});

test('template.json 损坏自动从版本恢复（不丢失数据、不报错）', () => {
  const a = saveTemplate({ name: '损坏恢复', template: tpl('损坏恢复', 'A') });
  saveTemplate({ name: '损坏恢复', template: tpl('损坏恢复', 'B') });
  // 写坏 template.json（模拟异常关机/写入中断）
  fs.writeFileSync(path.join(tmp, '损坏恢复', 'template.json'), '{ 坏掉的 json !!!');
  // listTemplates 跳过损坏模板（不崩溃）
  assert.ok(Array.isArray(listTemplates()));
  // getTemplate 自动恢复：从最新版本 v2 重建
  const t = getTemplate('损坏恢复');
  assert.ok(t, '损坏后应能恢复');
  assert.equal(t.template.version, 2);
  assert.equal(t.template.shapes[0].prompt, 'B 提示词');
  // 磁盘上的 template.json 已被修复为合法 JSON
  const raw = JSON.parse(fs.readFileSync(path.join(tmp, '损坏恢复', 'template.json'), 'utf8'));
  assert.equal(raw.versionId, 'v2');
  deleteTemplate('损坏恢复');
});

test('版本号冲突：版本文件被占用时自动分配到下一个可用版本号', () => {
  const a = saveTemplate({ name: '版本冲突', template: tpl('版本冲突', 'A') });
  assert.equal(a.version, 1);
  // 人为占用 v2（模拟跨进程并发已写入）
  fs.mkdirSync(path.join(tmp, '版本冲突', 'versions'), { recursive: true });
  fs.writeFileSync(path.join(tmp, '版本冲突', 'versions', 'v2.json'), JSON.stringify({ occupied: true }));
  const b = saveTemplate({ name: '版本冲突', template: tpl('版本冲突', 'B') });
  assert.equal(b.version, 3, 'v2 被占用时应自动分配 v3');
  assert.equal(getVersion('版本冲突', '', 'v3').version.shapes[0].prompt, 'B 提示词');
  deleteTemplate('版本冲突');
});

test('updateCurrent 修正当前版本（不新建版本），且保留原 changeNote', () => {
  const a = saveTemplate({ name: '补存修正', template: tpl('补存修正', 'A'), changeNote: '第一版' });
  assert.equal(a.version, 1);
  const b = saveTemplate({ name: '补存修正', template: tpl('补存修正', 'A-样式回读'), updateCurrent: true });
  assert.equal(b.version, 1, 'updateCurrent 不新建版本');
  assert.equal(b.created, false);
  assert.equal(listVersions('补存修正').versions.length, 1);
  const cur = getTemplate('补存修正').template;
  assert.equal(cur.shapes[0].prompt, 'A-样式回读 提示词');
  assert.equal(cur.changeNote, '第一版', '补存不能覆盖原版本说明');
  deleteTemplate('补存修正');
});

test('原子写入：保存后目录无 .tmp 残留；非法模板名被拒', () => {
  saveTemplate({ name: '原子写入', template: tpl('原子写入', 'x'), preview: 'data:image/png;base64,QUFBQQ==' });
  const entries = fs.readdirSync(path.join(tmp, '原子写入'));
  assert.ok(!entries.some((f) => f.includes('.tmp-')), '不应残留临时文件: ' + entries.join(','));
  const ventries = fs.readdirSync(path.join(tmp, '原子写入', 'versions'));
  assert.ok(!ventries.some((f) => f.includes('.tmp-')), 'versions 目录不应残留临时文件');
  assert.throws(() => saveTemplate({ name: '..', template: tpl('x', 'x') }), /invalid template id/);
  // 非法字符被安全清洗而不是拒绝（既有安全行为）：a/b → a_b
  const cleaned = saveTemplate({ name: 'a/b', template: tpl('x', 'x') });
  assert.ok(fs.existsSync(path.join(tmp, 'a_b')), '斜杠应被清洗为下划线');
  deleteTemplate(cleaned.id);
  deleteTemplate('原子写入');
});

test('folder 场景：版本操作在分类文件夹内正常', () => {
  saveTemplate({ name: '夹内模板', folder: '财务', template: tpl('夹内模板', 'A') });
  saveTemplate({ name: '夹内模板', folder: '财务', template: tpl('夹内模板', 'B') });
  const lv = listVersions('夹内模板', '财务');
  assert.equal(lv.versions.length, 2);
  restoreVersion('夹内模板', '财务', 'v1');
  assert.equal(getTemplate('夹内模板', '财务').template.versionId, 'v1');
  assert.throws(() => deleteVersion('夹内模板', '财务', 'v1'), /不能删除当前版本/);
  deleteTemplate('夹内模板', '财务');
});

test('HTTP: 版本端点全链路（保存/列表/详情/恢复/删除/非法参数）', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    // 保存 v1 + v2（带 changeNote）
    const p1 = await fetch(base + '/api/templates', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'HTTP版本', template: tpl('HTTP版本', 'A'), preview: '' })
    });
    const b1 = await p1.json();
    assert.equal(b1.version, 1);
    const p2 = await fetch(base + '/api/templates', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'HTTP版本', template: tpl('HTTP版本', 'B'), preview: '', changeNote: '第二版' })
    });
    const b2 = await p2.json();
    assert.equal(b2.version, 2);
    // 版本列表
    const lv = await (await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions')).json();
    assert.equal(lv.versions.length, 2);
    assert.equal(lv.versions[1].changeNote, '第二版');
    // 版本详情
    const gv = await (await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions/v1')).json();
    assert.equal(gv.version.shapes[0].prompt, 'A 提示词');
    // 恢复 v1
    const rv = await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/restore', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ versionId: 'v1' })
    });
    assert.equal(rv.status, 200);
    const cur = await (await fetch(base + '/api/templates/' + encodeURIComponent(b1.id))).json();
    assert.equal(cur.template.versionId, 'v1');
    // 删除当前版本 → 400
    const delCur = await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions/v1', { method: 'DELETE' });
    assert.equal(delCur.status, 400);
    // 切到 v2 后删除 v1 → 200
    await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/set-current', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ versionId: 'v2' })
    });
    const del = await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions/v1', { method: 'DELETE' });
    assert.equal(del.status, 200);
    const lv2 = await (await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions')).json();
    assert.equal(lv2.versions.length, 1);
    // 非法 versionId → 404
    const bad = await fetch(base + '/api/templates/' + encodeURIComponent(b1.id) + '/versions/v999');
    assert.equal(bad.status, 404);
    // 不存在的模板 → 404
    const nf = await fetch(base + '/api/templates/' + encodeURIComponent('不存在的') + '/versions');
    assert.equal(nf.status, 404);
    deleteTemplate('HTTP版本');
  } finally {
    server.close();
  }
});
