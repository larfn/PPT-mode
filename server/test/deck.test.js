// 套版（Deck）测试：存储 CRUD / 模板引用校验 / 多页构建 / HTTP 全链路
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const JSZip = require('jszip');
const { createApp } = require('../src/index.js');
const { setTemplateRoot, saveTemplate, deleteTemplate, getTemplate } = require('../src/templateStore.js');
const { listDecks, saveDeck, getDeck, deleteDeck, listDeckRecycleBin, restoreDeck, purgeDeck } = require('../src/deckStore.js');
const { buildDeckBase64 } = require('../src/slideBuilder.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-deck-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);

const tpl = (name, w, h, tag) => ({
  schemaVersion: 1, name,
  slideSize: { width: w, height: h },
  shapes: [
    { id: 't0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 0.5, width: 8, height: 1 }, textStyle: { size: 32 }, prompt: '标题' },
    { id: 'img0', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 2, width: 4, height: 3 },
      imageStyle: tag === 'styled' ? { srcRect: { l: 100, t: 0, r: 0, b: 0 }, spPrXml: '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' } : undefined }
  ]
});

async function unzip(b64) { return JSZip.loadAsync(Buffer.from(b64, 'base64')); }

test('deckStore: save/list/get/delete roundtrip + template reference validation', () => {
  const cover = saveTemplate({ name: '封面', template: tpl('封面', 13.33, 7.5, 'plain') });
  const body = saveTemplate({ name: '正文', template: tpl('正文', 13.33, 7.5, 'plain') });
  const saved = saveDeck({ name: '周报', deck: { pages: [
    { templateId: cover.id, variables: { date: '2026-08-20' } },
    { templateId: body.id }
  ] } });
  assert.ok(saved.id);
  const list = listDecks();
  assert.equal(list.length, 1);
  assert.equal(list[0].pageCount, 2);
  const got = getDeck(saved.id);
  assert.equal(got.deck.pages.length, 2);
  assert.equal(got.deck.pages[0].templateId, cover.id);
  assert.ok(got.deck.createdAt && got.deck.updatedAt);
  // 引用了不存在的模板 → 拒绝
  assert.throws(() => saveDeck({ name: '坏套版', deck: { pages: [{ templateId: '不存在' }] } }), /模板不存在/);
  // 固定版本不存在 → 拒绝
  assert.throws(() => saveDeck({ name: '坏版本', deck: { pages: [{ templateId: cover.id, templateVersion: 'v99' }] } }), /版本不存在/);
  deleteDeck(saved.id);
  assert.equal(listDecks().length, 0);
  deleteTemplate(cover.id); deleteTemplate(body.id);
});

test('deckStore: folder + traversal safety', () => {
  saveTemplate({ name: '夹内', folder: '项目', template: tpl('夹内', 13.33, 7.5, 'plain') });
  saveDeck({ name: '套版A', folder: '项目', deck: { pages: [{ templateId: '夹内', templateFolder: '项目' }] } });
  assert.equal(getDeck('套版A', '项目').deck.pages.length, 1);
  assert.throws(() => saveDeck({ name: '..', deck: { pages: [{ templateId: '夹内', templateFolder: '项目' }] } }), /invalid deck id/);
  deleteDeck('套版A', '项目');
  deleteTemplate('夹内', '项目');
});

test('deckStore: delete moves deck into recycle bin and can restore or purge', () => {
  const t = saveTemplate({ name: '套版回收模板', template: tpl('套版回收模板', 13.33, 7.5, 'plain') });
  const saved = saveDeck({ name: '套版回收项', deck: { pages: [{ templateId: t.id }] } });
  assert.equal(listDecks().some((d) => d.id === saved.id), true);
  deleteDeck(saved.id, '');
  assert.equal(listDecks().some((d) => d.id === saved.id), false);
  let rc = listDeckRecycleBin();
  assert.equal(rc.some((d) => d.id === saved.id && d.pageCount === 1), true);
  restoreDeck(rc.find((d) => d.id === saved.id).entryId);
  assert.equal(listDecks().some((d) => d.id === saved.id), true);
  deleteDeck(saved.id, '');
  rc = listDeckRecycleBin();
  assert.equal(purgeDeck(rc.find((d) => d.id === saved.id).entryId).ok, true);
  assert.equal(listDeckRecycleBin().some((d) => d.id === saved.id), false);
  deleteTemplate(t.id);
});

test('buildDeckBase64: multi-page pptx with per-page styles (no regression to single page)', async () => {
  const t1 = tpl('p1', 13.33, 7.5, 'styled');
  const t2 = tpl('p2', 13.33, 7.5, 'plain');
  const b64 = await buildDeckBase64([
    { template: t1, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=', texts: { t0: '第一页标题' }, vars: {} },
    { template: t2, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=', texts: { t0: '第二页标题' }, vars: {} }
  ]);
  const zip = await unzip(b64);
  assert.ok(zip.file('ppt/slides/slide1.xml') && zip.file('ppt/slides/slide2.xml'), '应有 2 张 slide');
  const s1 = await zip.file('ppt/slides/slide1.xml').async('string');
  const s2 = await zip.file('ppt/slides/slide2.xml').async('string');
  assert.ok(s1.includes('第一页标题'));
  assert.ok(s2.includes('第二页标题'));
  // 图片样式按页应用：第 1 页 roundRect + srcRect，第 2 页无
  assert.ok(s1.includes('prst="roundRect"') && s1.includes('srcRect'), '第 1 页应有样式');
  assert.ok(!s2.includes('roundRect') && !s2.includes('srcRect'), '第 2 页无样式');
  // 每页都含图片
  assert.ok(s1.includes('<a:blip') && s2.includes('<a:blip'));
});

test('buildDeckBase64: inconsistent slideSize rejected; empty pages rejected', async () => {
  const t1 = tpl('a', 13.33, 7.5, 'plain');
  const t2 = tpl('b', 10, 5.625, 'plain');
  await assert.rejects(() => buildDeckBase64([
    { template: t1, texts: {} },
    { template: t2, texts: {} }
  ]), /尺寸不一致/);
  await assert.rejects(() => buildDeckBase64([]), /no pages/);
});

test('HTTP: deck CRUD + build end-to-end (2 pages, per-page isolation)', async () => {
  const cover = saveTemplate({ name: 'HTTP封面', template: tpl('HTTP封面', 13.33, 7.5, 'styled'), preview: 'data:image/png;base64,iVBORw0KGgo=' });
  const body = saveTemplate({ name: 'HTTP正文', template: tpl('HTTP正文', 13.33, 7.5, 'plain') });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    // 保存套版
    const post = await fetch(base + '/api/decks', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'HTTP套版', deck: { pages: [
        { templateId: cover.id, variables: {} },
        { templateId: body.id }
      ] } })
    });
    assert.equal(post.status, 200);
    // 列表
    const list = await (await fetch(base + '/api/decks')).json();
    assert.equal(list.length, 1);
    assert.equal(list[0].pageCount, 2);
    const preview = await fetch(base + list[0].preview);
    assert.equal(preview.status, 200, '套版无独立预览图时应回退显示第一张模板预览');
    // build：第二页模板不存在（构造失败隔离）
    const build = await fetch(base + '/api/decks/build', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ pages: [
        { templateId: cover.id, texts: { t0: '封面标题' }, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=' },
        { templateId: '不存在', texts: {} }
      ] })
    });
    const buildBody = await build.json();
    assert.equal(buildBody.ok, true, '有页面成功即整体 ok');
    assert.equal(buildBody.pageCount, 1);
    assert.equal(buildBody.pageResults[0].ok, true);
    assert.equal(buildBody.pageResults[1].ok, false, '失败页隔离');
    const zip = await unzip(buildBody.base64);
    assert.ok(zip.file('ppt/slides/slide1.xml'));
    // 删除套版（不影响模板）
    const del = await fetch(base + '/api/decks/' + encodeURIComponent(list[0].id), { method: 'DELETE' });
    assert.equal(del.status, 200);
    assert.equal((await (await fetch(base + '/api/decks')).json()).length, 0);
    assert.ok(getTemplate('HTTP封面'), '删除套版不影响模板');
  } finally {
    server.close();
    deleteTemplate('HTTP封面'); deleteTemplate('HTTP正文');
  }
});
