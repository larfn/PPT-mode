'use strict';
// ===== 黄金路径 E2E（发布门禁）=====
// 6 条路径：单页模板 / 多图片 / 表格(CSV+扩展) / 套版 / 旧模板兼容 / 性能冒烟（大负载生成 < 预算）。
// 文件级断言必跑；PowerPoint COM 可用时自动追加「打开检查」（隐藏实例打开生成的 pptx 验证结构）。
//
// 运行：node server/e2e/golden-path.js [--no-com] [--verbose]
// 退出码：0 = 全部通过（含 SKIP 项）；1 = 存在失败项。
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');
const { spawn } = require('node:child_process');
const JSZip = require('../node_modules/jszip');
const pptxgen = require('../node_modules/pptxgenjs');
const { buildSlideBase64, buildDeckBase64 } = require('../src/slideBuilder.js');
const { parseSlideShapesDetailed } = require('../src/readStyles.js');
const { setTemplateRoot, saveTemplate, getTemplate } = require('../src/templateStore.js');
const { saveDeck, listDecks, getDeck } = require('../src/deckStore.js');
const { analyzeShapesByRules, HIGH_CONFIDENCE } = require('../../addin/src/lib/analyze.ts');
const { fitCrop, cropRectFromState, outputSize } = require('../../addin/src/lib/cropMath.ts');

const OPEN_CHECK = path.join(__dirname, 'open-check.ps1');
const ARGS = new Set(process.argv.slice(2));
const WANT_COM = !ARGS.has('--no-com');

// ---------- 环境隔离（临时模板库，不碰用户数据） ----------
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-e2e-'));
const TPL_ROOT = path.join(TMP, 'templates');
fs.mkdirSync(TPL_ROOT, { recursive: true });
setTemplateRoot(TPL_ROOT);

// ---------- 小工具 ----------
function makePng(r, g, b) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(1, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = 8; ihdr[9] = 2;
  const idat = zlib.deflateSync(Buffer.from([0, r, g, b]));
  function chunk(type, data) {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
    const t = Buffer.from(type, 'ascii');
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(Buffer.concat([t, data])) >>> 0, 0);
    return Buffer.concat([len, t, data, crc]);
  }
  const buf = Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', Buffer.alloc(0))]);
  return 'data:image/png;base64,' + buf.toString('base64');
}
const PNG_RED = makePng(200, 40, 40);
const PNG_BLUE = makePng(40, 80, 200);

// CSV → 二维数组（支持引号包裹、逗号/换行分隔；模拟前端「粘贴 CSV 导入」）
function csvToTable(csv) {
  const rows = []; let row = []; let cell = ''; let inQ = false;
  for (let i = 0; i < csv.length; i++) {
    const c = csv[i];
    if (inQ) {
      if (c === '"') { if (csv[i + 1] === '"') { cell += '"'; i++; } else inQ = false; }
      else cell += c;
    } else if (c === '"') inQ = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && csv[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
}

async function loadZip(base64) { return JSZip.loadAsync(Buffer.from(base64, 'base64')); }
function assert(cond, msg) { if (!cond) throw new Error(msg); }

// pptx 文件级探测：页数 / media 数 / 每页表格行数列数 / 每页文字 / pic 数
async function probePptx(base64) {
  const zip = await loadZip(base64);
  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
  const media = zip.file(/^ppt\/media\//).map((f) => f.name);
  const pages = [];
  for (const s of slideFiles) {
    const xml = await zip.file(s).async('string');
    const texts = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]);
    pages.push({
      slide: s,
      texts,
      tables: (xml.match(/<a:tbl>/g) || []).length,
      trCount: (xml.match(/<a:tr\b/g) || []).length,
      colCount: (xml.match(/<a:gridCol\b/g) || []).length,
      picCount: (xml.match(/<p:pic\b/g) || []).length
    });
  }
  return { slideCount: slideFiles.length, media, pages };
}

// ---------- PowerPoint COM 打开检查（隐藏实例） ----------
let comMode = WANT_COM ? 'probe' : 'off'; // probe → ok | unavailable
async function comOpen(base64, expected) {
  if (comMode === 'off') return { skipped: true, reason: '--no-com' };
  if (comMode === 'unavailable') return { skipped: true, reason: 'PowerPoint COM 不可用' };
  const pptxPath = path.join(TMP, 'chk-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7) + '.pptx');
  const outFile = path.join(TMP, 'chk-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7) + '.json');
  fs.writeFileSync(pptxPath, Buffer.from(base64, 'base64'));
  const res = await new Promise((resolve) => {
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', OPEN_CHECK, '-PptxPath', pptxPath, '-OutFile', outFile], { windowsHide: true });
    let err = '';
    child.stderr.on('data', (d) => { err += d; });
    child.on('close', (code) => {
      let data = null;
      try { data = JSON.parse(fs.readFileSync(outFile, 'utf8').replace(/^\uFEFF/, '')); } catch { /* no json */ }
      try { fs.rmSync(outFile, { force: true }); } catch { /* ignore */ }
      try { fs.rmSync(pptxPath, { force: true }); } catch { /* ignore */ }
      resolve({ code, data, err });
    });
  });
  if (res.code === 2) { comMode = 'unavailable'; return { skipped: true, reason: String(res.err || '').trim().slice(0, 120) }; }
  assert(res.code === 0 && res.data, 'COM 打开检查失败（exit ' + res.code + '）：' + String(res.err || '无输出').trim().slice(0, 200));
  const d = res.data;
  const checks = [];
  if (expected.slides !== undefined) { assert(d.slides === expected.slides, 'COM 页数=' + d.slides + '，期望 ' + expected.slides); checks.push('slides=' + d.slides); }
  if (expected.pictures !== undefined) {
    const pics = d.shapes.reduce((a, s) => a + s.pictures, 0);
    assert(pics === expected.pictures, 'COM 图片数=' + pics + '，期望 ' + expected.pictures); checks.push('pictures=' + pics);
  }
  if (expected.tables !== undefined) {
    const t = d.shapes.reduce((a, s) => a + s.tables, 0);
    assert(t === expected.tables, 'COM 表格数=' + t + '，期望 ' + expected.tables); checks.push('tables=' + t);
  }
  if (expected.containsTexts) {
    const all = d.shapes.flatMap((s) => s.texts || []).join('\n');
    for (const txt of expected.containsTexts) assert(all.includes(txt), 'COM 文本缺失: ' + txt);
    checks.push('texts ok');
  }
  return { ok: true, detail: checks.join(', ') };
}

// ---------- 测试运行器 ----------
const results = [];
async function runTest(name, fn) {
  const t0 = Date.now();
  try {
    const info = await fn();
    results.push({ name, ok: true, info: info || '', ms: Date.now() - t0 });
  } catch (e) {
    results.push({ name, ok: false, info: e.message, ms: Date.now() - t0 });
  }
}

// ---------- fixture：模拟用户设计的「模板页」 ----------
async function makeFixturePptx() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'T', width: 13.33, height: 7.5 });
  pptx.layout = 'T';
  const s = pptx.addSlide();
  s.addText('季度业绩汇报', { x: 1, y: 0.5, w: 10, h: 1, fontSize: 32, bold: true, fontFace: 'Microsoft YaHei' });
  s.addText('这是正文内容，用于验证黄金路径。', { x: 1, y: 2, w: 10, h: 1.5, fontSize: 14, fontFace: 'Microsoft YaHei' });
  s.addImage({ data: PNG_RED, x: 1, y: 4.5, w: 4, h: 2.5 });
  return pptx.write('base64');
}

// 模拟「读取当前页」：XML 解析文本框 + pic（EMU→英寸）；去重 sp/pic 重叠
function parsePageShapes(xml) {
  const shapes = parseSlideShapesDetailed(xml).map((b, i) => {
    const firstRun = (b.paragraphs || []).find((p) => p.runs && p.runs.length);
    const r0 = firstRun && firstRun.runs[0];
    return {
      id: 'shp' + i, name: b.name || ('shp' + i), type: 'text', source: 'slide',
      hasText: !!((b.paragraphs || []).some((p) => p.text && String(p.text).trim())),
      text: (b.paragraphs || []).map((p) => p.text || '').join('\n'),
      bounds: { left: b.left, top: b.top, width: b.width, height: b.height },
      textStyle: { size: (r0 && r0.sz) || 18, bold: !!(r0 && r0.b) }
    };
  });
  const pics = [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map((m, i) => {
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(m[0]);
    const ext = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(m[0]);
    const emu = (n) => (n === undefined ? 0 : Number(n) / 914400);
    return {
      id: 'pic' + i, name: 'Image ' + i, type: 'picture', source: 'slide', hasText: false, text: '',
      bounds: off && ext
        ? { left: emu(off[1]), top: emu(off[2]), width: emu(ext[1]), height: emu(ext[2]) }
        : { left: 0, top: 0, width: 1, height: 1 }
    };
  });
  // 去重：pptxgenjs 的图片同时被 <p:sp> 与 <p:pic> 捕获（真实 Office.js 链路只读一次），
  // 无文本且与 pic 位置重叠的 sp 视为同一图片 → 丢弃 sp 保留 pic。
  const overlap = (a, b) =>
    Math.abs(a.left - b.left) < 0.2 && Math.abs(a.top - b.top) < 0.2 &&
    Math.abs(a.width - b.width) < 0.2 && Math.abs(a.height - b.height) < 0.2;
  const filtered = shapes.filter((s) => !(String(s.text || '').trim() === '' && pics.some((p) => overlap(s.bounds, p.bounds))));
  return filtered.concat(pics);
}

const tests = [];

// ---------- 测试 1：单页模板（读取→标记→保存→重读→文字→图片→裁剪→生成→COM 检查） ----------
tests.push(['1/6 单页模板（读取→标记→保存→重读→生成→COM）', async () => {
  const PAGE = { width: 13.33, height: 7.5 };
  // 1. 读取当前页（fixture pptx 模拟用户设计的模板页）
  const fixtureB64 = await makeFixturePptx();
  const zip = await loadZip(fixtureB64);
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  const pageShapes = parsePageShapes(xml);
  assert(pageShapes.length >= 3, '读取当前页形状数=' + pageShapes.length + '，期望 >= 3（标题+正文+图片）');
  // 2. 自动标记（规则分类器，免费本地不联网）
  const recs = analyzeShapesByRules(pageShapes, PAGE);
  const applied = pageShapes.map((sh, i) => {
    const rec = recs.find((r) => r.idx === i);
    const role = rec && rec.confidence >= HIGH_CONFIDENCE ? rec.recommendedRole : 'fixed';
    return { shape: sh, role };
  });
  const imgApplied = applied.find((a) => a.shape.type === 'picture');
  assert(!!imgApplied && (imgApplied.role === 'ai_image' || imgApplied.role === 'fixed'), '图片位未获得推荐角色');
  assert(applied.filter((a) => a.role === 'ai_text').length >= 1, '标题/正文未标记为 ai_text');
  // 3. 保存模板（图片位 → ai_image + prompt；文本框 → ai_text / fixed content）
  const template = {
    schemaVersion: 1, name: '黄金路径-单页', slideSize: PAGE,
    shapes: applied.map((a) => {
      const sh = a.shape;
      if (sh.type === 'picture') return { id: sh.id, type: 'picture', role: a.role === 'ai_image' ? 'ai_image' : 'fixed', bounds: sh.bounds, prompt: '配一张与主题相关的图片' };
      if (a.role === 'ai_text') return { id: sh.id, type: 'text', role: 'ai_text', bounds: sh.bounds, textStyle: sh.textStyle, prompt: '根据主题生成简洁标题或正文' };
      return { id: sh.id, type: 'text', role: 'fixed', bounds: sh.bounds, textStyle: sh.textStyle, content: sh.text };
    })
  };
  const saved = await saveTemplate({ name: template.name, folder: '', template, preview: '' });
  // 4. 重新读取（模板库读取 + 版本迁移）
  const reread = await getTemplate(saved.id, '');
  assert(reread.template.shapes.length === template.shapes.length, '重读形状数=' + reread.template.shapes.length + '，期望 ' + template.shapes.length);
  const titleShape = reread.template.shapes.find((s) => s.role === 'ai_text');
  assert(!!titleShape, '重读后缺少 ai_text 位');
  // 5. 裁剪几何（模拟前端裁剪：模板图片位比例 → 裁剪矩形比例一致、输出有上限）
  const cropBase = { naturalW: 800, naturalH: 600, frameW: 400, frameH: 250, scale: 1, tx: 0, ty: 0 };
  const crop = cropRectFromState(fitCrop(cropBase));
  assert(Math.abs(crop.sw / crop.sh - 400 / 250) < 1e-6, '裁剪比例与图片位不一致');
  const out = outputSize(crop.sw, crop.sh, 2048);
  assert(out.w <= 2048 && out.h <= 2048, '裁剪输出超上限');
  // 6. 生成（文字 + 图片）
  const texts = {}; const imgIds = {};
  for (const s of reread.template.shapes) {
    if (s.role === 'ai_text') texts[s.id] = s.id === titleShape.id ? '黄金路径标题' : '黄金路径正文';
    if (s.role === 'ai_image') imgIds[s.id] = PNG_BLUE;
  }
  const b64 = await buildSlideBase64({ template: reread.template, images: imgIds, texts, vars: {} });
  const probe = await probePptx(b64);
  assert(probe.slideCount === 1, '生成页数=' + probe.slideCount);
  assert(probe.media.length === 1, 'media 数=' + probe.media.length + '，期望 1');
  assert(probe.pages[0].texts.join('|').includes('黄金路径标题'), '生成文本缺失标题');
  // 7. PowerPoint COM 打开检查
  const com = await comOpen(b64, { slides: 1, pictures: 1, containsTexts: ['黄金路径标题'] });
  return '标记=' + applied.map((a) => a.shape.name + ':' + a.role).join(',') + ' | 裁剪比例 ✓ | media=' + probe.media.length + (com.skipped ? ' | COM:' + com.reason : ' | COM:' + com.detail);
}]);

// ---------- 测试 2：多图片（2 图片位 → 两张图 → 生成 → media=2 → COM pictures=2） ----------
tests.push(['2/6 多图片（2 图片位 → media=2 → COM）', async () => {
  const tpl = {
    schemaVersion: 1, name: '黄金路径-双图', slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 't1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 0.5, width: 8, height: 0.8 }, textStyle: { size: 28, bold: true }, prompt: '标题' },
      { id: 'img1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 1.8, width: 5, height: 3 }, prompt: '左图' },
      { id: 'img2', type: 'picture', role: 'ai_image', bounds: { left: 6.5, top: 1.8, width: 5, height: 3 }, prompt: '右图' }
    ]
  };
  const b64 = await buildSlideBase64({ template: tpl, images: { img1: PNG_RED, img2: PNG_BLUE }, texts: { t1: '双图测试' }, vars: {} });
  const probe = await probePptx(b64);
  assert(probe.media.length === 2, 'media 数=' + probe.media.length + '，期望 2');
  assert(probe.pages[0].picCount === 2, 'slide XML pic 数=' + probe.pages[0].picCount + '，期望 2');
  const com = await comOpen(b64, { slides: 1, pictures: 2 });
  return 'media=' + probe.media.length + ' pic=' + probe.pages[0].picCount + (com.skipped ? ' | COM:' + com.reason : ' | COM:' + com.detail);
}]);

// ---------- 测试 3：表格（保存表格 → CSV 导入 → 扩展行列 → 生成 → COM） ----------
tests.push(['3/6 表格（CSV 导入 → 扩展行列 → 生成 → COM）', async () => {
  const tpl = {
    schemaVersion: 1, name: '黄金路径-表格', slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 't0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 0.4, width: 8, height: 0.8 }, textStyle: { size: 28, bold: true }, prompt: '标题' },
      { id: 'tbl1', type: 'table', role: 'table', bounds: { left: 1, top: 1.5, width: 8, height: 2.4 },
        table: {
          rows: 3, cols: 4, colWidths: [2, 2, 2, 2], rowHeights: [0.5, 0.4, 0.4],
          cells: [
            { row: 0, col: 0, colspan: 4, text: '季度经营数据', textStyle: { bold: true, size: 14 }, fill: '1F4E79' },
            { row: 1, col: 0, text: '营收' }, { row: 1, col: 1, text: '18.5%' }, { row: 1, col: 2, text: '利润' }, { row: 1, col: 3, text: '2.3 亿' },
            { row: 2, col: 0, text: '海外' }, { row: 2, col: 1, text: '32%' }, { row: 2, col: 2, text: '研发' }, { row: 2, col: 3, text: '1.1 亿' }
          ]
        } }
    ]
  };
  // CSV 导入（模拟前端「粘贴 CSV」）→ 比模板 3x4 多一行多一列 → 扩展行列
  const csv = '季度,营收(万),利润(万),海外占比,备注\nQ1,1000,180,32%,良好\nQ2,1200,210,35%,优秀\nQ3,1400,260,38%,优秀\nQ4,1600,310,41%,超预期';
  const tableData = { tbl1: csvToTable(csv) };
  assert(tableData.tbl1.length === 5 && tableData.tbl1[0].length === 5, 'CSV 解析 5x5 失败: ' + tableData.tbl1.length + 'x' + tableData.tbl1[0].length);
  const b64 = await buildSlideBase64({ template: tpl, texts: { t0: '季度经营分析' }, vars: {}, tableData });
  const probe = await probePptx(b64);
  const page = probe.pages[0];
  assert(page.tables === 1, '表格数=' + page.tables);
  assert(page.trCount === 5, '行数=' + page.trCount + '，期望 5（CSV 扩展）');
  assert(page.colCount === 5, '列数=' + page.colCount + '，期望 5（CSV 扩展）');
  assert(page.texts.join('|').includes('超预期'), '扩展单元格文字缺失');
  const com = await comOpen(b64, { slides: 1, tables: 1 });
  return 'CSV 5x5 → 表格 ' + page.trCount + 'x' + page.colCount + '（模板 3x4 已扩展）' + (com.skipped ? ' | COM:' + com.reason : ' | COM:' + com.detail);
}]);

// ---------- 测试 4：套版（3 模板 → 套版 → 生成 3 页 → COM） ----------
tests.push(['4/6 套版（3 模板 → 3 页 → COM）', async () => {
  const PAGE = { width: 13.33, height: 7.5 };
  const names = ['黄金套版-封面', '黄金套版-目录', '黄金套版-结尾'];
  const saved = [];
  for (let i = 0; i < names.length; i++) {
    const t = {
      schemaVersion: 1, name: names[i], slideSize: PAGE,
      shapes: [
        { id: 'tt' + i, type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 10, height: 1.2 }, textStyle: { size: 30, bold: true }, prompt: '标题' },
        { id: 'tb' + i, type: 'text', role: 'ai_text', bounds: { left: 1, top: 3, width: 10, height: 1 }, textStyle: { size: 16 }, prompt: '副标题' }
      ]
    };
    const r = await saveTemplate({ name: names[i], folder: '', template: t, preview: '' });
    saved.push(r.id);
  }
  // 套版存储（引用模板 id，不内嵌快照）
  const deck = { schemaVersion: 1, name: '黄金套版', pages: saved.map((id) => ({ templateId: id })) };
  await saveDeck({ name: deck.name, folder: '', deck, preview: '' });
  assert(listDecks().length === 1, '套版列表数=' + listDecks().length);
  const deckRead = await getDeck(deck.name, '');
  assert(deckRead.deck.pages.length === 3, '套版页数=' + deckRead.deck.pages.length);
  // 生成（从模板库取回模板对象）
  const pages = [];
  for (let i = 0; i < deckRead.deck.pages.length; i++) {
    const p = deckRead.deck.pages[i];
    const t = await getTemplate(p.templateId, '');
    pages.push({ template: t.template, texts: { ['tt' + i]: names[i], ['tb' + i]: '第 ' + (i + 1) + ' 页' }, vars: {} });
  }
  const b64 = await buildDeckBase64(pages);
  const probe = await probePptx(b64);
  assert(probe.slideCount === 3, '套版页数=' + probe.slideCount + '，期望 3');
  for (let i = 0; i < 3; i++) {
    assert(probe.pages[i].texts.join('|').includes(names[i]), '第 ' + (i + 1) + ' 页文字缺失: ' + names[i]);
  }
  const com = await comOpen(b64, { slides: 3 });
  return '3 页套版 ✓ 每页文字 ✓' + (com.skipped ? ' | COM:' + com.reason : ' | COM:' + com.detail);
}]);

// ---------- 测试 5：旧模板兼容（老 schema → 读取迁移 → 生成 → 不崩） ----------
tests.push(['5/6 旧模板兼容（老 schema → 迁移 → 生成不崩）', async () => {
  // 直接写「老结构」template.json：无 schemaVersion / 无版本控制 / 无语义层字段
  const legacyDir = path.join(TPL_ROOT, 'legacy-old');
  fs.mkdirSync(legacyDir, { recursive: true });
  fs.writeFileSync(path.join(legacyDir, 'template.json'), JSON.stringify({
    name: '旧版模板',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 'a', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 8, height: 1 }, textStyle: { size: 24 }, prompt: '标题' },
      { id: 'b', type: 'text', role: 'fixed', bounds: { left: 1, top: 2.5, width: 8, height: 1 }, textStyle: { size: 14 }, content: '固定文字内容' }
    ]
  }));
  // 读取（惰性迁移 v1 + normalize）
  const t = await getTemplate('legacy-old', '');
  assert(t.template.shapes.length === 2, '旧模板读取形状数=' + t.template.shapes.length);
  assert(t.template.shapes.some((s) => s.role === 'ai_text'), '旧模板 ai_text 角色丢失');
  assert(t.template.shapes.some((s) => s.role === 'fixed' && s.content === '固定文字内容'), '旧模板 fixed 内容丢失');
  // 生成 → 不崩 + 文字保留
  const b64 = await buildSlideBase64({ template: t.template, texts: { a: '旧模板新标题' }, vars: {} });
  const probe = await probePptx(b64);
  assert(probe.slideCount === 1, '旧模板生成页数=' + probe.slideCount);
  const tAll = probe.pages[0].texts.join('|');
  assert(tAll.includes('旧模板新标题') && tAll.includes('固定文字内容'), '旧模板文字缺失');
  return '迁移 v1 ✓ 角色保留 ✓ 生成不崩 ✓';
}]);

// ---------- 测试 6：性能冒烟（大负载生成 < 预算，防性能回归） ----------
// 单页重负载（24 文本位 + 6 图片位 + 表格 8x6）< 10s；5 页套版重负载 < 30s。
// 预算宽松（本机实测优化后单页约 0.3s、5 页约 1.5s），只拦「数量级回归」
//（例如后处理链退回 4 次全量 zip 往返、或每页单独建 pptx 实例）。
tests.push(['6/6 性能冒烟（大负载生成 < 预算）', async () => {
  const PAGE = { width: 13.33, height: 7.5 };
  const shapes = [];
  for (let i = 0; i < 24; i++) {
    shapes.push({ id: 't' + i, type: 'text', role: 'ai_text', bounds: { left: 0.5 + (i % 6) * 2.1, top: 0.5 + Math.floor(i / 6) * 1.7, width: 2, height: 1.4 }, textStyle: { size: 12 + (i % 3) * 4, font: 'Microsoft YaHei' }, prompt: '文本' });
  }
  for (let i = 0; i < 6; i++) {
    shapes.push({ id: 'img' + i, type: 'picture', role: 'ai_image', bounds: { left: 0.5 + (i % 3) * 4.2, top: 4.3 + Math.floor(i / 3) * 1.6, width: 4, height: 1.4 }, imageStyle: { shape: 'roundRect' } });
  }
  shapes.push({ id: 'tbl', type: 'table', role: 'table', bounds: { left: 8.3, top: 0.5, width: 4.5, height: 3.6 }, table: { rows: 5, cols: 4, colWidths: [1.2, 1.2, 1.2, 1.2], rowHeights: [0.7, 0.7, 0.7, 0.7, 0.7], tblPr: { firstRow: true }, cells: [] } });
  const template = { schemaVersion: 1, name: 'perf-smoke', slideSize: PAGE, shapes };
  const images = {};
  for (let i = 0; i < 6; i++) images['img' + i] = (i % 2 ? PNG_BLUE : PNG_RED);
  const texts = {};
  for (let i = 0; i < 24; i++) texts['t' + i] = '性能冒烟文本 ' + i + '：用于验证大负载生成耗时在预算内';
  const tableData = { tbl: Array.from({ length: 8 }, (_, r) => Array.from({ length: 6 }, (_, c) => '单元格' + r + '-' + c)) };
  // 单页重负载 < 10s
  const t0 = Date.now();
  const b64 = await buildSlideBase64({ template, images, texts, vars: {}, tableData });
  const ms1 = Date.now() - t0;
  assert(ms1 < 10000, '单页重负载生成 ' + ms1 + 'ms > 10s 预算');
  const probe = await probePptx(b64);
  assert(probe.slideCount === 1, '性能冒烟页数=' + probe.slideCount);
  assert(probe.pages[0].picCount === 6, '性能冒烟 pic 数=' + probe.pages[0].picCount);
  // 5 页套版重负载 < 30s
  const pages = [];
  for (let i = 0; i < 5; i++) pages.push({ template, images, texts, vars: {}, tableData });
  const t1 = Date.now();
  const b64d = await buildDeckBase64(pages);
  const ms5 = Date.now() - t1;
  assert(ms5 < 30000, '5 页套版生成 ' + ms5 + 'ms > 30s 预算');
  const probeD = await probePptx(b64d);
  assert(probeD.slideCount === 5, '5 页套版页数=' + probeD.slideCount);
  return '单页重负载 ' + ms1 + 'ms ✓ | 5 页套版 ' + ms5 + 'ms ✓（预算 10s/30s）';
} ]);
// ---------- 汇总 ----------
(async () => {
  for (const t of tests) await runTest(t[0], t[1]);
  console.log('===== 黄金路径 E2E（' + results.length + ' 条）=====');
  for (const r of results) {
    console.log((r.ok ? '[PASS] ' : '[FAIL] ') + r.name + '  (' + (r.ms / 1000).toFixed(1) + 's)');
    if (r.info) console.log('       ' + r.info);
  }
  const pass = results.filter((r) => r.ok).length;
  const fail = results.length - pass;
  console.log('===== 结果：' + pass + ' 通过 / ' + fail + ' 失败 =====');
  console.log('临时目录：' + TMP);
  if (fail > 0) process.exit(1);
})().catch((e) => { console.error('[E2E] 运行异常: ' + e.message); process.exit(1); });
