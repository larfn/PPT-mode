// MCP 服务器端到端测试（stdio JSON-RPC，真实子进程）：
// 验证 get_template 返回语义字段、generate_slide 应用语义约束并回传 warnings。
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');
const { createApp } = require('../src/index.js');
const { setTemplateRoot, saveTemplate, deleteTemplate } = require('../src/templateStore.js');
const { setPendingRoot, pendingDir } = require('../src/pendingStore.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-mcp-'));
process.env.PPT_TEMPLATE_ROOT = tmp;
setTemplateRoot(tmp);
const pendTmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-mcp-pend-'));
setPendingRoot(pendTmp);

const MCP_PATH = path.join(__dirname, '..', 'mcp', 'index.js');

function startMcp(backendUrl) {
  const child = spawn(process.execPath, [MCP_PATH], {
    env: { ...process.env, PPT_BACKEND: backendUrl },
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true
  });
  let buf = '';
  const waiters = new Map();
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (d) => {
    buf += d;
    let nl;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim();
      buf = buf.slice(nl + 1);
      if (!line) continue;
      let msg;
      try { msg = JSON.parse(line); } catch { continue; }
      const w = waiters.get(msg.id);
      if (w) { waiters.delete(msg.id); clearTimeout(w.timer); msg.error ? w.reject(new Error(JSON.stringify(msg.error))) : w.resolve(msg.result); }
    }
  });
  let stderrLog = '';
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (d) => { stderrLog += d; });
  return {
    child, stderrLog,
    request(method, params) {
      return new Promise((resolve, reject) => {
        const id = 'r' + Math.random().toString(36).slice(2, 10);
        const timer = setTimeout(() => {
          waiters.delete(id);
          reject(new Error('mcp request timeout: ' + method));
        }, 8000);
        waiters.set(id, { resolve, reject, timer });
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
      });
    },
    close() { try { child.kill(); } catch { /* ignore */ } }
  };
}

test('MCP get_template returns semantic fields; generate_slide applies constraints and reports warnings', async () => {
  const saved = saveTemplate({
    name: 'MCP语义模板',
    template: {
      schemaVersion: 1, name: 'MCP语义模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        {
          id: 'title1', type: 'text', role: 'ai_text',
          bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 32 },
          prompt: '写一个有力的标题',
          semanticRole: 'title', required: true,
          maxChars: 10, maxLines: 1, minChars: 3, preferredLength: 8,
          generationInstruction: '给一个有力的结论'
        },
        { id: 'fix1', type: 'picture', role: 'fixed', content: 'data:image/png;base64,AAAA', bounds: { left: 1, top: 3, width: 4, height: 3 } }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    // 握手
    const init = await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    assert.equal(init.serverInfo.name, 'ppt-ai-addin-mcp');
    // 工具列表
    const tools = await mcp.request('tools/list', {});
    const names = tools.tools.map((t) => t.name);
    assert.ok(names.includes('get_template') && names.includes('generate_slide'));
    // get_template：语义字段齐全
    const gt = await mcp.request('tools/call', { name: 'get_template', arguments: { templateId: saved.id } });
    assert.ok(!gt.isError, 'get_template should succeed');
    const summary = JSON.parse(gt.content[0].text);
    assert.equal(summary.name, 'MCP语义模板');
    const s = summary.shapes.find((x) => x.id === 'title1');
    assert.equal(s.semanticRole, 'title');
    assert.equal(s.contentType, undefined);
    assert.equal(s.required, true);
    assert.equal(s.maxChars, 10);
    assert.equal(s.maxLines, 1);
    assert.equal(s.minChars, 3);
    assert.equal(s.preferredLength, 8);
    assert.equal(s.generationInstruction, '给一个有力的结论');
    const f = summary.shapes.find((x) => x.id === 'fix1');
    assert.equal(f.semanticRole, undefined, 'fixed elements must not be forced to carry semantic fields');
    assert.equal(f.maxChars, undefined);
    // generate_slide：超限文本被截断 + warnings 回传
    const gs = await mcp.request('tools/call', {
      name: 'generate_slide',
      arguments: { templateId: saved.id, texts: { title1: '123456789012345' }, vars: {} }
    });
    const gsText = gs.content[0].text;
    assert.ok(gsText.includes('队列 id'), 'should return pending id');
    assert.ok(gsText.includes('约束提示'), 'warnings should be surfaced, got: ' + gsText);
    const list = await (await fetch(backend + '/api/ai/pending')).json();
    assert.equal(list.length, 1);
    const detail = await (await fetch(backend + '/api/ai/pending/' + list[0].id)).json();
    assert.equal(detail.texts.title1, '1234567890', 'text truncated to maxChars=10');
  } finally {
    mcp.close();
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('MCP generate_slide with old-style template passes texts through without warnings', async () => {
  const saved = saveTemplate({
    name: 'MCP旧模板',
    template: {
      schemaVersion: 1, name: 'MCP旧模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'a1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 } }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    const text = '长'.repeat(100);
    const gs = await mcp.request('tools/call', {
      name: 'generate_slide',
      arguments: { templateId: saved.id, texts: { a1: text }, vars: {} }
    });
    assert.ok(!gs.isError);
    assert.ok(!gs.content[0].text.includes('约束提示'), 'no constraints => no warnings');
    const list = await (await fetch(backend + '/api/ai/pending')).json();
    const detail = await (await fetch(backend + '/api/ai/pending/' + list[0].id)).json();
    assert.equal(detail.texts.a1, text, 'text passes through unchanged');
  } finally {
    mcp.close();
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— 当前演示文稿上下文（MCP 工具端到端，后端注入 mock runner）——

test('MCP context tools: get_presentation_context / get_current_slide / get_slide / inspect_slide', async () => {
  const { setPsRunner } = require('../src/pptContext.js');
  // mock runner：按 -Action 返回不同的结构化数据
  setPsRunner(async (args) => {
    const a = args.join(' ');
    if (a.includes('-Action') && a.includes('presentation')) {
      return {
        ok: true, kind: 'presentation', name: 'MCP上下文测试.pptx', saved: true,
        slideCount: 2, slideSize: { width: 13.33, height: 7.5 },
        currentSlide: { index: 1, slideId: 300, selectionType: 2 },
        slides: [
          { index: 1, slideId: 300, shapeCount: 2 },
          { index: 2, slideId: 301, shapeCount: 1 }
        ]
      };
    }
    if (a.includes('-Action') && a.includes('current-slide')) {
      return { ok: true, kind: 'current-slide', index: 1, slideId: 300, shapeCount: 2, shapes: [
        { id: 11, name: '标题 1', type: 'textBox', text: '当前页标题', left: 0.5, top: 0.5, width: 8, height: 1 },
        { id: 12, name: '表格 1', type: 'table', isTable: true, tableInfo: { rows: 3, cols: 2 }, left: 1, top: 2, width: 6, height: 3 }
      ] };
    }
    if (a.includes('-Action') && a.includes('slide')) {
      return { ok: true, kind: 'slide', index: 2, slideId: 301, shapeCount: 1, shapes: [
        { id: 21, name: '正文 1', type: 'textBox', text: '第二页正文内容', font: { name: '微软雅黑', size: 18 } }
      ] };
    }
    if (a.includes('-Action') && a.includes('inspect')) {
      return { ok: true, kind: 'inspect', index: 1, slideId: 300, shapeCount: 2, shapes: [
        { id: 11, name: '标题 1', type: 'textBox', text: '当前页标题', left: 0.5 },
        { id: 12, name: '表格 1', type: 'table', text: '', left: 1 }
      ] };
    }
    return { ok: false, error: 'unknown action: ' + a };
  });

  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    const tools = await mcp.request('tools/list', {});
    const names = tools.tools.map((t) => t.name);
    for (const n of ['get_presentation_context', 'get_current_slide', 'get_slide', 'inspect_slide']) {
      assert.ok(names.includes(n), 'tool missing: ' + n);
    }

    // get_presentation_context
    const pc = await mcp.request('tools/call', { name: 'get_presentation_context', arguments: {} });
    assert.ok(!pc.isError);
    const pcJson = JSON.parse(pc.content[0].text);
    assert.equal(pcJson.name, 'MCP上下文测试.pptx');
    assert.equal(pcJson.currentSlide.slideId, 300);
    assert.equal(pcJson.slides.length, 2);

    // get_current_slide
    const cs = await mcp.request('tools/call', { name: 'get_current_slide', arguments: {} });
    const csJson = JSON.parse(cs.content[0].text);
    assert.equal(csJson.kind, 'current-slide');
    assert.equal(csJson.shapes[0].text, '当前页标题');
    assert.equal(csJson.shapes[1].type, 'table');

    // get_slide by index
    const s1 = await mcp.request('tools/call', { name: 'get_slide', arguments: { index: 2 } });
    const s1Json = JSON.parse(s1.content[0].text);
    assert.equal(s1Json.slideId, 301);
    assert.equal(s1Json.shapes[0].font.name, '微软雅黑');

    // get_slide by slideId
    const s2 = await mcp.request('tools/call', { name: 'get_slide', arguments: { slideId: 301 } });
    const s2Json = JSON.parse(s2.content[0].text);
    assert.equal(s2Json.index, 2);

    // get_slide without params → friendly error text
    const s3 = await mcp.request('tools/call', { name: 'get_slide', arguments: {} });
    assert.ok(s3.isError === true || s3.content[0].text.includes('index'), 'should reject missing params: ' + s3.content[0].text);

    // inspect_slide default (current)
    const ins = await mcp.request('tools/call', { name: 'inspect_slide', arguments: {} });
    const insJson = JSON.parse(ins.content[0].text);
    assert.equal(insJson.kind, 'inspect');
    assert.equal(insJson.shapes.length, 2);
    assert.equal(insJson.shapes[0].text, '当前页标题');

    // inspect_slide by index
    const ins2 = await mcp.request('tools/call', { name: 'inspect_slide', arguments: { index: 1 } });
    assert.equal(JSON.parse(ins2.content[0].text).kind, 'inspect');
  } finally {
    mcp.close();
    server.close();
  }
});

// —— MCP 图片位生成（P1-D 补全）：generate_slide 透传 images（dataURL 直通）——
test('MCP generate_slide with images (dataURL) embeds image into generated slide', async () => {
  const saved = saveTemplate({
    name: 'MCP图片模板',
    template: {
      schemaVersion: 1, name: 'MCP图片模板',
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
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    // 1x1 红色 PNG dataURL
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const gs = await mcp.request('tools/call', {
      name: 'generate_slide',
      arguments: { templateId: saved.id, texts: { txt1: '带图文字' }, images: { img1: png } }
    });
    assert.ok(!gs.isError, 'generate_slide with images should succeed: ' + JSON.stringify(gs));
    const pendingId = gs.content[0].text.match(/队列 id：([^，。\s]+)/)[1];
    const detail = await (await fetch(backend + '/api/ai/pending/' + pendingId)).json();
    assert.ok(detail.images && detail.images.img1 === png, 'pending should carry the image dataURL');
    // base64 可解出 zip（pptx 内含 png 资源）
    const buf = Buffer.from(detail.base64, 'base64');
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);
    const pngEntries = Object.keys(zip.files).filter((n) => n.startsWith('ppt/media/') && /\.[a-z0-9]+$/i.test(n.split('?')[0]));
    assert.ok(pngEntries.length >= 1, 'generated pptx should embed the png, got: ' + Object.keys(zip.files).filter((n) => n.startsWith('ppt/media/')).join(','));
  } finally {
    mcp.close();
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— MCP 套版（P1-D 补全）：list_decks / get_deck / generate_deck ——
test('MCP deck tools: list_decks / get_deck / generate_deck → pending queue', async () => {
  const t1 = saveTemplate({
    name: 'MCP套版页1',
    template: {
      schemaVersion: 1, name: 'MCP套版页1',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'a1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } }]
    },
    preview: ''
  });
  const t2 = saveTemplate({
    name: 'MCP套版页2',
    template: {
      schemaVersion: 1, name: 'MCP套版页2',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [{ id: 'b1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } }]
    },
    preview: ''
  });
  const { saveDeck, deleteDeck } = require('../src/deckStore.js');
  const deckSaved = saveDeck({
    name: 'MCP测试套版',
    folder: '',
    deck: {
      name: 'MCP测试套版',
      pages: [
        { templateId: t1.id, texts: { a1: '第一页' } },
        { templateId: t2.id, texts: { b1: '第二页' } }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    const tools = await mcp.request('tools/list', {});
    const names = tools.tools.map((t) => t.name);
    for (const n of ['list_decks', 'get_deck', 'generate_deck']) assert.ok(names.includes(n), 'tool missing: ' + n);
    // list_decks
    const ld = await mcp.request('tools/call', { name: 'list_decks', arguments: {} });
    const list = JSON.parse(ld.content[0].text);
    assert.ok(list.some((d) => d.id === deckSaved.id && d.pageCount === 2), 'list_decks should include saved deck');
    // get_deck
    const gd = await mcp.request('tools/call', { name: 'get_deck', arguments: { deckId: deckSaved.id } });
    const deck = JSON.parse(gd.content[0].text);
    assert.equal(deck.deck.pages.length, 2);
    assert.equal(deck.deck.pages[0].templateId, t1.id);
    // generate_deck
    const gen = await mcp.request('tools/call', {
      name: 'generate_deck',
      arguments: {
        name: 'MCP套版生成测试',
        pages: [
          { templateId: t1.id, texts: { a1: '第一页内容' } },
          { templateId: t2.id, texts: { b1: '第二页内容' } }
        ]
      }
    });
    assert.ok(!gen.isError, 'generate_deck should succeed: ' + JSON.stringify(gen));
    assert.ok(gen.content[0].text.includes('2 页'), 'should report 2 pages: ' + gen.content[0].text);
    const id = gen.content[0].text.match(/队列 id：([^，。\s]+)/)[1];
    const detail = await (await fetch(backend + '/api/ai/pending/' + id)).json();
    assert.equal(detail.deck, true);
    assert.equal(detail.pageCount, 2);
    // 多页 base64 可解出 2 个 slide xml
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(Buffer.from(detail.base64, 'base64'));
    const slides = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
    assert.ok(slides.length >= 2, 'deck base64 should contain 2+ slides, got: ' + slides.length);
  } finally {
    mcp.close();
    server.close();
    deleteDeck(deckSaved.id, '');
    deleteTemplate(t1.id);
    deleteTemplate(t2.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— MCP 预设指令模板（prompts/list + prompts/get）——
test('MCP prompts: prompts/list returns templates; prompts/get returns content', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const mcp = startMcp('http://127.0.0.1:' + server.address().port);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    const pl = await mcp.request('prompts/list', {});
    const names = pl.prompts.map((p) => p.name);
    assert.ok(names.includes('ppt-gen-page') && names.includes('ppt-gen-deck'));
    const pg = await mcp.request('prompts/get', { name: 'ppt-gen-page' });
    assert.ok(pg.messages.length >= 1 && pg.messages[0].content.text.includes('generate_slide'));
    // 未知 prompt → JSON-RPC 错误（客户端 request 会 reject）
    let rejected = false;
    try { await mcp.request('prompts/get', { name: 'nope' }); } catch { rejected = true; }
    assert.ok(rejected, 'unknown prompt should error');
  } finally {
    mcp.close();
    server.close();
  }
});
// —— MCP 表格位（表格自动排版引擎）：get_template 暴露 tableInfo，generate_slide 透传 tables ——
test('MCP table slots: get_template exposes tableInfo; generate_slide passes tables to pending', async () => {
  const saved = saveTemplate({
    name: 'MCP表格模板',
    template: {
      schemaVersion: 1, name: 'MCP表格模板',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        {
          id: 'tbl1', type: 'table', role: 'table',
          bounds: { left: 1, top: 1, width: 4, height: 1.25 },
          table: {
            rows: 2, cols: 3,
            colWidths: [1.5, 1.5, 1],
            rowHeights: [0.5, 0.375],
            tblPr: { firstRow: true },
            cells: [
              { r: 0, c: 0, rowspan: 1, colspan: 1, text: '列1' },
              { r: 0, c: 1, rowspan: 1, colspan: 1, text: '列2' },
              { r: 0, c: 2, rowspan: 1, colspan: 1, text: '列3' },
              { r: 1, c: 0, rowspan: 1, colspan: 1, text: 'a' },
              { r: 1, c: 1, rowspan: 1, colspan: 1, text: 'b' },
              { r: 1, c: 2, rowspan: 1, colspan: 1, text: 'c' }
            ]
          }
        }
      ]
    },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    // tools/list：generate_slide / generate_deck 的 schema 暴露 tables
    const tools = await mcp.request('tools/list', {});
    const gsTool = tools.tools.find((t) => t.name === 'generate_slide');
    assert.ok(gsTool && gsTool.inputSchema.properties.tables, 'generate_slide schema should expose tables');
    const gdTool = tools.tools.find((t) => t.name === 'generate_deck');
    assert.ok(gdTool && gdTool.inputSchema.properties.pages.items.properties.tables, 'generate_deck page schema should expose tables');
    // get_template：表格位带 tableInfo（rows/cols/colWidths/rowHeights/bounds）+ description 提示
    const gt = await mcp.request('tools/call', { name: 'get_template', arguments: { templateId: saved.id } });
    assert.ok(!gt.isError, 'get_template should succeed');
    const summary = JSON.parse(gt.content[0].text);
    const tbl = summary.shapes.find((x) => x.id === 'tbl1');
    assert.equal(tbl.type, 'table');
    assert.ok(tbl.tableInfo, 'table shape should expose tableInfo');
    assert.equal(tbl.tableInfo.rows, 2);
    assert.equal(tbl.tableInfo.cols, 3);
    assert.deepEqual(tbl.tableInfo.colWidths, [1.5, 1.5, 1]);
    assert.deepEqual(tbl.tableInfo.rowHeights, [0.5, 0.375]);
    assert.deepEqual(tbl.tableInfo.bounds, { left: 1, top: 1, width: 4, height: 1.25 });
    assert.ok(tbl.description && tbl.description.includes('tables'), 'table shape description should hint the tables param');
    // generate_slide 带 tables（含 colspan=2 合并格）→ ok + 队列条目带 tables
    const tables = {
      tbl1: {
        rows: 2, cols: 3,
        colWidths: [1.5, 1.5, 1],
        rowHeights: [0.5, 0.375],
        fontSize: 14,
        cells: [
          { r: 0, c: 0, rowspan: 1, colspan: 2, text: '合并标题' },
          { r: 0, c: 2, rowspan: 1, colspan: 1, text: '数量' },
          { r: 1, c: 0, rowspan: 1, colspan: 1, text: '甲' },
          { r: 1, c: 1, rowspan: 1, colspan: 1, text: '乙' },
          { r: 1, c: 2, rowspan: 1, colspan: 1, text: '3' }
        ]
      }
    };
    const gs = await mcp.request('tools/call', {
      name: 'generate_slide',
      arguments: { templateId: saved.id, texts: {}, vars: {}, tables }
    });
    assert.ok(!gs.isError, 'generate_slide with tables should succeed: ' + JSON.stringify(gs));
    const pendingId = gs.content[0].text.match(/队列 id：([^，。\s]+)/)[1];
    assert.ok(pendingId, 'should return pending id');
    const detail = await (await fetch(backend + '/api/ai/pending/' + pendingId)).json();
    assert.ok(detail.tables && detail.tables.tbl1, 'pending entry should carry tables');
    assert.equal(detail.tables.tbl1.fontSize, 14);
    assert.equal(detail.tables.tbl1.cells.length, 5);
    assert.equal(detail.tables.tbl1.cells[0].colspan, 2, 'merged cell colspan preserved');
    assert.equal(detail.tables.tbl1.cells[0].text, '合并标题');
  } finally {
    mcp.close();
    server.close();
    deleteTemplate(saved.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// —— MCP 套版表格位：generate_deck 页面项带 tables 参数可正常生成 ——
test('MCP generate_deck accepts tables per page', async () => {
  const t = saveTemplate({
    name: 'MCP套版表格页',
    template: {
      schemaVersion: 1, name: 'MCP套版表格页',
      slideSize: { width: 13.33, height: 7.5 },
      shapes: [
        { id: 'titleA', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } },
        {
          id: 'tblA', type: 'table', role: 'table',
          bounds: { left: 1, top: 2.5, width: 4, height: 1.25 },
          table: {
            rows: 2, cols: 2,
            colWidths: [2, 2],
            rowHeights: [0.5, 0.375],
            cells: [
              { r: 0, c: 0, rowspan: 1, colspan: 1, text: 'x' },
              { r: 0, c: 1, rowspan: 1, colspan: 1, text: 'y' },
              { r: 1, c: 0, rowspan: 1, colspan: 1, text: '1' },
              { r: 1, c: 1, rowspan: 1, colspan: 1, text: '2' }
            ]
          }
        }
      ]
    },
    preview: ''
  });
  const { saveDeck, deleteDeck } = require('../src/deckStore.js');
  const deckSaved = saveDeck({
    name: 'MCP表格套版',
    folder: '',
    deck: { name: 'MCP表格套版', pages: [{ templateId: t.id, texts: { titleA: '表格页' } }] },
    preview: ''
  });
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    const gen = await mcp.request('tools/call', {
      name: 'generate_deck',
      arguments: {
        name: 'MCP表格套版生成',
        pages: [{
          templateId: t.id,
          texts: { titleA: '表格页内容' },
          tables: {
            tblA: {
              rows: 2, cols: 2,
              colWidths: [2, 2],
              rowHeights: [0.5, 0.375],
              fontSize: 14,
              cells: [
                { r: 0, c: 0, rowspan: 1, colspan: 2, text: '合并格' },
                { r: 1, c: 0, rowspan: 1, colspan: 1, text: '甲' },
                { r: 1, c: 1, rowspan: 1, colspan: 1, text: '乙' }
              ]
            }
          }
        }]
      }
    });
    assert.ok(!gen.isError, 'generate_deck with tables should succeed: ' + JSON.stringify(gen));
    assert.ok(gen.content[0].text.includes('1 页'), 'should report 1 page: ' + gen.content[0].text);
    const id = gen.content[0].text.match(/队列 id：([^，。\s]+)/)[1];
    const detail = await (await fetch(backend + '/api/ai/pending/' + id)).json();
    assert.equal(detail.deck, true);
    assert.equal(detail.pageCount, 1);
  } finally {
    mcp.close();
    server.close();
    deleteDeck(deckSaved.id, '');
    deleteTemplate(t.id);
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

test('MCP fit_table: 计算合并格表格布局（colW/rowH/fontSize/overflow），可被 generate_slide 直接消费', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const backend = 'http://127.0.0.1:' + server.address().port;
  const mcp = startMcp(backend);
  try {
    await mcp.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    const tools = await mcp.request('tools/list', {});
    assert.ok(tools.tools.some((t) => t.name === 'fit_table'), 'tools/list 包含 fit_table');
    const fit = await mcp.request('tools/call', {
      name: 'fit_table',
      arguments: {
        cells: [
          { r: 0, c: 0, colspan: 2, text: '营业收入' }, { r: 0, c: 2, text: '单位' },
          { r: 1, c: 0, rowspan: 2, text: '华东' }, { r: 1, c: 1, text: '18.5%' }, { r: 1, c: 2, text: '万元' },
          { r: 2, c: 1, text: '22.3%' }, { r: 2, c: 2, text: '万元' }
        ],
        slotWidthIn: 4,
        slotHeightIn: 1.25,
        baseFontSizePt: 14,
        fontFloorPt: 10
      }
    });
    assert.ok(!fit.isError, 'fit_table should succeed: ' + JSON.stringify(fit));
    const r = JSON.parse(fit.content[0].text);
    assert.equal(r.cols, 3);
    assert.equal(r.rows, 3);
    assert.ok(Math.abs(r.colWidths.reduce((a, b) => a + b, 0) - 4) < 1e-6, 'colW 总和 = 槽宽');
    assert.ok(r.fontSize >= 10 && r.fontSize <= 14, '字号在 [10,14] 内');
    assert.ok(Array.isArray(r.rowHeights) && r.rowHeights.length === 3);
    assert.equal(typeof r.overflow, 'boolean');
  } finally {
    mcp.close();
    server.close();
    try { fs.rmSync(pendingDir(), { recursive: true, force: true }); } catch { /* ignore */ }
  }
});
