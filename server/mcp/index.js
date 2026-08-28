// 模板助手 MCP Server（stdio JSON-RPC 2.0）
// 让 ChatGPT / Claude / Copilot 等 MCP 客户端使用本插件的模板库生成 PPT 页面。
// 调用链：AI → MCP 工具 → 本地后端(127.0.0.1) → 生成页面 → 待写队列 → COM/任务窗格写入 PPT
//
// 启动：node server\mcp\index.js
// ChatGPT 配置（桌面版 → 设置 → MCP 服务器 → 添加）：
//   command: node
//   args: ["<本目录绝对路径>\index.js"]
'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');

// P2-E 安全加固：优先从 runtime.json（%APPDATA%/ppt-ai-addin/runtime.json，后端启动时写入）
// 读取实际端口与一次性 token，与后端「随机端口 + 鉴权」保持一致；PPT_BACKEND 可显式覆盖端口。
let RUNTIME = null;
let DATA_DIR = null; // 模块级：try 块内 const 不会泄漏到块外（心跳依赖它）
try {
  ({ DATA_DIR } = require('../src/config.js'));
  const rtFile = path.join(DATA_DIR, 'runtime.json');
  if (fs.existsSync(rtFile)) RUNTIME = JSON.parse(fs.readFileSync(rtFile, 'utf8'));
} catch { /* 读不到就用默认端口 + 无 token（后端未启用鉴权时兼容） */ }

// 心跳：每 30s touch %APPDATA%/ppt-ai-addin/mcp-seen（内容=本进程 PID），
// 后端 /api/version 据此报告 MCP 在线状态（75s 未 touch 视为离线）。
// MCP 是 stdio 服务，仅在 AI 客户端会话期间存活；进程退出后心跳自然停止。
if (typeof DATA_DIR === 'string') {
  const MCP_SEEN = path.join(DATA_DIR, 'mcp-seen');
  const touchSeen = () => {
    try {
      const now = new Date();
      fs.utimesSync(MCP_SEEN, now, now);
    } catch {
      try { fs.writeFileSync(MCP_SEEN, String(process.pid)); } catch { /* 忽略：心跳失败不影响 MCP 功能 */ }
    }
  };
  try {
    fs.writeFileSync(MCP_SEEN, String(process.pid));
    touchSeen();
    const hb = setInterval(touchSeen, 30000);
    if (hb.unref) hb.unref(); // 不阻止进程退出
  } catch { /* 忽略 */ }
}

const BACKEND = process.env.PPT_BACKEND
  || (RUNTIME && Number.isInteger(RUNTIME.port) ? ('http://127.0.0.1:' + RUNTIME.port) : 'http://127.0.0.1:3788');
const AUTH_TOKEN = (RUNTIME && RUNTIME.token) || '';
const PROTOCOL_VERSION = '2024-11-05';
// 统一版本号（scripts/release.js 生成 server/src/version.js；MCP 随后端/前端同版本发布）
const MCP_VERSION = (() => { try { return require('../src/version.js').VERSION || '0.1.0'; } catch { return '0.1.0'; } })();

// ---------- 表格自动排版引擎（fit_table 工具） ----------
// Node 24 type-stripping 可直接 require addin 的纯 TS 引擎模块（与 server 测试同机制）。
// 引擎为纯函数：输入内容格 + 槽位尺寸 → 输出 colW/rowH/fontSize/overflow，供 generate_slide 的 tables 参数使用。
let fitEngine = null;
function loadFitEngine() {
  if (fitEngine) return fitEngine;
  try { fitEngine = require('../../addin/src/lib/tableModel.ts'); }
  catch (e) { throw new Error('无法加载表格排版引擎（需在源码目录运行 MCP）：' + (e && e.message ? e.message : String(e))); }
  return fitEngine;
}

// ---------- 基础工具 ----------
function jsonLines(logObj) { process.stderr.write(JSON.stringify(logObj) + '\n'); }

async function httpJson(method, urlPath, body) {
  const headers = {};
  if (body) headers['Content-Type'] = 'application/json';
  if (AUTH_TOKEN) headers['X-Auth-Token'] = AUTH_TOKEN;
  const res = await fetch(BACKEND + urlPath, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined
  });
  const text = await res.text();
  let data = null;
  try { data = JSON.parse(text); } catch { data = { raw: text.slice(0, 500) }; }
  if (!res.ok) throw new Error((data && data.error) || ('HTTP ' + res.status + ': ' + text.slice(0, 200)));
  return data;
}

// ---------- 工具实现 ----------
const TOOLS = [
  {
    name: 'list_templates',
    description: '列出本机 PPT 模板库的全部模板（含分类文件夹、名称、更新时间）。生成页面之前先调用本工具选择模板。',
    inputSchema: { type: 'object', properties: {} }
  },
  {
    name: 'get_template',
    description: '读取模板详情：幻灯片尺寸、全部内容位（AI 文本位及其提示词、手动变量位及变量名、AI 图片位）与语义约束（semanticRole 语义角色、maxChars 最多字符数、maxLines 最多行数、minChars 最少字符数、preferredLength 建议长度、required 必填、generationInstruction 生成指令、translate 自动翻译副标题开关与 translateSource 翻译原文来源）。生成文字时请遵守各内容位的语义约束；约束缺失表示不限制。',
    inputSchema: {
      type: 'object',
      properties: {
        templateId: { type: 'string', description: '模板 id（list_templates 返回）' },
        folder: { type: 'string', description: '分类文件夹名（list_templates 返回，未分类可省略）' }
      },
      required: ['templateId']
    }
  },
  {
    name: 'fit_table',
    description: '计算表格内容的自动排版布局（列宽/行高/全表统一字号/是否超出），供 generate_slide 的 tables 参数直接使用。输入：cells（含合并格 rowspan/colspan）+ 槽位宽高（英寸）+ 可选基准字号（默认 14pt）/下限（默认 10pt）；输出：{ rows, cols, colWidths, rowHeights, fontSize, cells, overflow, scaleRatio }。外部 AI 粘贴/生成表格内容后，先调用本工具得到布局，再把整个对象传给 generate_slide 的 tables[表格位 id]。',
    inputSchema: {
      type: 'object',
      properties: {
        cells: {
          type: 'array',
          description: '单元格列表：{ r, c, rowspan?, colspan?, text }（r/c 从 0 起；合并格在主格位置声明）',
          items: {
            type: 'object',
            properties: {
              r: { type: 'number' }, c: { type: 'number' },
              rowspan: { type: 'number', description: '纵向跨行数，默认 1' },
              colspan: { type: 'number', description: '横向跨列数，默认 1' },
              text: { type: 'string' }
            },
            required: ['r', 'c', 'text']
          }
        },
        slotWidthIn: { type: 'number', description: '表格位宽度（英寸，模板 bounds.width，见 get_template 的 tableInfo）' },
        slotHeightIn: { type: 'number', description: '表格位高度（英寸，模板 bounds.height）' },
        baseFontSizePt: { type: 'number', description: '基准字号（磅），默认 14' },
        fontFloorPt: { type: 'number', description: '字号下限（磅），默认 10' }
      },
      required: ['cells', 'slotWidthIn', 'slotHeightIn']
    }
  },
  {
    name: 'generate_slide',
    description: '用指定模板 + AI 提供的文本内容生成一页 PPT（复用模板的精确样式），生成结果进入「待写队列」。texts 的键为 AI 文本位 id（见 get_template，请遵守其语义约束：semanticRole/maxChars/maxLines 等；超限文本会被后端自动截断并返回 warnings），vars 的键为变量位 id。图片位可选：images 的键为 AI 图片位 id，值为图片 dataURL 或 http(s) 图片地址（URL 由后端安全下载后嵌入）。表格位可选：tables 的键为表格位 id（见 get_template 的 tableInfo），值为 fit_table 工具输出的布局对象（含合并格；样式自动套模板表格位、源样式不保留）。',
    inputSchema: {
      type: 'object',
      properties: {
        templateId: { type: 'string', description: '模板 id' },
        folder: { type: 'string' },
        texts: { type: 'object', description: 'AI 文本位内容，键=内容位 id，值=要写入的文字' },
        vars: { type: 'object', description: '手动变量位内容，键=变量位 id，值=文字' },
        images: { type: 'object', description: 'AI 图片位内容（可选），键=图片位 id，值=图片 dataURL 或 http(s) 图片地址' },
        tables: { type: 'object', description: '表格位内容（可选），键=表格位 shape id（见 get_template 的 tableInfo），值=布局对象 { rows, cols, colWidths, rowHeights, fontSize, cells }（来自 fit_table 或人工/外部计算）；样式自动套模板表格位，源样式不保留' }
      },
      required: ['templateId', 'texts']
    }
  },
  {
    name: 'list_decks',
    description: '列出本机全部套版（整份生成模板序列）：名称、分类文件夹、页数、更新时间。生成整份 PPT 之前先调用本工具选择套版。',
    inputSchema: { type: 'object', properties: {} }
  },
  {
    name: 'get_deck',
    description: '读取套版详情：有序页面序列（每页引用的模板 id/folder/可选固定版本、变量、图片说明等）。生成整份之前先调用本工具了解每页用的模板。',
    inputSchema: {
      type: 'object',
      properties: {
        deckId: { type: 'string', description: '套版 id（list_decks 返回）' },
        folder: { type: 'string', description: '分类文件夹名（list_decks 返回，未分类可省略）' }
      },
      required: ['deckId']
    }
  },
  {
    name: 'generate_deck',
    description: '按套版序列整份生成多页 PPT（复用每页模板的精确样式），生成结果进入「待写队列」，可一次写入当前 PPT。pages 的每一项对应一页：templateId + 可选 templateFolder/templateVersion（固定版本保证可复现）+ texts（AI 文本位内容）+ vars（变量位）+ images（图片位 dataURL/URL）。',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: '套版说明（可选，用于待写队列显示）' },
        pages: {
          type: 'array',
          description: '页面序列，每页一个对象',
          items: {
            type: 'object',
            properties: {
              templateId: { type: 'string', description: '模板 id（必填）' },
              templateFolder: { type: 'string', description: '模板分类文件夹' },
              templateVersion: { type: 'string', description: '固定模板版本（可选，如 v3）' },
              texts: { type: 'object', description: 'AI 文本位内容，键=内容位 id' },
              vars: { type: 'object', description: '手动变量位内容，键=变量位 id' },
              images: { type: 'object', description: '图片位内容，键=图片位 id，值=图片 dataURL 或 http(s) 地址' },
              tables: { type: 'object', description: '表格位内容（可选），键=表格位 shape id，值=布局对象 { rows, cols, colWidths, rowHeights, fontSize, cells }（来自 fit_table 或人工/外部计算）；样式自动套模板表格位，源样式不保留' }
            },
            required: ['templateId']
          }
        }
      },
      required: ['pages']
    }
  },
  {
    name: 'list_pending',
    description: '列出待写队列（AI 已生成、尚未写入 PPT 的页面）。',
    inputSchema: { type: 'object', properties: {} }
  },
  {
    name: 'write_slide',
    description: '把待写队列中的页面直接插入当前打开的 PowerPoint 演示文稿（自动插入到当前选中页之后）。需要 PowerPoint 正在运行并打开了目标文档。若失败（如 COM 权限问题），提示用户打开本插件的任务窗格点击写入。',
    inputSchema: {
      type: 'object',
      properties: { pendingId: { type: 'string', description: '待写队列 id（generate_slide 或 list_pending 返回）' } },
      required: ['pendingId']
    }
  },
  {
    name: 'delete_pending',
    description: '丢弃待写队列中的一个页面。',
    inputSchema: {
      type: 'object',
      properties: { pendingId: { type: 'string' } },
      required: ['pendingId']
    }
  },
  {
    name: 'save_template',
    description: '把当前 PowerPoint 页面的全部形状与文字保存为新模板（元素默认标记为固定元素，AI 文本位/图片位等角色可在插件「保存模板」页再调整）。需要 PowerPoint 正在运行并打开了目标文档。',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: '模板名称（必填）' },
        folder: { type: 'string', description: '分类文件夹名（可选，默认保存到根目录；list_templates 可查看现有文件夹）' },
        changeNote: { type: 'string', description: '版本说明（可选）' }
      },
      required: ['name']
    }
  },
  {
    name: 'get_presentation_context',
    description: '获取当前打开的 PowerPoint 演示文稿的结构化上下文：文件名、总页数、当前选中页（index + 稳定 slideId）、页面尺寸（英寸）、保存状态、全部页面的索引列表（含每页 slideId 与元素数，裁剪上限 500 页）、当前选中元素摘要。用于了解文档全貌后决定操作哪一页。注意：slideId 是稳定 ID（移动/插入后不变），index 是当前页码（会变）。若 PowerPoint 未运行或未打开文档，返回错误提示。',
    inputSchema: { type: 'object', properties: {} }
  },
  {
    name: 'get_current_slide',
    description: '获取 PowerPoint 当前页（当前查看/选中页）的完整结构化内容：页面 index/slideId/版式名，以及每个元素的 id/name/type/位置(英寸)/尺寸/旋转/文本(≤300字)/字体/图片/表格/组合等。用于详细分析当前页内容。',
    inputSchema: { type: 'object', properties: {} }
  },
  {
    name: 'get_slide',
    description: '按页码（index，1 起）或稳定 slideId 获取指定页的完整结构化内容（元素列表，同 get_current_slide 的格式）。两个参数至少填一个；slideId 优先（移动后仍稳定）。',
    inputSchema: {
      type: 'object',
      properties: {
        index: { type: 'number', description: '页码，从 1 开始' },
        slideId: { type: 'number', description: '稳定幻灯片 ID（get_presentation_context 返回的 slideId）' }
      }
    }
  },
  {
    name: 'inspect_slide',
    description: '获取指定页（缺省当前页）的紧凑结构摘要：每个元素一行（id/名称/类型/文本前60字/位置/尺寸），适合快速了解页面布局；需要完整文本与样式时用 get_slide 或 get_current_slide。',
    inputSchema: {
      type: 'object',
      properties: {
        index: { type: 'number', description: '页码，从 1 开始' },
        slideId: { type: 'number', description: '稳定幻灯片 ID' }
      }
    }
  }
];

// ---------- 预设指令模板（prompts）----------
// 给 AI 客户端的「一键式」工作流提示：把「先看模板 → 生成 → 写入」的标准流程固化成可复用模板，
// 客户端（如 ChatGPT/Claude）的 /prompts 里可直接选用，降低误用工具的概率。
const PROMPTS = [
  {
    name: 'ppt-gen-page',
    description: '用模板库中的一个模板生成一页 PPT（模板 + 文本 + 可选图片，写入当前演示文稿）',
    template: '请按以下步骤生成一页 PPT：\n1. 调用 list_templates 查看模板库，选择与需求最匹配的模板；\n2. 调用 get_template 读取该模板的幻灯片尺寸与内容位（AI 文本位/变量位/图片位及语义约束）；\n3. 调用 generate_slide 生成页面：texts 填 AI 文本位内容（遵守 maxChars/maxLines 等约束），vars 填变量位，images 可给图片位传图片 dataURL 或图片地址；若模板含表格位，可用 tables 参数提供表格内容（键=表格位 id，值=含 rows/cols/colWidths/rowHeights/fontSize/cells 的布局对象，样式自动适配模板）；\n4. 向用户确认内容后调用 write_slide 把该页写入当前 PowerPoint（自动插入当前页之后）。'
  },
  {
    name: 'ppt-gen-deck',
    description: '按套版整份生成多页 PPT（一次生成并写入当前演示文稿）',
    template: '请按以下步骤整份生成 PPT：\n1. 调用 list_decks 查看可用套版，选择与需求最匹配的套版；\n2. 调用 get_deck 读取套版的页面序列（每页引用的模板）；\n3. 调用 generate_deck 整份生成：pages 里每页填 texts/vars/images，若该页模板含表格位可加 tables（键=表格位 id，值=含 rows/cols/colWidths/rowHeights/fontSize/cells 的布局对象，样式自动适配模板）；可指定 templateVersion 固定版本；\n4. 向用户确认后调用 write_slide 一次性写入当前 PowerPoint。'
  }
];

// pkg 打包后 __dirname 是 exe 内虚拟路径（C:\\snapshot\\...），powershell -File 无法执行虚拟路径；
// 与 pptContext.js 相同策略：磁盘上不存在时，先从打包资源读内容解压到临时目录再执行。
// 注意用完整路径字面量（path.join(__dirname, '..', 'mcp', 'insert.ps1')），pkg 静态分析才能识别并打进包。
const INSERT_PS = path.join(__dirname, '..', 'mcp', 'insert.ps1');
function resolveInsertScript() {
  // 同 pptContext.js：pkg 下 existsSync 对虚拟资产返回 true，必须用 process.pkg 判定打包环境
  const isPkg = typeof process.pkg !== 'undefined';
  if (!isPkg && fs.existsSync(INSERT_PS)) return INSERT_PS; // 开发环境（node 直接跑）直接可用
  try {
    const content = fs.readFileSync(INSERT_PS, 'utf8'); // pkg assets 虚拟路径可读
    const body = content.charCodeAt(0) === 0xFEFF ? content.slice(1) : content; // 去重：资产自带 BOM 时不重复加
    const tmp = path.join(os.tmpdir(), 'ppt-ins-' + process.pid + '-' + Date.now() + '.ps1');
    fs.writeFileSync(tmp, '\uFEFF' + body, 'utf8');
    return tmp;
  } catch {
    return INSERT_PS; // 兜底：交给 spawn 报错
  }
}
// 用 PowerShell COM 把生成的 pptx 插入当前演示文稿
function insertViaCom(pptxPath) {
  return new Promise((resolve) => {
    const ps = resolveInsertScript();
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps, '-PptxPath', pptxPath], {
      windowsHide: true
    });
    let out = '', err = '';
    child.stdout.on('data', (d) => { out += d; });
    child.stderr.on('data', (d) => { err += d; });
    child.on('error', (e) => resolve({ ok: false, message: '无法启动 PowerShell（COM 插入失败）：' + e.message }));
    child.on('close', (code) => {
      resolve(code === 0 ? { ok: true, message: (out || '已插入').trim() } : { ok: false, message: (err || out || 'COM 插入失败').trim() });
    });
  });
}

async function callTool(name, args) {
  switch (name) {
    case 'list_templates': {
      const list = await httpJson('GET', '/api/templates');
      return JSON.stringify(list, null, 2);
    }
    case 'get_template': {
      const { templateId, folder = '' } = args || {};
      if (!templateId) throw new Error('templateId is required');
      const { template } = await httpJson('GET', '/api/templates/' + encodeURIComponent(templateId) + '?folder=' + encodeURIComponent(folder));
      const summary = {
        name: template.name,
        slideSize: template.slideSize,
        background: template.background ? template.background.type : null,
        shapes: (template.shapes || []).map((s) => ({
          id: s.id,
          role: s.role,
          type: s.type,
          name: s.name,
          prompt: s.prompt || undefined,
          varName: s.varName || undefined,
          placeholder: s.placeholder || undefined,
          // 模板语义层：AI 文本位可带语义角色与生成约束（旧模板无这些字段时为 undefined，AI 按不限制处理）
          semanticRole: s.semanticRole || undefined,
          contentType: s.contentType || undefined,
          required: s.required || undefined,
          maxChars: s.maxChars || undefined,
          maxLines: s.maxLines || undefined,
          minChars: s.minChars || undefined,
          preferredLength: s.preferredLength || undefined,
          generationInstruction: s.generationInstruction || undefined,
          // 自动翻译副标题（translate=true 时该文本位生成的是 translateSource 所指原文的英文翻译；'theme' = 全局主题）
          translate: s.translate || undefined,
          translateSource: s.translateSource || undefined,
          // 表格位（type='table'）：暴露布局信息（行列数/列宽/行高/位置，取自模板 shape.table；无表格数据时为 undefined），
          // AI 可用 generate_slide/generate_deck 的 tables 参数填入内容（自动适配模板表格位样式与尺寸，源样式不保留）
          tableInfo: s.type === 'table' && s.table ? {
            rows: s.table.rows,
            cols: s.table.cols,
            colWidths: s.table.colWidths,
            rowHeights: s.table.rowHeights,
            bounds: s.bounds
          } : undefined,
          description: s.type === 'table'
            ? '表格位：可用 generate_slide/generate_deck 的 tables 参数填入表格内容（键=本表格位 id，值=含 rows/cols/colWidths/rowHeights/fontSize/cells 的布局对象，来自 fit_table 或人工/外部计算；自动适配模板样式与尺寸，源样式不保留）。'
            : undefined
        }))
      };
      return JSON.stringify(summary, null, 2);
    }
    case 'fit_table': {
      const { cells, slotWidthIn, slotHeightIn, baseFontSizePt, fontFloorPt } = args || {};
      if (!Array.isArray(cells) || !cells.length) throw new Error('cells array is required');
      if (!(Number.isFinite(slotWidthIn) && slotWidthIn > 0) || !(Number.isFinite(slotHeightIn) && slotHeightIn > 0)) {
        throw new Error('slotWidthIn / slotHeightIn 必须是正数（英寸）');
      }
      const eng = loadFitEngine();
      const fit = eng.fitTableLayout(cells, {
        slotWidthIn,
        slotHeightIn,
        ...(Number.isFinite(baseFontSizePt) && baseFontSizePt > 0 ? { baseFontSizePt } : {}),
        ...(Number.isFinite(fontFloorPt) && fontFloorPt > 0 ? { fontFloorPt } : {})
      });
      return JSON.stringify(fit, null, 2);
    }
    case 'generate_slide': {
      const { templateId, folder = '', texts = {}, vars = {}, images, tables } = args || {};
      if (!templateId) throw new Error('templateId is required');
      const r = await httpJson('POST', '/api/ai/generate', { templateId, folder, texts, vars, images, tables });
      let msg = '已生成页面（模板：' + r.templateName + '，队列 id：' + r.pendingId + '）。如需写入当前 PPT，调用 write_slide 工具。';
      if (r.warnings && r.warnings.length) msg += '\n约束提示：' + r.warnings.join('；');
      return msg;
    }
    case 'list_decks': {
      const list = await httpJson('GET', '/api/decks');
      return JSON.stringify(list, null, 2);
    }
    case 'get_deck': {
      const { deckId, folder = '' } = args || {};
      if (!deckId) throw new Error('deckId is required');
      return JSON.stringify(await httpJson('GET', '/api/decks/' + encodeURIComponent(deckId) + '?folder=' + encodeURIComponent(folder)), null, 2);
    }
    case 'generate_deck': {
      const { pages, name } = args || {};
      if (!Array.isArray(pages) || !pages.length) throw new Error('pages array is required');
      const r = await httpJson('POST', '/api/ai/generate-deck', { pages, name });
      let msg = '已生成整份套版（' + r.pageCount + ' 页，队列 id：' + r.pendingId + '）。如需一次性写入当前 PPT，调用 write_slide 工具。';
      const failed = (r.pageResults || []).filter((p) => !p.ok);
      if (failed.length) msg += '\n有 ' + failed.length + ' 页生成失败：' + failed.map((p) => '第' + (p.index + 1) + '页 ' + p.error).join('；');
      return msg;
    }
    case 'list_pending': {
      const list = await httpJson('GET', '/api/ai/pending');
      return JSON.stringify(list, null, 2);
    }
    case 'write_slide': {
      const { pendingId } = args || {};
      if (!pendingId) throw new Error('pendingId is required');
      const entry = await httpJson('GET', '/api/ai/pending/' + encodeURIComponent(pendingId));
      if (entry.written) return '该页面已写入过（writtenAt=' + entry.writtenAt + '），如需要可重新生成。';
      const tmp = path.join(os.tmpdir(), 'ppt-ai-mcp-' + pendingId + '.pptx');
      fs.writeFileSync(tmp, Buffer.from(entry.base64, 'base64'));
      const r = await insertViaCom(tmp);
      try { fs.rmSync(tmp, { force: true }); } catch { /* ignore */ }
      if (!r.ok) {
        return 'COM 直接插入失败：' + r.message + '。请打开本插件的任务窗格（模板库页会提示「AI 已生成」），点击写入当前 PPT；或检查 PowerPoint 是否以管理员身份运行（与本进程权限不一致时 COM 不可用）。';
      }
      await httpJson('POST', '/api/ai/pending/' + encodeURIComponent(pendingId) + '/write');
      return '已插入当前 PPT：' + r.message;
    }
    case 'delete_pending': {
      const { pendingId } = args || {};
      if (!pendingId) throw new Error('pendingId is required');
      await httpJson('DELETE', '/api/ai/pending/' + encodeURIComponent(pendingId));
      return '已丢弃';
    }
    case 'save_template': {
      const { name, folder = '', changeNote } = args || {};
      if (!name || !String(name).trim()) throw new Error('name is required');
      const r = await httpJson('POST', '/api/templates/from-slide', { name: String(name).trim(), folder, changeNote });
      return '已保存模板「' + r.name + '」（id=' + r.id + (r.folder ? '，文件夹=' + r.folder : '，根目录') + '）。可用 list_templates 查看。';
    }
    case 'get_presentation_context': {
      return JSON.stringify(await httpJson('GET', '/api/context/presentation'), null, 2);
    }
    case 'get_current_slide': {
      return JSON.stringify(await httpJson('GET', '/api/context/current-slide'), null, 2);
    }
    case 'get_slide': {
      const { index, slideId } = args || {};
      if (index === undefined && slideId === undefined) throw new Error('index 或 slideId 至少填一个');
      const q = slideId !== undefined ? 'id=' + encodeURIComponent(slideId) : 'index=' + encodeURIComponent(index);
      return JSON.stringify(await httpJson('GET', '/api/context/slide?' + q), null, 2);
    }
    case 'inspect_slide': {
      const { index, slideId } = args || {};
      const q = slideId !== undefined ? 'id=' + encodeURIComponent(slideId)
        : index !== undefined ? 'index=' + encodeURIComponent(index) : '';
      return JSON.stringify(await httpJson('GET', '/api/context/inspect' + (q ? '?' + q : '')), null, 2);
    }
    default:
      throw new Error('unknown tool: ' + name);
  }
}

// ---------- stdio JSON-RPC 循环 ----------
let buf = '';
// 消息串行队列：handleMessage 是异步的（HTTP 工具），若不排队，响应会按 async 完成顺序发送（可能乱序，违反 JSON-RPC 按 id 对应）；
// 串行化保证响应与请求同序，客户端按 id 匹配也绝对安全。
let msgQueue = Promise.resolve();
process.stdin.setEncoding('utf8');
process.stdin.on('data', (chunk) => {
  buf += chunk;
  let nl;
  while ((nl = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, nl).trim();
    buf = buf.slice(nl + 1);
    if (!line) continue;
    let msg;
    try { msg = JSON.parse(line); } catch { continue; }
    msgQueue = msgQueue.then(() => handleMessage(msg)).catch((e) => jsonLines({ handlerError: e.message }));
  }
});

function send(obj) { process.stdout.write(JSON.stringify(obj) + '\n'); }

async function handleMessage(msg) {
  if (msg.method === 'initialize') {
    send({
      jsonrpc: '2.0',
      id: msg.id,
      result: {
        // 版本协商：回显客户端声明的协议版本（MCP 2024-11-05 客户端只发一个版本；旧版客户端未发时用默认）
        protocolVersion: (msg.params && msg.params.protocolVersion) || PROTOCOL_VERSION,
        // 能力声明：tools（全部工具）+ prompts（预设指令模板）+ logging（服务器→客户端日志通知）。
        // 只声明已实现的能力，客户端可据此决定是否调用 prompts/list。
        capabilities: { tools: {}, prompts: {}, logging: {} },
        serverInfo: { name: 'ppt-ai-addin-mcp', version: MCP_VERSION }
      }
    });
    return;
  }
  if (msg.method === 'notifications/initialized') return; // 通知无需响应
  if (msg.method === 'ping') { send({ jsonrpc: '2.0', id: msg.id, result: {} }); return; } // JSON-RPC 请求必须响应
  if (msg.method === 'tools/list') {
    send({ jsonrpc: '2.0', id: msg.id, result: { tools: TOOLS } });
    return;
  }
  if (msg.method === 'prompts/list') {
    send({ jsonrpc: '2.0', id: msg.id, result: { prompts: PROMPTS } });
    return;
  }
  if (msg.method === 'prompts/get') {
    const name = msg.params && msg.params.name;
    const p = PROMPTS.find((x) => x.name === name);
    if (!p) return send({ jsonrpc: '2.0', id: msg.id, error: { code: -32602, message: 'unknown prompt: ' + name } });
    send({ jsonrpc: '2.0', id: msg.id, result: { description: p.description, messages: [{ role: 'user', content: { type: 'text', text: p.template } }] } });
    return;
  }
  if (msg.method === 'tools/call') {
    const { name, arguments: args } = msg.params || {};
    try {
      const text = await callTool(name, args || {});
      send({ jsonrpc: '2.0', id: msg.id, result: { content: [{ type: 'text', text }] } });
    } catch (e) {
      send({ jsonrpc: '2.0', id: msg.id, result: { content: [{ type: 'text', text: '错误：' + e.message }], isError: true } });
    }
    return;
  }
  // 其他方法：返回 MethodNotFound
  send({ jsonrpc: '2.0', id: msg.id, error: { code: -32601, message: 'method not found: ' + msg.method } });
}

// 启动提示（stderr，不污染 stdio 协议）
jsonLines({ boot: 'ppt-ai-addin MCP server ready', version: MCP_VERSION, backend: BACKEND, auth: AUTH_TOKEN ? 'token' : 'none' });
