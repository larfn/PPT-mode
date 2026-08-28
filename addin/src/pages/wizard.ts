import { Api, TemplateDoc, TemplateShape, TemplateMeta, getDefaultImageProvider } from '../api.js';
import { renderPreview } from '../previewPanel.js';
import { writePendingSlide, testFileOpenable } from '../office/writeSlide.js';
import type { PendingSlide } from '../office/writeSlide.js';
import { showProgress, showModal, showPreviewModal, showToast } from '../ui.js';
import { escapeHtml, escapeAttr } from '../lib/html.js';
import { constraintSummary, shapeConstraints, semanticRoleLabel, sortShapesByPosition } from '../lib/semantic.js';
import { openImageCropEditor } from '../lib/cropEditor.js';
import { runQualityChecks, clearImageSizeCache, QUALITY_CATEGORIES } from '../lib/qualityCheck.js';
import type { QualityReport, QualityLevel } from '../lib/qualityCheck.js';
import { markInputError, clearInputError } from '../lib/formHint.js';
import { infoTip } from '../lib/tooltip.js';
import { loadOutputMode, resolveSlotMode, LIMIT_CHOICES } from '../lib/outputMode.js';
import type { OutputMode } from '../lib/outputMode.js';
import { fitTableLayout, cellsFromGrid, mergeSameTextCells } from '../lib/tableModel.js';
import type { FitCell, FitResult } from '../lib/tableModel.js';
import { parseTableHtml, parseTableCsv } from '../lib/tableClipboard.js';
import { localTextToLatex, todayStr } from '../lib/formula.js';
import { defaultImgState, normalizeImagePageSize, thumbnailUrlOf } from './wizard/imageState.js';
import type { ImgSlotState } from './wizard/imageState.js';
import {
  cloneTemplateWithWizardBackground,
  defaultWizardBackgroundState
} from './wizard/backgroundState.js';
import type { WizardBackgroundState } from './wizard/backgroundState.js';

// ================= 页面生成工作台 =================
// 最终形态不是"向导"：所有编辑区（主题/文字/图片/表格）平铺在一页，
// 下方常驻「实时预览」（防抖自动刷新，点击可放大），底部一键「生成并插入 PPT」。

let template: TemplateDoc | null = null;
// 每个 AI 图片位各自的图片 dataURL：{ [shapeId]: dataURL }（支持一个模板多个图片位）
let images: Record<string, string> = {};
let texts: Record<string, string> = {};   // 生成的文本（每个 AI 文本位）
let prompts: Record<string, string> = {}; // 各文本位的提示词输入（进入工作台后保留）
let formulaLatex: Record<string, string> = {}; // 公式位 AI 生成的 LaTeX（生成 PPT 时优先使用，否则本地规范化）
let vars: Record<string, string> = {};    // 旧模板手动变量位（写回 vars，兼容后端）
let tableData: Record<string, string[][]> = {}; // 表格位数据：{ [shapeId]: string[][] }
let tableMerges: Record<string, FitCell[]> = {}; // 表格位合并结构（主格列表，来自 HTML 粘贴/一键合并）
let tableFits: Record<string, FitResult> = {};   // 表格位最近一次 fit 结果（预览/信息行/提交复用）
let tableMergeDir: Record<string, 'vertical' | 'horizontal' | 'auto'> = {}; // 一键合并方向（重绘保留）
let lastTableMerges: Record<string, FitCell[]> = {}; // 一键合并前快照（撤销用）
let aiResult: Record<string, string[][]> = {}; // AI 生成结果预览（未应用前保留）
let pasteBuffers: Record<string, string> = {}; // 粘贴面板输入（重绘时保留）
let imgProvider = 'baidu_page'; // 搜图供应商（从配置读取）
const croppedImages = new Set<string>(); // 已人工裁剪的 ai_image 位（生成前清除旧 srcRect）
let globalTheme = ''; // 全局主题（作为各段 AI 生成的统一背景）
let backgroundState: WizardBackgroundState = defaultWizardBackgroundState(null);
const expandedTables = new Set<string>(); // 已展开的表格位（重绘后保持展开状态）
let previewTimer: number | undefined; // 实时预览防抖计时器
let previewCollapsed = true; // 预览默认收起（需要时点「展开 ▴」查看实时预览）
let lastQuality: QualityReport | null = null; // 最近一次质量检查结果（写入前门禁）
let qualitySeq = 0; // 质量检查刷新序号（防异步竞态，只采用最后一次结果）
let defaultOutputMode: OutputMode = loadOutputMode(); // 新文本位初始输出模式
let slotOutputModes: Record<string, OutputMode> = {}; // 每个文本位独立的 AI 输出模式

// —— 输入法（IME）全局守卫 ——
// 症状：拼音候选框不显示/闪没（但按空格能出中文）。根因：组词过程中任何布局变化
// （autoResize 撑高、预览重渲染）都会让候选框跟随错位/隐藏。方案：compositionstart→compositionend
// 期间一律跳过渲染/预览刷新（所有输入框统一生效），组词结束后补一次收尾刷新。
let imeComposing = false;   // 组词中（compositionstart 置 true）
let imeDirty = false;       // 组词期间被跳过的刷新（compositionend 后补一次）
let imeInit = false;        // 全局监听只注册一次
let workbenchContainer: HTMLElement | null = null; // 当前工作台容器（compositionend 收尾刷新用）

// 每个图片位独立的搜图状态：{ [shapeId]: { q, page, images, pageSize, selected, providerError } }
let imgStates: Record<string, ImgSlotState> = {};
let imagePageSize = 9;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// 生成向导文字位排序：序号 → 主标题 → 副标题 → 正文 → 图片诠释 → 公式 → 日期（最后）
// （同角色保持模板原始「上到下、左到右」顺序——输入已按位置排序，稳定排序不破坏；不指定/其他 → 日期后；旧手动变量位 → 最末）
const TEXT_SLOT_ROLE_ORDER: Record<string, number> = { seq: 0, title: 1, subtitle: 2, body: 3, caption: 4, formula: 5, date: 6 };
function sortTextSlots(shapes: TemplateShape[]): TemplateShape[] {
  return [...shapes].sort((a, b) => {
    const pa = a.role === 'manual_var' ? 8 : (TEXT_SLOT_ROLE_ORDER[a.semanticRole || ''] ?? 7);
    const pb = b.role === 'manual_var' ? 8 : (TEXT_SLOT_ROLE_ORDER[b.semanticRole || ''] ?? 7);
    return pa - pb;
  });
}

function cloneOutputMode(m: OutputMode): OutputMode {
  return { plain: m.plain, bullets: m.bullets, condense: m.condense, maxChars: m.maxChars, touched: m.touched };
}

function modeForSlot(slotId: string): OutputMode {
  if (!slotOutputModes[slotId]) slotOutputModes[slotId] = cloneOutputMode(defaultOutputMode);
  return slotOutputModes[slotId];
}

// 表格位编辑上限（防粘贴超大表格卡死）
const MAX_TABLE_ROWS = 40;
const MAX_TABLE_COLS = 20;
// 表格排版基准（阶段 1.2：默认 14pt，备选 12pt；下限 10pt；全表统一缩放）
const TABLE_BASE_FONT_PT = 14;
const TABLE_FONT_FLOOR_PT = 10;
const TABLE_FONT_FACE = 'Microsoft YaHei'; // 兜底字体（优先取模板表格位首格样式）

// 从保存的表格结构生成默认 2D 数据（首行为表头，其余为模板原文）
function defaultTableData(s: TemplateShape): string[][] {
  const t = s.table;
  if (!t) return [[]];
  const rows = t.rows || 1, cols = t.cols || 1;
  const out: string[][] = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));
  for (const c of t.cells || []) {
    if (c.row >= 0 && c.row < rows && c.col >= 0 && c.col < cols) out[c.row][c.col] = c.text || '';
  }
  return out;
}

// ================= 表格排版引擎接入（fit 结果供预览/信息行/提交复用） =================
// canvas 测量：不可用时回退引擎默认 em 估算（measure 传 undefined 即默认）
let _tblMeasureCtx: CanvasRenderingContext2D | null | undefined;
function getTblMeasureCtx(): CanvasRenderingContext2D | null {
  if (_tblMeasureCtx === undefined) {
    try { const c = document.createElement('canvas'); _tblMeasureCtx = c.getContext('2d'); }
    catch { _tblMeasureCtx = null; }
  }
  return _tblMeasureCtx;
}
// 引擎约定：measure 返回文本宽度（磅）。canvas 在 96dpi 下 px = pt×96/72，故 px×72/96 即磅。
function canvasMeasure(text: string, fontSizePt: number, fontFace?: string): number {
  const ctx = getTblMeasureCtx();
  if (!ctx) return 0;
  try {
    ctx.font = fontSizePt + 'pt ' + (fontFace || TABLE_FONT_FACE);
    return ctx.measureText(text).width * (72 / 96);
  } catch { return 0; }
}

// 表格位字体：优先模板表格位首格样式（源样式不保留，统一套模板样式）
function tableFontFace(s: TemplateShape): string {
  const first = (s.table && s.table.cells && s.table.cells[0]) || undefined;
  return (first && (first.textStyle && (first.textStyle.eaFont || first.textStyle.font))) || TABLE_FONT_FACE;
}

// 以合并结构为主构建 FitCell（无 merges 时用 cellsFromGrid）；主格文字取 grid[r][c]（越界/undefined → ''）
function buildFitCellsFor(shape: TemplateShape, grid: string[][], merges: FitCell[]): FitCell[] {
  const masters = (merges && merges.length) ? merges : cellsFromGrid(grid || []);
  return masters.map((m) => ({
    r: m.r, c: m.c, rowspan: m.rowspan, colspan: m.colspan,
    text: (grid && grid[m.r] && grid[m.r][m.c] != null) ? String(grid[m.r][m.c]) : ''
  }));
}

// 对某表格位跑一遍排版引擎（基准 14pt / 下限 10pt / 全表统一缩放）
function computeFit(shape: TemplateShape, grid: string[][], merges: FitCell[]): FitResult {
  const b = shape.bounds || { left: 0, top: 0, width: 4, height: 1.5 };
  return fitTableLayout(buildFitCellsFor(shape, grid, merges), {
    slotWidthIn: b.width,
    slotHeightIn: b.height,
    baseFontSizePt: TABLE_BASE_FONT_PT,
    fontFloorPt: TABLE_FONT_FLOOR_PT,
    fontFace: tableFontFace(shape),
    measure: getTblMeasureCtx() ? canvasMeasure : undefined
  });
}

// 取某表格位最近一次 fit；无缓存时计算并缓存（首次渲染即显示适配信息）
function getTableFit(shape: TemplateShape): FitResult {
  if (!tableFits[shape.id]) tableFits[shape.id] = computeFit(shape, tableData[shape.id] || [[]], tableMerges[shape.id]);
  return tableFits[shape.id];
}
// 重算某表格位 fit（数据/合并变化后，保证提交的 tables 与当前数据一致）
function recomputeTableFit(shape: TemplateShape): FitResult {
  tableFits[shape.id] = computeFit(shape, tableData[shape.id] || [[]], tableMerges[shape.id]);
  return tableFits[shape.id];
}

// FitCell[]（含合并）→ 2D 网格：主格文字写入，被覆盖格留空（预览按合并渲染）；超上限裁剪
function gridFromFitCells(cells: FitCell[]): string[][] {
  const rows = Math.min(cells.length ? Math.max(...cells.map((c) => c.r + Math.max(1, c.rowspan))) : 0, MAX_TABLE_ROWS);
  const cols = Math.min(cells.length ? Math.max(...cells.map((c) => c.c + Math.max(1, c.colspan))) : 0, MAX_TABLE_COLS);
  const grid: string[][] = Array.from({ length: rows }, () => Array<string>(cols).fill(''));
  for (const cell of cells) {
    if (cell.r >= rows || cell.c >= cols) continue;
    grid[cell.r][cell.c] = cell.text || '';
  }
  return grid;
}

// HTML 粘贴导入：FitCell 结构 → tableData + tableMerges（合并保留，被覆盖格留空）
function importFitCells(t: string, cells: FitCell[], rawText: string): void {
  const capped = cells.filter((c) => c.r < MAX_TABLE_ROWS && c.c < MAX_TABLE_COLS).map((c) => ({ ...c }));
  tableData[t] = gridFromFitCells(capped);
  tableMerges[t] = capped;
  pasteBuffers[t] = rawText;
  const shape = template?.shapes.find((x) => x.id === t);
  if (shape) recomputeTableFit(shape);
  expandedTables.add(t);
  const hasMerge = capped.some((c) => c.rowspan > 1 || c.colspan > 1);
  const nr = tableData[t].length, nc = (tableData[t][0] || []).length;
  showToast('已导入 ' + nr + ' × ' + nc + (hasMerge ? '（含合并）' : '') + ' ✓', 2200);
}

// 文本导入（CSV/TSV 降级）：无合并结构
function importFromGrid(t: string, grid: string[][], rawText: string): void {
  tableData[t] = grid.slice(0, MAX_TABLE_ROWS).map((row) => row.slice(0, MAX_TABLE_COLS));
  delete tableMerges[t];
  pasteBuffers[t] = rawText;
  const shape = template?.shapes.find((x) => x.id === t);
  if (shape) recomputeTableFit(shape);
  expandedTables.add(t);
  showToast('已导入 ' + tableData[t].length + ' 行 × ' + (tableData[t][0] || []).length + ' 列 ✓', 2000);
}

// fit 信息行（摘要下）：字号/缩字比例 · 行列 · 超出
function tableFitInfoHtml(s: TemplateShape, fit: FitResult): string {
  const sumH = fit.rowHeights.reduce((a, b) => a + b, 0);
  const parts: string[] = [];
  if (fit.fontSize >= fit.baseFontSizePt - 1e-9) {
    parts.push('<span class="fit-ok">已适配：字号 ' + Math.round(fit.fontSize) + 'pt（未缩字）</span>');
  } else {
    parts.push('<span class="fit-shrink">已适配：字号 ' + Math.round(fit.fontSize) + 'pt（' + Math.round(fit.scaleRatio * 100) + '%）</span>');
  }
  parts.push('<span class="fit-ok">' + fit.rows + ' × ' + fit.cols + '</span>');
  if (fit.overflow) {
    const extend = Math.max(0, sumH - (s.bounds.height || 0));
    parts.push('<span class="fit-overflow">⚠ 超出：表格向下延伸 ' + extend.toFixed(1) + '"</span>');
    parts.push(infoTip('仍会写入，表体会自动撑高'));
  }
  return parts.join(' · ');
}

// 重算所有表格位 fit 并刷新各自信息行（数据/合并变化后，防抖回调内调用）
function refreshTableFits(container: HTMLElement): void {
  if (!template) return;
  const tableShapes = template.shapes.filter((s) => s.role === 'table' && s.table);
  for (const s of tableShapes) {
    const fit = recomputeTableFit(s);
    const el = container.querySelector('.wb-tbl-fit[data-tbl="' + s.id + '"]') as HTMLElement | null;
    if (el) el.innerHTML = tableFitInfoHtml(s, fit);
  }
}

// 解析 AI 返回：去掉代码块标记，取第一个 [ 到最后一个 ] 的 JSON 二维数组
function parseTableAiJson(text: string): string[][] {
  let s = (text || '').trim();
  s = s.replace(/^\`\`\`(?:json)?/i, '').replace(/\`\`\`$/, '').trim();
  const start = s.indexOf('[');
  const end = s.lastIndexOf(']');
  if (start < 0 || end <= start) return [];
  try {
    const arr = JSON.parse(s.slice(start, end + 1));
    if (!Array.isArray(arr)) return [];
    return arr.filter((row) => Array.isArray(row)).map((row) => row.map((c) => String(c ?? '')));
  } catch { return []; }
}

// AI 生成结果预览（只读网格 + 应用/放弃按钮）
function aiPreviewHtml(arr: string[][], tblId: string): string {
  const nc = Math.max(1, (arr[0] || []).length);
  return '<div style="margin-top:6px;border:1px solid #ddd;padding:6px">' +
    '<b>生成结果（' + arr.length + ' 行 × ' + nc + ' 列，不含表头）</b>' +
    '<div style="display:grid;grid-template-columns:repeat(' + nc + ',1fr);gap:1px;background:#eee;margin:4px 0">' +
    arr.map((row) => row.map((cell) => '<div style="background:#fff;padding:2px 4px;font-size:12px;overflow:hidden">' + escapeHtml(cell || '') + '</div>').join('')).join('') +
    '</div>' +
    '<div style="display:flex;gap:6px">' +
    '<button class="secondary tbl-ai-apply" data-tbl="' + tblId + '">应用（保留表头）</button>' +
    '<button class="secondary tbl-ai-discard" data-tbl="' + tblId + '">放弃</button>' +
    '</div></div>';
}

// 生成结果框高度随内容自动撑开：不出现框内滚动条，文字完整显示
function autoResize(ta: HTMLTextAreaElement): void {
  ta.style.height = 'auto';
  ta.style.height = `${ta.scrollHeight}px`;
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch { ok = false; }
    ta.remove();
    return ok;
  }
}

// 写入失败后的处理：根据错误类型给出「重试 / 插入到开头 / 复制错误信息」等操作
async function handleWriteError(e: Error): Promise<'retry' | 'start' | 'cancel'> {
  const err = e as Error & { kind?: string; detail?: string; exportBase64?: string };
  const detail = err.detail ? `\n\n详情：${err.detail}` : '';
  const copy = async () => {
    await copyText(`${err.message}${err.detail || ''}`);
    showToast('错误信息已复制', 1500);
  };
  if (err.kind === 'selection') {
    const choice = await showModal({
      title: '未检测到选中的页面',
      message: `${err.message}${detail}\n\n也可以选择「插入到开头」，直接把页面写入演示文稿开头。`,
      buttons: [
        { id: 'retry', label: '重试', kind: 'primary' },
        { id: 'start', label: '插入到开头' },
        { id: 'copy', label: '复制错误信息' },
        { id: 'cancel', label: '取消', kind: 'danger' }
      ]
    });
    if (choice === 'copy') { await copy(); return 'cancel'; }
    return choice as 'retry' | 'start' | 'cancel';
  }
  const buttons: { id: string; label: string; kind?: 'primary' | 'secondary' | 'danger' }[] = [
    { id: 'retry', label: '重试', kind: 'primary' },
    { id: 'copy', label: '复制错误信息' }
  ];
  if (err.exportBase64) {
    buttons.push({ id: 'testfile', label: '测试文件能否打开' });
    buttons.push({ id: 'export', label: '导出诊断文件' });
  }
  buttons.push({ id: 'cancel', label: '取消', kind: 'danger' });
  const choice = await showModal({
    title: '写入 PPT 失败',
    message: `${err.message}${detail}`,
    buttons
  });
  if (choice === 'copy') { await copy(); return 'cancel'; }
  if (choice === 'testfile') {
    try {
      await testFileOpenable(err.exportBase64 || '');
      showModal({
        title: '测试结果：文件可以打开',
        message: '生成的文件能正常在新窗口打开！说明文件本身没有问题，问题出在 PowerPoint 的「插入幻灯片」API（可能是当前 Office 版本的问题）。\n建议：完全退出并重开 PowerPoint 后再试；若仍失败，把最新弹窗的文字复制给我。',
        buttons: [{ id: 'ok', label: '知道了' }]
      });
    } catch (te) {
      showModal({
        title: '测试结果：文件打不开',
        message: `生成的文件无法打开（${(te as Error).message}）——说明文件本身有问题。\n请点击「复制错误信息」把详情发给我。`,
        buttons: [{ id: 'ok', label: '知道了' }]
      });
    }
    return 'cancel';
  }
  if (choice === 'export') {
    try {
      const r = await Api.exportDebugSlide(err.exportBase64 || '');
      showToast('诊断文件已保存：' + r.filePath, 3000);
    } catch (ee) {
      showToast('导出失败：' + (ee as Error).message, 3000);
    }
    return 'cancel';
  }
  return choice === 'retry' ? 'retry' : 'cancel';
}

// —— 自动翻译副标题（AI 文本位 translate/translateSource）——
function findTranslateSource(s: TemplateShape): TemplateShape | undefined {
  if (!template) return undefined;
  const direct = template.shapes.find((x) => x.id === s.translateSource);
  if (direct) return direct;
  if (/^\d+$/.test(String(s.translateSource || ''))) {
    return template.shapes.find((x) => x.id === 'shp' + s.translateSource);
  }
  return undefined;
}
function translateSourceLabel(s: TemplateShape): string {
  if (!s.translate) return '';
  if (s.translateSource === 'theme') return '全局主题';
  const src = findTranslateSource(s);
  if (!src) return '未指定原文';
  if (src.role === 'manual_var') return '变量「' + (src.varName || '变量') + '」';
  const aiIdx = (template?.shapes || []).findIndex((x) => x.id === src.id);
  return '文本位 ' + (aiIdx + 1);
}
function resolveTranslateText(s: TemplateShape, theme: string): string {
  if (s.translateSource === 'theme') return theme;
  const src = findTranslateSource(s);
  if (!src) return '';
  if (src.role === 'manual_var') return vars[src.id] || '';
  if (src.role === 'ai_text') return texts[src.id] || prompts[src.id] || '';
  return '';
}

// 读取本地图片文件 → dataURL
function readLocalImage(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function cropWizardBackgroundDataUrl(dataUrl: string): Promise<string | null> {
  if (!template?.slideSize?.width || !template.slideSize.height) return dataUrl;
  const wPx = Math.round(template.slideSize.width * 96);
  const hPx = Math.round(template.slideSize.height * 96);
  const crop = await openImageCropEditor({
    imageDataUrl: dataUrl,
    frameRatio: template.slideSize.width / template.slideSize.height,
    frameSizeLabel: `页面背景：${wPx} × ${hPx} px`
  });
  return !crop.canceled && crop.dataUrl ? crop.dataUrl : null;
}

async function cropWizardBackground(container: HTMLElement): Promise<void> {
  if (!backgroundState.customImageDataUrl) return;
  try {
    const cropped = await cropWizardBackgroundDataUrl(backgroundState.customImageDataUrl);
    if (!cropped) return;
    backgroundState = { followDocument: false, customImageDataUrl: cropped };
    const thumb = container.querySelector('#wb-bg-thumb') as HTMLElement | null;
    if (thumb) {
      thumb.innerHTML = '<img src="' + escapeAttr(cropped) + '" alt="背景预览" />'
        + '<button id="wb-bg-crop" class="secondary" title="按页面比例重新裁剪">✂ 裁剪</button>'
        + '<button id="wb-bg-clear" class="ghost">清除</button>';
      thumb.querySelector('#wb-bg-crop')?.addEventListener('click', () => { void cropWizardBackground(container); });
      thumb.querySelector('#wb-bg-clear')?.addEventListener('click', () => {
        backgroundState = defaultWizardBackgroundState(template);
        const follow = container.querySelector('#wb-bg-follow') as HTMLInputElement | null;
        const pick = container.querySelector('#wb-bg-pick') as HTMLButtonElement | null;
        const hint = container.querySelector('#wb-bg-hint') as HTMLElement | null;
        if (follow) follow.checked = true;
        if (pick) pick.disabled = true;
        if (hint) hint.textContent = '';
        thumb.innerHTML = '';
        thumb.style.display = 'none';
        schedulePreview(container);
      });
    }
    schedulePreview(container);
    showToast('已按页面比例裁剪背景 ✓', 1500);
  } catch {
    showToast('裁剪失败', 2000);
  }
}

// ================= 实时预览 =================
function refreshPreview(container: HTMLElement): void {
  const box = container.querySelector('#wb-preview-inner') as HTMLElement | null;
  if (!box || !template) return;
  box.innerHTML = renderPreview(cloneTemplateWithWizardBackground(template, backgroundState), images, texts, vars, tableData, tableFits, tableMerges);
  void refreshQuality(container); // 预览刷新后同步跑一遍质量检查
}
// 输入法组词守卫：document 级 composition 事件，全局只注册一次；
// 组词期间（imeComposing=true）所有输入框的渲染/预览刷新统一跳过，结束补一次收尾。
function initImeGuard(): void {
  if (imeInit) return;
  imeInit = true;
  document.addEventListener('compositionstart', () => { imeComposing = true; });
  document.addEventListener('compositionend', () => {
    imeComposing = false;
    if (imeDirty && workbenchContainer) {
      imeDirty = false;
      // 收尾刷新：先把所有可见 gen-ta 撑到内容高度，再刷预览（组词已结束，不会再抖动）
      const page = document.querySelector('#page');
      page?.querySelectorAll('textarea.gen-ta').forEach((ta) => autoResize(ta as HTMLTextAreaElement));
      schedulePreview(workbenchContainer);
    }
  });
}

function schedulePreview(container: HTMLElement): void {
  if (imeComposing) { imeDirty = true; return; } // 组词中：跳过刷新，compositionend 后补
  if (previewTimer) window.clearTimeout(previewTimer);
  previewTimer = window.setTimeout(() => {
    refreshTableFits(container); // 先按最新数据/合并重算各表 fit 并刷新信息行
    refreshPreview(container);   // 再按最新 fit 渲染预览
  }, 250);
}

// ================= 生成质量检查（预览下方小字面板）=================
// 把模板语义约束（maxChars/maxLines/minChars/preferredLength/required）从「控制 AI 生成」
// 升级为「检查最终排版」，并检查：文本溢出 / 行数·字数限制 / 图片留白 / 图片裁剪异常 / 表格超出 / 元素重叠。
async function refreshQuality(container: HTMLElement): Promise<void> {
  if (!template) return;
  if (imeComposing) { imeDirty = true; return; } // 组词中：跳过质量检查刷新
  const box = container.querySelector('#wb-quality') as HTMLElement | null;
  if (!box) return;
  const seq = ++qualitySeq;
  const report = await runQualityChecks(template, images, texts, vars, tableData, croppedImages, tableFits);
  if (seq !== qualitySeq) return; // 已被更新的刷新取代，丢弃过期结果
  lastQuality = report;
  box.innerHTML = qualityHtml(report);
}

function qualityHtml(r: QualityReport): string {
  const worstOf = (cat: string): QualityLevel => {
    let w: QualityLevel = 'pass';
    for (const x of r.issues) {
      if (x.category !== cat) continue;
      if (x.level === 'error') w = 'error';
      else if (x.level === 'warn' && w !== 'error') w = 'warn';
    }
    return w;
  };
  // 只显示有警告/异常的类别；通过的（✓）不占位置，全部通过时整个面板不渲染
  const active = QUALITY_CATEGORIES.filter((cat) => worstOf(cat) !== 'pass');
  if (!active.length) return '';
  const chip = (cat: string): string => {
    const w = worstOf(cat);
    const cls = w === 'error' ? 'err' : 'warn';
    const icon = w === 'error' ? '✗' : '⚠';
    return '<span class="wb-q-chip ' + cls + '">' + icon + ' ' + cat + '</span>';
  };
  const iconOf = (lv: QualityLevel): string => lv === 'error' ? '✗' : '⚠';
  const detail = r.issues.map((x) =>
    '<div class="wb-q-item ' + x.level + '">' + iconOf(x.level) + ' <span><b>' + escapeHtml(x.label) + '</b>：' + escapeHtml(x.message) + '</span></div>').join('');
  // 不显示「🔎 生成质量检查 / N 项警告」标题行，只留警告内容（类别 + 条目）
  return '<div class="wb-quality">' +
    '<div class="wb-quality-chips">' + active.map(chip).join('') + '</div>' +
    '<div class="wb-quality-detail">' + detail + '</div>' +
    '</div>';
}

// ================= 工作台渲染 =================
export async function renderWizard(container: HTMLElement): Promise<void> {
  const id = sessionStorage.getItem('templateId');
  if (!id) {
    container.innerHTML = '<h1 class="page-title">生成向导</h1><div class="card"><p>请先在「模板库」选择模板。</p><button class="primary" id="go-library">去模板库</button></div>';
    container.querySelector('#go-library')!.addEventListener('click', () => { location.hash = '#library'; });
    return;
  }
  // 模板列表（顶部下拉切换模板用；失败时静默降级为仅当前模板）
  let tplList: TemplateMeta[] = [];
  try { tplList = await Api.listTemplates(); } catch { /* 后端未就绪：仅当前模板 */ }
  try {
    const folder = sessionStorage.getItem('templateFolder') || '';
    const res = await Api.getTemplate(id, folder);
    template = res.template;
    template.shapes = sortShapesByPosition(template.shapes); // 读取顺序稳定：上到下、左到右
  } catch (e) {
    container.innerHTML = `<div class="card"><p class="error">加载模板失败：${escapeHtml((e as Error).message)}</p></div>`;
    return;
  }
  resetAll();
  // 读取「每次显示图片数量」与搜图供应商配置
  try {
    const cfg = await Api.getConfig();
    imagePageSize = normalizeImagePageSize(cfg.image?.pageSize);
    imgProvider = await getDefaultImageProvider();
  } catch { /* 后端未就绪时使用默认值 */ }
  renderWorkbench(container, tplList);
}

function resetAll(): void {
  images = {}; texts = {}; prompts = {}; formulaLatex = {}; vars = {};
  tableData = {}; aiResult = {}; pasteBuffers = {};
  tableMerges = {}; tableFits = {}; tableMergeDir = {}; lastTableMerges = {};
  imgStates = {}; imagePageSize = 9; croppedImages.clear(); globalTheme = '';
  backgroundState = defaultWizardBackgroundState(template);
  defaultOutputMode = loadOutputMode();
  slotOutputModes = {};
  expandedTables.clear(); previewTimer = undefined;
  lastQuality = null; qualitySeq = 0;
  clearImageSizeCache(); // 图片尺寸缓存随模板切换清空
}

function renderWorkbench(container: HTMLElement, tplList: TemplateMeta[]): void {
  if (!template) return;
  workbenchContainer = container; // compositionend 收尾刷新定位容器
  initImeGuard(); // 输入法组词守卫（全局只注册一次）
  const hasImage = template.shapes.some((s) => s.role === 'ai_image');
  const textSlots = sortTextSlots(template.shapes.filter((s) => s.role === 'ai_text' || s.role === 'manual_var'));
  const imageShapes = template.shapes.filter((s) => s.role === 'ai_image');
  const tableShapes = template.shapes.filter((s) => s.role === 'table' && s.table);
  // 表格位默认数据（保存的单元格文字）；首次进入时初始化
  tableShapes.forEach((s) => { if (!tableData[s.id]) tableData[s.id] = defaultTableData(s); });
  // 每个图片位初始化独立搜图状态
  imageShapes.forEach((s) => { if (!imgStates[s.id]) imgStates[s.id] = defaultImgState(imagePageSize); });
  // 提示词可见可用：模板自带的提示词自动填入「AI服务」输入框（未手动改过才填；公式位用内置提示词不预填）
  textSlots.forEach((s) => {
    if (s.prompt && !(s.id in prompts) && s.semanticRole !== 'formula') prompts[s.id] = s.prompt;
  });
  // 日期位：自动取当日日期（YYYY/MM/DD），无输入框
  textSlots.forEach((s) => {
    if (s.semanticRole === 'date' && !texts[s.id]) texts[s.id] = todayStr();
  });

  const currentId = sessionStorage.getItem('templateId') || '';
  const currentFolder = sessionStorage.getItem('templateFolder') || '';
  const tplOptions = tplList.length
    ? `<select id="wb-template" style="flex:1">${tplList.map((t) =>
        `<option value="${escapeAttr(t.id)}|${escapeAttr(t.folder)}"${t.id === currentId && t.folder === currentFolder ? ' selected' : ''}>${escapeHtml(t.name)}${t.folder ? '（' + escapeHtml(t.folder) + '）' : ''}</option>`).join('')}</select>`
    : `<span class="wb-tpl-name"><b>${escapeHtml(template.name)}</b></span>`;
  // 输出模式控件：移入每个「AI服务」折叠内（提示词文本框左下的 chips + 右侧 AI 生成按钮），不再常驻顶部

  container.innerHTML = `
    <h1 class="page-title">生成向导</h1>
    <div class="card wb-tpl-card">
      <label>模板</label>
      <div style="display:flex;gap:8px;align-items:center">${tplOptions}
        <button class="secondary" id="wb-tpl-back" title="返回模板库">模板库</button>
      </div>
    </div>
    <div class="card wb-write-bar">
      <button class="primary" id="write-ppt" style="width:100%">生成并插入 PPT</button>
    </div>
    <div class="card">
      <details class="wb-adv gp-adv"><summary>全局提示词${infoTip('作为各段 AI 生成的统一背景；不填时按各字段单独要求生成')}</summary>
        <input id="global-theme" placeholder="例如：输电线路异物清除机器人设计" style="margin-top:6px" />
      </details>
    </div>
    <div class="card">
      ${backgroundPanelHtml()}
    </div>
    ${textSlots.length ? `<div class="card wb-sec"><div class="wb-sec-title"><span>文字</span></div>${textSlots.map((s, i) => textSlotHtml(s, i)).join('')}</div>` : ''}
    ${hasImage ? `<div class="card wb-sec"><div class="wb-sec-title"><span>图片${infoTip('把本地图片拖到这里直接上传，或点「本地」选择')}</span></div>${imageShapes.map((s, i) => imgSlotHtml(s, i)).join('')}</div>` : ''}
    ${tableShapes.length ? `<div class="card wb-sec"><div class="wb-sec-title">表格</div><div id="wb-tables">${tableShapes.map((s, ti) => tableSlotHtml(s, ti)).join('')}</div></div>` : ''}
    <div class="card wb-preview-card">
      <div class="wb-preview-head">
        <b>实时预览${infoTip('输入文字 / 换图 / 改表格时自动刷新；点击预览可放大查看。')}</b>
        <span style="display:flex;gap:6px;align-items:center">
          <button class="secondary" id="wb-preview-toggle" title="收起 / 展开预览">${previewCollapsed ? '展开 ▴' : '收起 ▾'}</button>
          <button class="secondary" id="wb-preview-zoom" title="放大预览">放大</button>
        </span>
      </div>
      <div id="wb-preview-body" style="display:${previewCollapsed ? 'none' : 'block'}">
        <div id="wb-preview-inner"></div>
        <div id="wb-quality"></div>
      </div>
    </div>
`;
  refreshPreview(container);
  bindWorkbench(container);
}

// —— 文字位：按语义角色渲染不同输入 UI ——
// 主标题/副标题/正文/不指定：标签 + 输入框（占满整行）+ AI服务（提示词自动填入；AI 生成按钮收进折叠）
// 勾选「自动翻译」的位：无 AI 生成按钮，源文本输入完成后自动调用服务翻译；输入框占满整行
// 序号：一行内 6 个 01~06 按钮 + 1 个仅限数字的输入框（无 AI服务）；日期：自动取当日（只读）；
// 图片诠释：单输入框（图片搜索关键词自动 copy，无 AI服务）；公式：输入框 + AI服务（内置提示词转专业型）
function textSlotHtml(s: TemplateShape, i: number): string {
  const isVar = s.role === 'manual_var';
  const cur = isVar ? (vars[s.id] ?? '') : (texts[s.id] ?? '');
  const roleLabel = semanticRoleLabel(s.semanticRole);
  const label = isVar
    ? (s.varName ? '变量「' + escapeHtml(s.varName) + '」' : '变量 ' + (i + 1))
    : (roleLabel || '文本位 ' + (i + 1)) + (s.name ? '：' + escapeHtml(s.name) : '');
  const cons = constraintSummary(s);
  const sem = s.semanticRole;
  const isSpecial = sem === 'seq' || sem === 'date' || sem === 'caption' || sem === 'formula';
  const isTranslate = !isSpecial && !!(s.translate && s.translateSource);
  const tr = isTranslate ? ' <span class="wb-translate-tag">🔁 自动翻译</span>' : '';
  const lab = `<label class="wb-slot-label"><b>${label}</b>${cons ? ` <span class="wb-constraint">${escapeHtml(cons)}</span>` : ''}${tr}</label>`;
  const status = `<span class="hint slot-status" data-id="${s.id}"></span>`;
  // —— 序号：01~06 按钮 + 数字输入框（同一行；选中/输入即写入该位内容）——
  if (sem === 'seq') {
    return `
    <div class="wb-text-slot" data-id="${s.id}">
      ${lab}
      <div class="wb-seq-row" data-id="${s.id}">
        ${['01', '02', '03', '04', '05', '06'].map((n) => `<button type="button" class="om-chip seq-chip${cur === n ? ' active' : ''}" data-id="${s.id}" data-seq="${n}">${n}</button>`).join('')}
        <input class="seq-input" data-id="${s.id}" inputmode="numeric" pattern="[0-9]*" maxlength="6" placeholder="自定义" value="${escapeAttr(/^\d*$/.test(cur) ? cur : '')}" title="仅限阿拉伯数字" />
      </div>
      ${status}
    </div>`;
  }
  // —— 日期：自动取当日（只读，无输入框 / 无 AI服务）——
  if (sem === 'date') {
    return `
    <div class="wb-text-slot" data-id="${s.id}">
      ${lab}
      <div class="wb-date-val" data-id="${s.id}">📅 ${escapeHtml(texts[s.id] || todayStr())}</div>
      ${status}
    </div>`;
  }
  // —— 图片诠释：单输入框（无 AI服务；图片搜索时关键词自动 copy）——
  if (sem === 'caption') {
    return `
    <div class="wb-text-slot" data-id="${s.id}">
      ${lab}
      <div class="wb-slot-row">
        <input class="gen-in wb-gen" data-gen="${s.id}" value="${escapeAttr(cur)}" placeholder="填写图片诠释；图片搜索的关键词会自动填入（留空则不带）" />
      </div>
      ${status}
    </div>`;
  }
  // —— 公式：输入框（可粘贴）+ AI服务（内置提示词转完整专业型公式，输出 LaTeX）——
  if (sem === 'formula') {
    return `
    <div class="wb-text-slot" data-id="${s.id}">
      ${lab}
      <div class="wb-slot-row">
        <textarea class="gen-ta wb-gen" data-gen="${s.id}" rows="2" placeholder="粘贴或输入公式，如 𝑇_allow=(𝜋[𝜏] 𝑑^3)/(16𝐾_𝜏×10^3 )">${escapeAttr(cur)}</textarea>
      </div>
      <details class="wb-adv"><summary>AI服务</summary>
        <p class="hint" style="margin:4px 0">内置提示词：把输入的公式/描述转换为完整规范的专业型公式（输出 LaTeX），生成 PPT 时转为专业型数学排版。</p>
        <div class="wb-adv-actions"><button class="secondary slot-gen" data-id="${s.id}" title="用 AI 把上面的公式转成完整专业型公式">AI 生成公式</button></div>
      </details>
      ${status}
    </div>`;
  }
  // —— 普通文字位（主标题/副标题/正文/不指定）：原有 UI ——
  const genBtn = isTranslate ? '' : `<button class="secondary slot-gen" data-id="${s.id}" title="用主题（或下方单独要求）生成这一段">AI 生成</button>`;
  // 输出模式 chips（整段/分点/精简/限字数）：每个 AI服务 折叠内、提示词文本框左下方；AI 生成按钮在同排右侧
  const slotMode = modeForSlot(s.id);
  const modeChips = `<span class="wb-mode-chips">
        <button type="button" class="om-chip${slotMode.plain ? ' active' : ''}" data-slot="${s.id}" data-om="plain" title="整段：一段连续正文，不用 Markdown 和列表符号">整段</button>
        <button type="button" class="om-chip${slotMode.bullets ? ' active' : ''}" data-slot="${s.id}" data-om="bullets" title="分点：每个要点单独一行">分点</button>
        <button type="button" class="om-chip${slotMode.condense ? ' active' : ''}" data-slot="${s.id}" data-om="condense" title="精简：删冗余、句子短">精简</button>
        <select class="om-limit" data-slot="${s.id}" title="限制 AI 输出总字数">${LIMIT_CHOICES.map((n) => `<option value="${n}"${slotMode.maxChars === n ? ' selected' : ''}>${n === 0 ? '不限字数' : '≤' + n + ' 字'}</option>`).join('')}</select>
      </span>`;
  const adv = isTranslate ? '' : `
    <details class="wb-adv"><summary>AI服务</summary>
      <textarea class="slot-prompt" data-id="${s.id}" rows="2" placeholder="为该段单独写要求（可选），留空则用主题生成">${escapeAttr(prompts[s.id] || '')}</textarea>
      <div class="wb-adv-row">
        ${modeChips}
        <span class="wb-adv-actions">${genBtn}</span>
      </div>
    </details>`;
  return `
  <div class="wb-text-slot" data-id="${s.id}">
    ${lab}
    <div class="wb-slot-row">
      <textarea class="gen-ta wb-gen" data-gen="${s.id}" rows="1" placeholder="${isTranslate ? '完成原文输入后自动翻译，也可手动输入译文' : '输入内容，或展开「AI服务」生成'}">${escapeAttr(cur)}</textarea>
    </div>
    ${adv}
    ${status}
  </div>`;
}

// —— 图片位：已选图预览 + 本地/搜索/拖拽 三种方式 ——
function imgSlotHtml(s: TemplateShape, i: number): string {
  const st = imgStates[s.id] || defaultImgState();
  const chosen = images[s.id] || '';
  return `
  <div class="img-slot wb-img-slot" data-id="${s.id}">
    <b>图片位 ${i + 1}${s.name ? '：' + escapeHtml(s.name) : ''}</b>
    <div class="img-chosen" data-id="${s.id}" style="margin:4px 0${chosen ? '' : ';display:none'}">
      <img src="${chosen}" style="max-width:120px;max-height:80px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />
      <button class="secondary img-crop" data-id="${s.id}" title="按图片位比例重新裁剪">✂ 裁剪</button>
      <button class="secondary img-clear" data-id="${s.id}">✕ 清除</button>
    </div>
    <div class="wb-img-actions">
      <button class="secondary img-pick" data-id="${s.id}" title="从文件夹选择本地图片">本地</button>
      <input class="img-q" data-id="${s.id}" value="${escapeAttr(st.q)}" placeholder="描述图片，如：深蓝色科技感机房" />
      <button class="secondary img-search" data-id="${s.id}">搜索</button>
    </div>
    <input type="file" class="img-file" data-id="${s.id}" accept="image/*" style="display:none" />
    <div class="img-dropzone" data-id="${s.id}"></div>
    <div class="img-results" data-id="${s.id}" style="display:grid;grid-template-columns:repeat(3,1fr);gap:4px"></div>
    <div class="img-nav" data-id="${s.id}"></div>
  </div>`;
}

function backgroundPanelHtml(): string {
  if (!template) return '';
  const bg = template.background;
  const bgDesc = bg
    ? (bg.type === 'solid' && bg.color ? '纯色（' + bg.color + '）' : bg.type === 'picture' ? '图片' : bg.type)
    : '无';
  const checked = backgroundState.followDocument ? ' checked' : '';
  const disabled = backgroundState.followDocument ? ' disabled' : '';
  const custom = backgroundState.customImageDataUrl;
  const thumb = custom
    ? '<div id="wb-bg-thumb" class="wb-bg-thumb"><img src="' + escapeAttr(custom) + '" alt="背景预览" /><button id="wb-bg-crop" class="secondary" title="按页面比例重新裁剪">✂ 裁剪</button><button id="wb-bg-clear" class="ghost">清除</button></div>'
    : '<div id="wb-bg-thumb" class="wb-bg-thumb" style="display:none"></div>';
  return `
    <details class="wb-adv gp-adv wb-bg-adv">
      <summary>背景${infoTip('默认跟随模板保存的文档背景；取消勾选后可选择本地图片作为本次生成背景。')}</summary>
      <div class="page-info wb-bg-panel">
        <div class="pi-row">
          <span class="pi-k">图片</span>
          <span class="pi-v wb-bg-desc" id="wb-bg-desc">${escapeHtml(bgDesc)}</span>
          <span class="flex-spacer"></span>
          <label class="fixed-cb-wrap wb-bg-follow">跟随文档 <input type="checkbox" id="wb-bg-follow"${checked} /></label>
          <button id="wb-bg-pick" class="ghost"${disabled}>文件夹选择</button>
          <input type="file" id="wb-bg-file" accept="image/*" style="display:none" />
        </div>
        <p class="pi-hint" id="wb-bg-hint">${backgroundState.followDocument ? '' : (custom ? '已选择本地背景图片' : '请选择一张背景图片')}</p>
        ${thumb}
      </div>
    </details>`;
}

// —— 表格位：摘要卡片（已导入 RxC + fit 信息行）+ 可展开编辑（手动/粘贴/AI + 一键合并）——
function tableSlotHtml(s: TemplateShape, ti: number): string {
  const data = tableData[s.id] || defaultTableData(s);
  const nc = Math.max(1, (data[0] || []).length);
  const aiRes = aiResult[s.id];
  const expanded = expandedTables.has(s.id);
  const fit = getTableFit(s); // 计算/复用 fit（信息行 + 预览 + 提交复用）
  const merges = tableMerges[s.id];
  const mergeCount = merges ? merges.filter((c) => c.rowspan > 1 || c.colspan > 1).length : 0;
  const dirSel = tableMergeDir[s.id] || 'auto';
  return `
  <div class="wb-tbl-slot" data-tbl="${s.id}">
    <div class="wb-tbl-summary">
      <span class="wb-tbl-info"><b>表格 ${ti + 1}${s.name ? '：' + escapeHtml(s.name) : ''}</b>　<span class="wb-tbl-size">已导入 ${data.length} × ${nc} 数据</span></span>
      <button class="secondary tbl-toggle" data-tbl="${s.id}">${expanded ? '收起 ▲' : '展开编辑 ▾'}</button>
    </div>
    <div class="wb-tbl-fit" data-tbl="${s.id}">${tableFitInfoHtml(s, fit)}</div>
    <div class="wb-tbl-editor" data-tbl="${s.id}" style="display:${expanded ? 'block' : 'none'}">
      ${mergeCount ? '<div class="hint" style="margin:2px 0">该表格含 ' + mergeCount + ' 处合并单元格（预览按合并渲染；逐格编辑以左上主格文字为准）</div>' : ''}
      <div class="wb-tbl-merge-row">
        <span class="hint">相同字样一键合并：</span>
        <select class="tbl-merge-dir" data-tbl="${s.id}" title="合并方向：自动 = 先横向后纵向">
          <option value="auto"${dirSel === 'auto' ? ' selected' : ''}>自动</option>
          <option value="vertical"${dirSel === 'vertical' ? ' selected' : ''}>纵向</option>
          <option value="horizontal"${dirSel === 'horizontal' ? ' selected' : ''}>横向</option>
        </select>
        <button class="secondary tbl-merge-do" data-tbl="${s.id}" title="把内容完全一致的相邻单元格合并（trim 后相同、连续 ≥2 格）">一键合并</button>
        <button class="secondary tbl-merge-undo" data-tbl="${s.id}" title="撤销最近一次一键合并">撤销</button>
      </div>
      <div class="tbl-tabs" data-tbl="${s.id}" style="display:flex;gap:4px;margin:6px 0">
        <button class="secondary tbl-tab" data-tbl="${s.id}" data-mode="manual">手动编辑</button>
        <button class="secondary tbl-tab active" data-tbl="${s.id}" data-mode="paste">粘贴</button>
        <button class="secondary tbl-tab" data-tbl="${s.id}" data-mode="ai">AI 生成</button>
      </div>
      <div class="tbl-panel" data-tbl="${s.id}" data-mode="manual" style="display:none">
        <div class="tbl-grid" style="display:grid;grid-template-columns:repeat(${nc}, 1fr);gap:2px;margin:4px 0">
          ${data.map((row, r) => row.map((cell, c) =>
            `<textarea class="tbl-cell" data-tbl="${s.id}" data-r="${r}" data-c="${c}" rows="1" placeholder=" ">${escapeAttr(cell)}</textarea>`).join('')).join('')}
        </div>
      </div>
      <div class="tbl-panel" data-tbl="${s.id}" data-mode="paste" style="display:block">
        <textarea class="tbl-paste-ta" data-tbl="${s.id}" rows="6" placeholder="在此 Ctrl+V 粘贴 Excel/网页复制的表格（优先读剪贴板 HTML，自动识别合并格；无 HTML 时按制表符/逗号分隔解析）">${escapeAttr(pasteBuffers[s.id] || '')}</textarea>
        <div style="display:flex;gap:6px;margin-top:6px">
          <button class="secondary tbl-paste-do" data-tbl="${s.id}">导入到表格</button>
          <button class="secondary tbl-paste-clear" data-tbl="${s.id}">清空全部单元格</button>
        </div>
      </div>
      <div class="tbl-panel" data-tbl="${s.id}" data-mode="ai" style="display:none">
        <textarea class="tbl-ai-req" data-tbl="${s.id}" rows="2" placeholder="生成要求（可选），如：补充 5 行近三年各季度营收数据，数值用万元"></textarea>
        <div style="display:flex;gap:6px;margin-top:6px">
          <button class="secondary tbl-ai-do" data-tbl="${s.id}">AI 生成数据</button>
        </div>
        <div class="tbl-ai-out" data-tbl="${s.id}">${aiRes ? aiPreviewHtml(aiRes, s.id) : ''}</div>
      </div>
    </div>
  </div>`;
}

// 局部重绘表格区（粘贴导入/清空/AI 应用后行列变化），不影响其它编辑区状态
function rerenderTables(container: HTMLElement): void {
  if (!template) return;
  const box = container.querySelector('#wb-tables') as HTMLElement | null;
  if (!box) return;
  const tableShapes = template.shapes.filter((s) => s.role === 'table' && s.table);
  box.innerHTML = tableShapes.map((s, ti) => tableSlotHtml(s, ti)).join('');
  bindTableEvents(container);
}
// ================= 事件绑定 =================
function bindWorkbench(container: HTMLElement): void {
  if (!template) return;
  bindTemplateSelect(container);
  bindBackgroundEvents(container);
  bindOutputMode(container);
  bindTextEvents(container);
  bindImageEvents(container);
  bindTableEvents(container);
  bindPreviewAndWrite(container);
}

function bindBackgroundEvents(container: HTMLElement): void {
  if (!template) return;
  const follow = container.querySelector('#wb-bg-follow') as HTMLInputElement | null;
  const pick = container.querySelector('#wb-bg-pick') as HTMLButtonElement | null;
  const file = container.querySelector('#wb-bg-file') as HTMLInputElement | null;
  const hint = container.querySelector('#wb-bg-hint') as HTMLElement | null;
  const thumb = container.querySelector('#wb-bg-thumb') as HTMLElement | null;
  const sync = (): void => {
    if (follow) follow.checked = backgroundState.followDocument;
    if (pick) pick.disabled = backgroundState.followDocument;
    if (hint) hint.textContent = backgroundState.followDocument
      ? ''
      : (backgroundState.customImageDataUrl ? '已选择本地背景图片' : '请选择一张背景图片');
    if (thumb) {
      if (backgroundState.customImageDataUrl) {
        thumb.style.display = '';
        thumb.innerHTML = '<img src="' + escapeAttr(backgroundState.customImageDataUrl) + '" alt="背景预览" />'
          + '<button id="wb-bg-crop" class="secondary" title="按页面比例重新裁剪">✂ 裁剪</button>'
          + '<button id="wb-bg-clear" class="ghost">清除</button>';
        thumb.querySelector('#wb-bg-crop')?.addEventListener('click', () => { void cropWizardBackground(container); });
        thumb.querySelector('#wb-bg-clear')?.addEventListener('click', () => {
          backgroundState = defaultWizardBackgroundState(template);
          if (file) file.value = '';
          sync();
          schedulePreview(container);
        });
      } else {
        thumb.innerHTML = '';
        thumb.style.display = 'none';
      }
    }
  };
  follow?.addEventListener('change', () => {
    backgroundState.followDocument = follow.checked;
    if (backgroundState.followDocument) backgroundState.customImageDataUrl = '';
    sync();
    schedulePreview(container);
  });
  pick?.addEventListener('click', () => file?.click());
  file?.addEventListener('change', () => {
    const f = file.files?.[0];
    if (!f) return;
    void readLocalImage(f).then(async (dataUrl) => {
      const cropped = await cropWizardBackgroundDataUrl(dataUrl);
      if (!cropped) return;
      backgroundState = { followDocument: false, customImageDataUrl: cropped };
      sync();
      schedulePreview(container);
      showToast('已选择本地背景 ✓', 1500);
    }).catch(() => alert('读取图片失败')).finally(() => { file.value = ''; });
  });
  sync();
}

// —— 输出模式 chips：每个文本位独立，整段/分点/精简 互斥与组合 + 限字数下拉 ——
function bindOutputMode(container: HTMLElement): void {
  const chips = Array.from(container.querySelectorAll<HTMLElement>('.om-chip'));
  const limits = Array.from(container.querySelectorAll<HTMLSelectElement>('.om-limit'));
  const refresh = (slotId?: string) => {
    chips.forEach((c) => {
      const k = c.dataset.om as 'plain' | 'bullets' | 'condense' | undefined;
      if (!k) return; // 非输出模式 chip（如序号位 01~06 按钮共用 .om-chip 类）
      const sid = c.dataset.slot || '';
      if (slotId && sid !== slotId) return;
      c.classList.toggle('active', !!modeForSlot(sid)[k]);
    });
    limits.forEach((l) => {
      const sid = l.dataset.slot || '';
      if (slotId && sid !== slotId) return;
      l.value = String(modeForSlot(sid).maxChars);
    });
  };
  chips.forEach((c) => {
    c.addEventListener('click', () => {
      const k = c.dataset.om as 'plain' | 'bullets' | 'condense' | undefined;
      if (!k) return; // 序号按钮：不参与输出模式
      const sid = c.dataset.slot || '';
      const mode = modeForSlot(sid);
      if (k === 'plain') { mode.plain = true; mode.bullets = false; }
      else if (k === 'bullets') { mode.bullets = true; mode.plain = false; }
      else mode[k] = !mode[k];
      mode.touched = true;
      refresh(sid);
    });
  });
  limits.forEach((l) => {
    l.addEventListener('change', () => {
      const sid = l.dataset.slot || '';
      const mode = modeForSlot(sid);
      mode.maxChars = Number(l.value) || 0;
      mode.touched = true;
      refresh(sid);
    });
  });
  refresh();
}

function bindTemplateSelect(container: HTMLElement): void {
  const sel = container.querySelector('#wb-template') as HTMLSelectElement | null;
  sel?.addEventListener('change', async () => {
    const [id, folder] = (sel.value || '').split('|');
    if (!id) return;
    sessionStorage.setItem('templateId', id);
    sessionStorage.setItem('templateFolder', folder || '');
    try {
      const res = await Api.getTemplate(id, folder || '');
      template = res.template;
      template.shapes = sortShapesByPosition(template.shapes); // 读取顺序稳定：上到下、左到右
    } catch (e) {
      showToast('加载模板失败：' + (e as Error).message, 3000);
      return;
    }
    resetAll();
    // 重新读取配置（pageSize/provider）
    try {
      const cfg = await Api.getConfig();
      imagePageSize = normalizeImagePageSize(cfg.image?.pageSize);
      for (const k of Object.keys(imgStates)) imgStates[k].pageSize = imagePageSize;
      if (cfg.image?.provider) {
        try {
          const pl = await Api.getImageProviders();
          if (pl.providers.some((p) => p.id === cfg.image.provider)) imgProvider = cfg.image.provider;
        } catch { /* 注册表不可达：保留默认 */ }
      }
    } catch { /* 默认值 */ }
    try {
      const list = await Api.listTemplates();
      renderWorkbench(container, list);
    } catch {
      renderWorkbench(container, []);
    }
  });
  const back = container.querySelector('#wb-tpl-back') as HTMLElement | null;
  back?.addEventListener('click', () => { location.hash = '#library'; });
}

function bindTextEvents(container: HTMLElement): void {
  if (!template) return;
  const themeInput = container.querySelector('#global-theme') as HTMLInputElement | null;
  if (themeInput) {
    themeInput.value = globalTheme; // 进入/切模板后保留
    themeInput.addEventListener('input', () => { globalTheme = themeInput.value; });
  }
  const textSlots = sortTextSlots(template.shapes.filter((s) => s.role === 'ai_text' || s.role === 'manual_var'));
  textSlots.forEach((s) => {
    const isVar = s.role === 'manual_var';
    const pta = container.querySelector(`textarea.slot-prompt[data-id="${s.id}"]`) as HTMLTextAreaElement | null;
    pta?.addEventListener('input', () => { prompts[s.id] = pta.value; });
    const ta = container.querySelector(`textarea.gen-ta[data-gen="${s.id}"]`) as HTMLTextAreaElement | null;
    if (ta) {
      ta.addEventListener('input', (e) => {
        if (isVar) vars[s.id] = ta.value; else texts[s.id] = ta.value;
        // 输入法组词中（isComposing 或全局 composition 标记）：跳过 autoResize 与预览刷新，
        // 避免布局抖动导致候选框消失/错位；组词结束后由 compositionend 统一补一次收尾
        if (imeComposing || (e as InputEvent).isComposing === true) { imeDirty = true; return; }
        autoResize(ta);
        schedulePreview(container); // 输入即刷新实时预览
      });
      autoResize(ta);
    }
    // —— 序号位：01~06 按钮 + 仅数字输入框（选中/输入即写入该位内容）——
    if (s.semanticRole === 'seq') {
      const chips = container.querySelectorAll(`button.seq-chip[data-id="${s.id}"]`);
      chips.forEach((btn) => {
        btn.addEventListener('click', () => {
          const n = (btn as HTMLButtonElement).getAttribute('data-seq') || '';
          texts[s.id] = n;
          const inp = container.querySelector(`input.seq-input[data-id="${s.id}"]`) as HTMLInputElement | null;
          if (inp) inp.value = n;
          chips.forEach((b) => b.classList.toggle('active', b === btn));
          schedulePreview(container);
        });
      });
      const seqInp = container.querySelector(`input.seq-input[data-id="${s.id}"]`) as HTMLInputElement | null;
      const clearChips = () => chips.forEach((b) => b.classList.remove('active'));
      seqInp?.addEventListener('input', () => {
        const v = seqInp.value.replace(/\D/g, ''); // 仅限阿拉伯数字
        if (seqInp.value !== v) seqInp.value = v;
        texts[s.id] = v;
        clearChips();
        schedulePreview(container);
      });
      seqInp?.addEventListener('paste', (e) => {
        e.preventDefault();
        const clip = (e as ClipboardEvent).clipboardData || ((window as unknown as { clipboardData?: DataTransfer }).clipboardData);
        const t = (clip ? clip.getData('text') : '').replace(/\D/g, '');
        seqInp.value = t;
        texts[s.id] = t;
        clearChips();
        schedulePreview(container);
      });
    }
    // —— 图片诠释位：单输入框（无 AI服务）——
    const cin = container.querySelector(`input.gen-in[data-gen="${s.id}"]`) as HTMLInputElement | null;
    cin?.addEventListener('input', () => { texts[s.id] = cin.value; schedulePreview(container); });
    const genBtn = container.querySelector(`button.slot-gen[data-id="${s.id}"]`) as HTMLButtonElement | null;
    const statusEl = container.querySelector(`.slot-status[data-id="${s.id}"]`) as HTMLElement | null;
    genBtn?.addEventListener('click', async () => {
      // —— 公式位：内置提示词把输入转成完整专业型公式（LaTeX），生成 PPT 时转为专业型排版 ——
      if (s.semanticRole === 'formula') {
        const cfg0 = await Api.getConfig().catch(() => null);
        if (!cfg0 || !cfg0.text?.apiKey) {
          const choice = await showModal({
            title: '文本 AI 未配置',
            message: '尚未配置文本生成服务（API Key），无法把公式转为专业型。也可以直接粘贴 LaTeX 公式，生成时会本地转换。',
            buttons: [
              { id: 'ok', label: '知道了', kind: 'primary' },
              { id: 'config', label: '前往配置', kind: 'danger' }
            ]
          });
          if (choice === 'config') location.hash = '#settings';
          return;
        }
        genBtn.disabled = true;
        const genText = genBtn.textContent || '';
        genBtn.textContent = '转换中…';
        if (statusEl) statusEl.textContent = '';
        try {
          const raw = texts[s.id] || '';
          const { text: t } = await Api.generateText(
            '你是数学公式转换助手。把用户给出的公式文本或描述转换为完整、规范、可直接使用的专业型数学公式，输出 LaTeX 代码（正确使用 \\frac、下标 _{}、上标 ^{}、\\pi、\\tau、\\times 等语法）。只输出 LaTeX 本体，不要任何解释、不要 markdown 代码块标记。',
            raw ? '请把下面的公式转换为完整专业型 LaTeX：\n' + raw : '请生成一个完整的专业型数学公式（LaTeX）',
            shapeConstraints(s)
          );
          const latex = t.trim().replace(/^```(latex)?/i, '').replace(/```$/, '').trim();
          if (latex) {
            formulaLatex[s.id] = latex;
            if (statusEl) statusEl.textContent = '✓ 已生成完整公式（生成 PPT 时转为专业型）';
          } else {
            if (statusEl) statusEl.textContent = '✗ 未得到公式';
          }
        } catch (e) {
          if (statusEl) statusEl.textContent = '✗ 转换失败：' + ((e as Error).message);
        } finally {
          genBtn.disabled = false;
          genBtn.textContent = genText;
        }
        return;
      }
      const theme = globalTheme;
      // 未配置文本 AI：提示前往配置
      const cfg = await Api.getConfig().catch(() => null);
      if (!cfg || !cfg.text?.apiKey) {
        const choice = await showModal({
          title: '文本 AI 未配置',
          message: '尚未配置文本生成服务（API Key），无法自动生成文字。请先前往配置。',
          buttons: [
            { id: 'ok', label: '知道了', kind: 'primary' },
            { id: 'config', label: '前往配置', kind: 'danger' }
          ]
        });
        if (choice === 'config') location.hash = '#settings';
        return;
      }
      genBtn.disabled = true;
      const genBtnText = genBtn.textContent || '';
      genBtn.textContent = '生成中…';
      if (statusEl) statusEl.textContent = '';
      try {
        let text = '';
        // 输出模式：解析该文本位实际生效的风格指令与清洗规则（整段/分点/精简/限字数）
        const { instruction, clean } = resolveSlotMode(modeForSlot(s.id), s.semanticRole);
        if (s.translate && s.translateSource) {
          const srcText = resolveTranslateText(s, theme);
          if (!srcText) {
            showToast('「' + (s.prompt || '翻译位') + '」原文为空，请先填写/生成原文', 3000);
          } else {
            const { text: t } = await Api.generateText(
              '你是 PPT 英文副标题翻译助手。把给定原文翻译成简洁、地道的英文副标题；保留专有名词、术语、数字与单位；不添加解释、不逐字直译，只输出译文。' + (instruction ? '\n' + instruction : ''),
              '请把下面的原文翻译为英文副标题，只输出译文：\n' + srcText,
              shapeConstraints(s),
              clean
            );
            text = t.trim();
          }
        } else {
          const per = prompts[s.id]?.trim() || '';
          const userPrompt = per || `请围绕主题「${theme || '模板主题'}」生成`;
          const { text: t } = await Api.generateText(
            '你是 PPT 文案助手，输出简洁专业的中文文案，不要多余解释。' + (instruction ? '\n' + instruction : ''),
            `${userPrompt}\n段落要求：${s.prompt || '内容简洁，符合演示文稿风格'}`,
            shapeConstraints(s),
            clean
          );
          text = t.trim();
        }
        if (text) {
          if (isVar) vars[s.id] = text; else texts[s.id] = text;
          if (ta) { ta.value = text; autoResize(ta); }
          if (statusEl) statusEl.textContent = '✓ 已生成';
          schedulePreview(container);
        }
      } catch (e) {
        if (statusEl) statusEl.textContent = '✗ 生成失败：' + ((e as Error).message);
      } finally {
        genBtn.disabled = false;
        genBtn.textContent = genBtnText;
      }
    });
    // —— 自动翻译副标题（translate 位）：源文本变化自动调用 AI 翻译（防抖 700ms）——
    const translateSlots = textSlots.filter((s) => s.translate && s.translateSource);
    if (translateSlots.length) {
      const timers: Record<string, number> = {};
      const debounceTranslate = (s: TemplateShape, statusEl: HTMLElement | null) => {
        if (timers[s.id]) window.clearTimeout(timers[s.id]);
        timers[s.id] = window.setTimeout(() => { void runAutoTranslate(s, statusEl); }, 700);
      };
      const runAutoTranslate = async (s: TemplateShape, statusEl: HTMLElement | null) => {
        if (statusEl) statusEl.textContent = '';
        const srcText = resolveTranslateText(s, globalTheme);
        if (!srcText) return;
        const cfg = await Api.getConfig().catch(() => null);
        if (!cfg || !cfg.text?.apiKey) { if (statusEl) statusEl.textContent = '⚠ 未配置文本 AI，无法自动翻译'; return; }
        if (statusEl) statusEl.textContent = '🔁 自动翻译中…';
        try {
          const { instruction, clean } = resolveSlotMode(modeForSlot(s.id), s.semanticRole);
          const { text: t } = await Api.generateText(
            '你是 PPT 英文副标题翻译助手。把给定原文翻译成简洁、地道的英文副标题；保留专有名词、术语、数字与单位；不添加解释、不逐字直译，只输出译文。' + (instruction ? '\n' + instruction : ''),
            '请把下面的原文翻译为英文副标题，只输出译文：\n' + srcText,
            shapeConstraints(s),
            clean
          );
          const h = t.trim();
          if (h) {
            texts[s.id] = h;
            const taEl = container.querySelector(`textarea.gen-ta[data-gen="${s.id}"]`) as HTMLTextAreaElement | null;
            if (taEl) { taEl.value = h; autoResize(taEl); }
            if (statusEl) statusEl.textContent = '✓ 已自动翻译';
            schedulePreview(container);
          }
        } catch (e) {
          if (statusEl) statusEl.textContent = '✗ 自动翻译失败：' + ((e as Error).message);
        }
      };
      for (const s of translateSlots) {
        const stEl = container.querySelector(`.slot-status[data-id="${s.id}"]`) as HTMLElement | null;
        if (s.translateSource === 'theme') {
          themeInput?.addEventListener('input', () => debounceTranslate(s, stEl));
        } else {
          const src = findTranslateSource(s);
          if (src) {
            const srcTa = container.querySelector(`textarea.gen-ta[data-gen="${src.id}"]`) as HTMLTextAreaElement | null;
            srcTa?.addEventListener('input', () => debounceTranslate(s, stEl));
            const srcPta = container.querySelector(`textarea.slot-prompt[data-id="${src.id}"]`) as HTMLTextAreaElement | null;
            srcPta?.addEventListener('input', () => debounceTranslate(s, stEl));
          }
        }
      }
    }
  });
}

// —— 图片位：本地选图（按钮/拖拽）+ 搜图 + 清除 + 裁剪 ——
function bindImageEvents(container: HTMLElement): void {
  if (!template) return;
  template.shapes.filter((s) => s.role === 'ai_image').forEach((s) => {
    const id = s.id;
    const pickBtn = container.querySelector(`button.img-pick[data-id="${id}"]`) as HTMLButtonElement | null;
    const fileInp = container.querySelector(`input.img-file[data-id="${id}"]`) as HTMLInputElement | null;
    const dropzone = container.querySelector(`.img-dropzone[data-id="${id}"]`) as HTMLElement | null;
    pickBtn?.addEventListener('click', () => fileInp?.click());
    dropzone?.addEventListener('click', () => fileInp?.click());
    fileInp?.addEventListener('change', () => {
      const f = fileInp.files?.[0];
      if (f) void readLocalImage(f).then((dataUrl) => setImage(container, id, dataUrl, '本地图片'));
      fileInp.value = '';
    });
    dropzone?.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.style.borderColor = '#1f3864'; });
    dropzone?.addEventListener('dragleave', () => { dropzone.style.borderColor = '#ccc'; });
    dropzone?.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.style.borderColor = '#ccc';
      const f = e.dataTransfer?.files?.[0];
      if (f && /^image\//.test(f.type)) void readLocalImage(f).then((dataUrl) => setImage(container, id, dataUrl, '本地图片'));
      else showToast('请拖入图片文件', 1800);
    });
    const qInp = container.querySelector(`input.img-q[data-id="${id}"]`) as HTMLInputElement | null;
    const searchBtn = container.querySelector(`button.img-search[data-id="${id}"]`) as HTMLButtonElement | null;
    searchBtn?.addEventListener('click', async () => {
      const q = (qInp?.value || '').trim();
      if (!q) { markInputError(qInp, '！需要输入内容！请描述你想要的图片'); return; }
      clearInputError(qInp);
      const st = (imgStates[id] = imgStates[id] || defaultImgState());
      st.q = q;
      st.page = 1;
      await doSearch(container, id);
    });
    const clearBtn = container.querySelector(`button.img-clear[data-id="${id}"]`) as HTMLButtonElement | null;
    clearBtn?.addEventListener('click', () => {
      delete images[id];
      const st = imgStates[id];
      if (st) st.selected = '';
      renderChosen(container, id);
      schedulePreview(container);
    });
    const cropBtn = container.querySelector(`button.img-crop[data-id="${id}"]`) as HTMLButtonElement | null;
    cropBtn?.addEventListener('click', () => { void cropImage(container, id); });
  });
}

// —— 表格位：展开/编辑、单元格输入、粘贴导入（HTML/CSV）、一键合并/撤销、AI 生成数据 ——
function bindTableEvents(container: HTMLElement): void {
  if (!template) return;
  container.querySelectorAll('button.tbl-toggle').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl');
      const editor = tbl ? container.querySelector(`.wb-tbl-editor[data-tbl="${tbl}"]`) as HTMLElement | null : null;
      if (!editor || !tbl) return;
      const hidden = editor.style.display === 'none';
      editor.style.display = hidden ? 'block' : 'none';
      if (hidden) expandedTables.add(tbl); else expandedTables.delete(tbl);
      btn.textContent = hidden ? '收起 ▲' : '展开编辑 ▾';
    });
  });
  container.querySelectorAll('button.tbl-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl');
      const mode = (btn as HTMLButtonElement).getAttribute('data-mode');
      container.querySelectorAll(`button.tbl-tab[data-tbl="${tbl}"]`).forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      container.querySelectorAll(`div.tbl-panel[data-tbl="${tbl}"]`).forEach((panel) => {
        (panel as HTMLElement).style.display = (panel as HTMLElement).getAttribute('data-mode') === mode ? 'block' : 'none';
      });
    });
  });
  container.querySelectorAll('textarea.tbl-cell').forEach((ta) => {
    ta.addEventListener('input', () => {
      const tbl = (ta as HTMLTextAreaElement).getAttribute('data-tbl') || '';
      const r = Number((ta as HTMLTextAreaElement).getAttribute('data-r'));
      const c = Number((ta as HTMLTextAreaElement).getAttribute('data-c'));
      if (!tableData[tbl] || !tableData[tbl][r]) return;
      tableData[tbl][r][c] = (ta as HTMLTextAreaElement).value;
      schedulePreview(container);
    });
  });
  container.querySelectorAll('textarea.tbl-paste-ta').forEach((ta) => {
    ta.addEventListener('input', () => { pasteBuffers[(ta as HTMLTextAreaElement).getAttribute('data-tbl') || ''] = (ta as HTMLTextAreaElement).value; });
    ta.addEventListener('paste', (e) => {
      const tbl = (ta as HTMLTextAreaElement).getAttribute('data-tbl') || '';
      const clip = (e as ClipboardEvent).clipboardData;
      if (!clip) return;
      const html = clip.getData('text/html');
      const plain = clip.getData('text/plain');
      if (html) {
        const cells = parseTableHtml(html);
        if (cells.length) {
          e.preventDefault();
          importFitCells(tbl, cells, plain);
          rerenderTables(container);
          schedulePreview(container);
          return;
        }
      }
      const grid = parseTableCsv(plain);
      if (grid.length) {
        e.preventDefault();
        importFromGrid(tbl, grid, plain);
        rerenderTables(container);
        schedulePreview(container);
      }
    });
  });
  container.querySelectorAll('button.tbl-paste-do').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl') || '';
      const ta = container.querySelector(`textarea.tbl-paste-ta[data-tbl="${tbl}"]`) as HTMLTextAreaElement | null;
      const raw = (ta?.value || '').trim();
      if (!raw) { showToast('请先在上方粘贴内容', 2000); return; }
      const grid = parseTableCsv(raw);
      if (!grid.length) { showToast('未识别到表格内容（支持制表符/逗号分隔）', 2500); return; }
      importFromGrid(tbl, grid, raw);
      rerenderTables(container);
      schedulePreview(container);
    });
  });
  container.querySelectorAll('button.tbl-paste-clear').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl') || '';
      if (!tableData[tbl]) return;
      tableData[tbl] = tableData[tbl].map((row) => row.map(() => ''));
      pasteBuffers[tbl] = '';
      delete tableMerges[tbl];
      const shape = template?.shapes.find((x) => x.id === tbl);
      if (shape) recomputeTableFit(shape);
      expandedTables.add(tbl);
      showToast('已清空全部单元格', 1500);
      rerenderTables(container);
      schedulePreview(container);
    });
  });
  container.querySelectorAll('select.tbl-merge-dir').forEach((sel) => {
    const s = sel as HTMLSelectElement;
    const tbl = s.getAttribute('data-tbl') || '';
    s.value = tableMergeDir[tbl] || 'auto';
    s.addEventListener('change', () => { tableMergeDir[tbl] = s.value as 'vertical' | 'horizontal' | 'auto'; });
  });
  container.querySelectorAll('button.tbl-merge-do').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl') || '';
      const shape = template?.shapes.find((x) => x.id === tbl);
      if (!shape) return;
      const dir = tableMergeDir[tbl] || 'auto';
      const cells = buildFitCellsFor(shape, tableData[tbl] || [[]], tableMerges[tbl]);
      const merged = mergeSameTextCells(cells, dir);
      const mergedCount = merged.filter((c) => c.rowspan > 1 || c.colspan > 1).length;
      if (!mergedCount) { showToast('未发现可合并的相同文字（需相邻、trim 后完全一致、连续 ≥2 格）', 2800); return; }
      lastTableMerges[tbl] = tableMerges[tbl] ? tableMerges[tbl].map((c) => ({ ...c })) : [];
      tableMerges[tbl] = merged;
      recomputeTableFit(shape);
      showToast('已合并 ' + mergedCount + ' 处相同文字 ✓', 2000);
      rerenderTables(container);
      schedulePreview(container);
    });
  });
  container.querySelectorAll('button.tbl-merge-undo').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl') || '';
      if (!(tbl in lastTableMerges)) { showToast('没有可撤销的合并', 1600); return; }
      if (lastTableMerges[tbl].length) tableMerges[tbl] = lastTableMerges[tbl];
      else delete tableMerges[tbl];
      delete lastTableMerges[tbl];
      const shape = template?.shapes.find((x) => x.id === tbl);
      if (shape) recomputeTableFit(shape);
      showToast('已撤销最近一次合并', 1500);
      rerenderTables(container);
      schedulePreview(container);
    });
  });
  // —— AI 生成表格数据：按表头+要求生成示例数据（预览后应用/放弃）——
  const aiGen = async (tbl: string) => {
    const shape = template?.shapes.find((x) => x.id === tbl);
    const req = container.querySelector(`textarea.tbl-ai-req[data-tbl="${tbl}"]`) as HTMLTextAreaElement | null;
    const reqText = (req?.value || '').trim();
    const cfg = await Api.getConfig().catch(() => null);
    if (!cfg || !cfg.text?.apiKey) {
      const choice = await showModal({
        title: '文本 AI 未配置',
        message: '尚未配置文本生成服务（API Key），无法自动生成表格数据。请先前往配置。',
        buttons: [
          { id: 'ok', label: '知道了', kind: 'primary' },
          { id: 'config', label: '前往配置', kind: 'danger' }
        ]
      });
      if (choice === 'config') location.hash = '#settings';
      return;
    }
    const btn = container.querySelector(`button.tbl-ai-do[data-tbl="${tbl}"]`) as HTMLButtonElement | null;
    if (btn) btn.disabled = true;
    const prog = btn ? showProgress(btn, '正在生成表格数据…') : null;
    try {
      const data = tableData[tbl] || defaultTableData(shape!);
      const headers = (data[0] || []).map((h, i) => `列${i + 1}${h ? '（' + h + '）' : ''}`).join('、') || '（无表头）';
      const { text } = await Api.generateText(
        '你是数据表格助手。根据表头与要求生成表格数据，只输出一个 JSON 二维字符串数组（如 [["a","b"],["1","2"]]），不要任何解释、文字或代码块标记，输出数组不要包含表头行。',
        `表头：${headers}\n当前已有 ${data.length} 行。\n要求：${reqText || '生成符合表头结构的示例数据'}`
      );
      const rows = parseTableAiJson(text);
      if (!rows.length) throw new Error('AI 返回的不是有效二维数组');
      aiResult[tbl] = rows;
      prog?.done();
      if (btn) btn.disabled = false;
      expandedTables.add(tbl);
      rerenderTables(container);
      const aiTab = container.querySelector(`button.tbl-tab[data-tbl="${tbl}"][data-mode="ai"]`) as HTMLButtonElement | null;
      aiTab?.click();
    } catch (e) {
      prog?.done();
      if (btn) btn.disabled = false;
      showToast('AI 生成失败：' + (e as Error).message, 4000);
    }
  };
  container.querySelectorAll('button.tbl-ai-do').forEach((btn) => {
    btn.addEventListener('click', () => { void aiGen((btn as HTMLButtonElement).getAttribute('data-tbl') || ''); });
  });
  container.querySelectorAll('button.tbl-ai-apply').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tbl = (btn as HTMLButtonElement).getAttribute('data-tbl') || '';
      const header = (tableData[tbl] || [])[0] || [];
      const aiRows = aiResult[tbl] || [];
      tableData[tbl] = [header, ...aiRows].slice(0, MAX_TABLE_ROWS).map((row) => row.slice(0, MAX_TABLE_COLS));
      delete aiResult[tbl];
      delete tableMerges[tbl];
      const shape = template?.shapes.find((x) => x.id === tbl);
      if (shape) recomputeTableFit(shape);
      expandedTables.add(tbl);
      showToast(`已应用 AI 数据（${aiRows.length} 行）✓`, 1800);
      rerenderTables(container);
      schedulePreview(container);
    });
  });
  container.querySelectorAll('button.tbl-ai-discard').forEach((btn) => {
    btn.addEventListener('click', () => {
      delete aiResult[(btn as HTMLButtonElement).getAttribute('data-tbl') || ''];
      rerenderTables(container);
    });
  });
}

// —— 预览区（实时预览/放大）+ 底部「生成并插入 PPT」 ——
function bindPreviewAndWrite(container: HTMLElement): void {
  const zoomBtn = container.querySelector('#wb-preview-zoom') as HTMLButtonElement | null;
  const inner = container.querySelector('#wb-preview-inner') as HTMLElement | null;
  const previewNow = () => { if (template) showPreviewModal(renderPreview(template, images, texts, vars, tableData, tableFits, tableMerges)); };
  zoomBtn?.addEventListener('click', previewNow);
  inner?.addEventListener('click', previewNow);
  const toggleBtn = container.querySelector('#wb-preview-toggle') as HTMLButtonElement | null;
  const body = container.querySelector('#wb-preview-body') as HTMLElement | null;
  toggleBtn?.addEventListener('click', () => {
    previewCollapsed = !previewCollapsed;
    if (body) body.style.display = previewCollapsed ? 'none' : 'block';
    if (toggleBtn) toggleBtn.textContent = previewCollapsed ? '展开 ▴' : '收起 ▾';
  });
  const writeBtn = container.querySelector('#write-ppt') as HTMLButtonElement;
  writeBtn.addEventListener('click', async () => {
    try {
      if (!backgroundState.followDocument && !backgroundState.customImageDataUrl) {
        showToast('请先选择一张背景图片，或勾选「跟随文档」。', 2500);
        return;
      }
      // 校验：图片位必须全部选图
      const missingImgs = (template?.shapes || []).filter((m) => m.role === 'ai_image' && !(images[m.id] || '').trim()).map((m) => m.name || m.id);
      if (missingImgs.length) { showToast('还有图片位未选择图片：' + missingImgs.join('、'), 2500); return; }
      // 校验：文本位（除序号/日期/图注/公式）必须非空
      const special = (m: TemplateShape) => m.semanticRole === 'seq' || m.semanticRole === 'date' || m.semanticRole === 'caption' || m.semanticRole === 'formula';
      const emptyTexts = (template?.shapes || []).filter((m) =>
        (m.role === 'ai_text' || m.role === 'manual_var') && !special(m) && !(m.role === 'manual_var' ? vars[m.id] || '' : texts[m.id] || '').trim()
      ).map((m) => m.name || m.id);
      if (emptyTexts.length && await showModal({
        title: '有文本位为空',
        message: `以下文本位还没有内容，写入后会显示为空，确定继续？\n` + emptyTexts.join('\n'),
        buttons: [
          { id: 'cancel', label: '取消' },
          { id: 'continue', label: '仍然写入', kind: 'primary' }
        ]
      }) !== 'continue') return;
      // 质量检查门禁：存在 error 级问题且用户选择继续才写入
      if (lastQuality && lastQuality.error > 0) {
        const issues = lastQuality.issues.filter((w) => w.level === 'error').slice(0, 4).map((w) => '· ' + w.label + '：' + w.message).join('\n');
        if (await showModal({
          title: '生成质量检查发现异常',
          message: '以下 ' + lastQuality.error + ` 项检查未通过，生成后页面可能存在排版问题：\n` + issues + (lastQuality.error > 4 ? '\n…' : '') + '\n\n仍然写入？',
          buttons: [
            { id: 'cancel', label: '取消' },
            { id: 'continue', label: '仍然写入', kind: 'primary' }
          ]
        }) !== 'continue') return;
      }
      // 表格：写入前重算 fit，保证提交数据与预览一致
      const tableShapes = (template?.shapes || []).filter((m) => m.role === 'table' && m.table);
      for (const m of tableShapes) recomputeTableFit(m);
      // 公式位：优先使用 AI 生成的 LaTeX，否则本地规范化
      const textsOut = { ...texts };
      for (const m of template?.shapes || []) {
        if (m.role === 'ai_text' && m.semanticRole === 'formula') {
          const h = texts[m.id] || '';
          if (!h.trim()) continue;
          textsOut[m.id] = (formulaLatex[m.id] || '').trim() || localTextToLatex(h);
        }
      }
      const payload: PendingSlide = {
        templateId: sessionStorage.getItem('templateId') || '',
        folder: sessionStorage.getItem('templateFolder') || '',
        template: template ? cloneTemplateWithWizardBackground(template, backgroundState) : undefined,
        images,
        texts: textsOut,
        vars,
        tableData,
        tables: Object.keys(tableFits).length ? { ...tableFits } : undefined,
        clearSrcRectFor: croppedImages.size ? Array.from(croppedImages) : undefined
      };
      let insertAtStart = false;
      for (;;) {
        writeBtn.disabled = true;
        const prog = showProgress(writeBtn, '正在读取模板…');
        try {
          await writePendingSlide(payload, { insertAtStart, onStage: (stage) => prog.setText(`正在${stage}…`) });
          prog.done();
          writeBtn.disabled = false;
          showToast('已写入当前 PPT', 1500);
          setTimeout(() => { location.hash = '#library'; }, 800);
          return;
        } catch (e) {
          prog.done();
          writeBtn.disabled = false;
          const choice = await handleWriteError(e as Error);
          if (choice === 'retry') continue;
          if (choice === 'start') { insertAtStart = true; continue; }
          return;
        }
      }
    } catch (e) {
      showToast(`写入失败：${(e as Error).message}`, 4000);
    }
  });
}

// ================= 图片搜索 / 下载 / 裁剪（与旧向导一致） =================
async function doSearch(container: HTMLElement, shapeId: string): Promise<void> {
  const st = imgStates[shapeId] || defaultImgState();
  const box = container.querySelector(`.img-results[data-id="${shapeId}"]`) as HTMLElement;
  const nav = container.querySelector(`.img-nav[data-id="${shapeId}"]`) as HTMLElement;
  box.innerHTML = '<p>搜索中…</p>';
  nav.innerHTML = '';
  const searchBtn = container.querySelector(`button.img-search[data-id="${shapeId}"]`) as HTMLButtonElement;
  searchBtn.disabled = true;
  const searchProgress = showProgress(searchBtn, '正在搜索图片…');
  try {
    // 一次拉全量（百度 60 张 / Bing 约 35 张），翻页在本地按 pageSize 切片
    const res = await Api.searchImages(st.q, 60, 1, imgProvider);
    searchProgress.done();
    searchBtn.disabled = false;
    st.images = res.images;
    st.page = 1;
    st.providerError = res.error ? (res.error.message || '搜索来源暂时不可用') : '';
    renderResults(container, shapeId);
    // 图片诠释位：把搜索关键词自动 copy 到诠释输入框（仅当该字段为空；本地选图不填）
    const q = (st.q || '').trim();
    if (q && template) {
      let captionChanged = false;
      for (const cs of template.shapes) {
        if (cs.role === 'ai_text' && cs.semanticRole === 'caption' && !(texts[cs.id] || '').trim()) {
          texts[cs.id] = q;
          const cin = container.querySelector(`input.gen-in[data-gen="${cs.id}"]`) as HTMLInputElement | null;
          if (cin) cin.value = q;
          captionChanged = true;
        }
      }
      if (captionChanged) schedulePreview(container);
    }
  } catch (e) {
    searchProgress.done();
    searchBtn.disabled = false;
    box.innerHTML = `<p class="error">搜索失败：${escapeHtml((e as Error).message)}</p>`;
  }
}

function renderResults(container: HTMLElement, shapeId: string): void {
  const st = imgStates[shapeId] || defaultImgState();
  const box = container.querySelector(`.img-results[data-id="${shapeId}"]`) as HTMLElement;
  const nav = container.querySelector(`.img-nav[data-id="${shapeId}"]`) as HTMLElement;
  if (!st.images.length) {
    box.innerHTML = '<p class="error">没有搜到图片，换个关键词试试。</p>';
    nav.innerHTML = '';
    return;
  }
  const pageSize = st.pageSize;
  const totalPages = Math.max(1, Math.ceil(st.images.length / pageSize));
  if (st.page > totalPages) st.page = totalPages;
  const slice = st.images.slice((st.page - 1) * pageSize, st.page * pageSize);
  const errTip = st.providerError
    ? `<p class="error" style="margin:0 0 6px">⚠ ${escapeHtml(st.providerError)}</p>`
    : '';
  // 只显示图片缩略图（不显示每张的文字介绍）；右上角「＋」可预览整张图（不下载）
  box.innerHTML = errTip + slice.map((img) => `
    <div class="card" style="padding:4px;position:relative">
      <button type="button" class="img-zoom" data-full="${escapeAttr(img.imageUrl)}" title="预览整张图">＋</button>
      <img class="thumb" src="${escapeAttr(thumbnailUrlOf(img))}" data-full="${escapeAttr(img.imageUrl)}" style="cursor:pointer${img.imageUrl === st.selected ? ';outline:3px solid #1f3864' : ''}" />
    </div>`).join('');
  nav.innerHTML = `
    <button class="secondary" id="img-prev-${shapeId}" ${st.page <= 1 ? 'disabled' : ''}>上一页</button>
    <span class="page-num">第 ${st.page} / ${totalPages} 页</span>
    <button class="secondary" id="img-next-${shapeId}" ${st.page >= totalPages ? 'disabled' : ''}>下一页</button>`;
  nav.querySelector(`#img-prev-${shapeId}`)!.addEventListener('click', () => { st.page -= 1; renderResults(container, shapeId); });
  nav.querySelector(`#img-next-${shapeId}`)!.addEventListener('click', () => { st.page += 1; renderResults(container, shapeId); });
  box.querySelectorAll('img[data-full]').forEach((el) => {
    el.addEventListener('click', () => downloadAndSelect(container, el as HTMLImageElement, shapeId));
  });
  // 右上角「＋」：预览整张原图（只预览，不下载）
  box.querySelectorAll('button.img-zoom').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const url = btn.getAttribute('data-full') || '';
      if (url) showPreviewModal(`<div style="text-align:center"><img src="${url}" style="max-width:100%;max-height:70vh;border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.25)" /></div>`);
    });
  });
}

const downloadLocks = new Set<string>(); // 每个图片位独立防重复点击

// 点击缩略图：下载原图（走安全下载）→ 设为本位图片（比例不符自动弹裁剪框）
async function downloadAndSelect(container: HTMLElement, el: HTMLImageElement, shapeId: string): Promise<void> {
  const url = el.getAttribute('data-full') || '';
  if (!url) return;
  const st = imgStates[shapeId] || defaultImgState();
  if ((st.selected === url && images[shapeId]) || downloadLocks.has(shapeId)) return;
  downloadLocks.add(shapeId);
  const box = container.querySelector(`.img-results[data-id="${shapeId}"]`) as HTMLElement;
  const card = el.closest('.card');
  const old = box.querySelector('.progress-wrap');
  if (old) old.remove();
  const wrap = document.createElement('div');
  wrap.className = 'progress-wrap';
  wrap.innerHTML = '<div class="progress-track"><div class="progress-bar indeterminate"></div></div><div class="progress-text">正在下载…</div>';
  card?.appendChild(wrap);
  const bar = wrap.querySelector('.progress-bar') as HTMLElement;
  const text = wrap.querySelector('.progress-text') as HTMLElement;
  try {
    const { taskId } = await Api.downloadImage(url, imgProvider);
    let status: { done: boolean; error?: string; received: number; total: number | null; dataUrl: string } = { done: false, received: 0, total: null, dataUrl: '' };
    for (;;) {
      await sleep(200);
      status = await Api.getDownloadStatus(taskId);
      if (status.total) {
        const pct = Math.min(100, Math.round((status.received / Math.max(status.total, 1)) * 100));
        bar.classList.remove('indeterminate');
        bar.style.width = `${pct}%`;
        text.textContent = `正在下载 ${pct}%`;
      } else {
        text.textContent = '正在下载…';
      }
      if (status.done || status.error) break;
    }
    if (status.error) throw new Error(status.error);
    images[shapeId] = status.dataUrl;
    st.selected = url;
    try {
      const shape = template?.shapes.find((x) => x.id === shapeId && x.role === 'ai_image');
      if (shape?.bounds?.width && shape.bounds.height) {
        const frameRatio = shape.bounds.width / shape.bounds.height;
        const wPx = Math.round(shape.bounds.width * 96);
        const hPx = Math.round(shape.bounds.height * 96);
        const crop = await openImageCropEditor({ imageDataUrl: status.dataUrl, frameRatio, frameSizeLabel: `模板图片位：${wPx} × ${hPx} px` });
        if (!crop.canceled && crop.dataUrl) { images[shapeId] = crop.dataUrl; croppedImages.add(shapeId); }
      }
    } catch { /* 裁剪失败不影响选图 */ }
    box.querySelectorAll('img[data-full]').forEach((h) => { (h as HTMLElement).style.outline = ''; });
    el.style.outline = '3px solid #1f3864';
    renderChosen(container, shapeId);
    schedulePreview(container);
    text.textContent = '下载完成 ✓';
    setTimeout(() => wrap.remove(), 900);
  } catch (e) {
    wrap.remove();
    alert(`图片下载失败：${(e as Error).message}`);
  } finally {
    downloadLocks.delete(shapeId);
  }
}

// 本地图片/拖拽：dataURL 直接设为图片位（比例不符自动弹裁剪框）
async function setImage(container: HTMLElement, shapeId: string, dataUrl: string, label: string): Promise<void> {
  images[shapeId] = dataUrl;
  try {
    const shape = template?.shapes.find((x) => x.id === shapeId && x.role === 'ai_image');
    if (shape?.bounds?.width && shape.bounds.height) {
      const frameRatio = shape.bounds.width / shape.bounds.height;
      const wPx = Math.round(shape.bounds.width * 96);
      const hPx = Math.round(shape.bounds.height * 96);
      const crop = await openImageCropEditor({ imageDataUrl: dataUrl, frameRatio, frameSizeLabel: `模板图片位：${wPx} × ${hPx} px` });
      if (!crop.canceled && crop.dataUrl) { images[shapeId] = crop.dataUrl; croppedImages.add(shapeId); }
    }
  } catch { /* 裁剪失败不影响选图 */ }
  const st = (imgStates[shapeId] = imgStates[shapeId] || defaultImgState());
  st.selected = 'local:' + label;
  renderChosen(container, shapeId);
  schedulePreview(container);
  showToast('已选择' + label + ' ✓', 1500);
}

// 重新按图片位比例裁剪当前已选图片
async function cropImage(container: HTMLElement, shapeId: string): Promise<void> {
  const dataUrl = images[shapeId];
  if (!dataUrl) return;
  const shape = template?.shapes.find((s) => s.id === shapeId && s.role === 'ai_image');
  if (!shape?.bounds?.width || !shape.bounds.height) { showToast('该图片位没有尺寸信息，无法裁剪', 2500); return; }
  try {
    const crop = await openImageCropEditor({
      imageDataUrl: dataUrl,
      frameRatio: shape.bounds.width / shape.bounds.height,
      frameSizeLabel: '模板图片位：' + Math.round(shape.bounds.width * 96) + ' × ' + Math.round(shape.bounds.height * 96) + ' px'
    });
    if (!crop.canceled && crop.dataUrl) {
      images[shapeId] = crop.dataUrl;
      croppedImages.add(shapeId);
      renderChosen(container, shapeId);
      schedulePreview(container);
      showToast('已按图片位比例裁剪 ✓', 1500);
    }
  } catch {
    showToast('裁剪失败', 2000);
  }
}

// 已选图片展示（缩略图 + 裁剪/清除按钮）
function renderChosen(container: HTMLElement, shapeId: string): void {
  const el = container.querySelector(`.img-chosen[data-id="${shapeId}"]`) as HTMLElement | null;
  if (!el) return;
  const dataUrl = images[shapeId];
  if (!dataUrl) { el.style.display = 'none'; return; }
  el.style.display = '';
  el.innerHTML = '<img src="' + dataUrl + '" style="max-width:120px;max-height:80px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />'
    + ' <button class="secondary img-crop" data-id="' + shapeId + '" title="按图片位比例重新裁剪">✂ 裁剪</button>'
    + ' <button class="secondary img-clear" data-id="' + shapeId + '">✕ 清除</button>';
  el.querySelector(`button.img-crop[data-id="${shapeId}"]`)?.addEventListener('click', () => { void cropImage(container, shapeId); });
  el.querySelector(`button.img-clear[data-id="${shapeId}"]`)?.addEventListener('click', () => {
    delete images[shapeId];
    const st = imgStates[shapeId];
    if (st) st.selected = '';
    renderChosen(container, shapeId);
    schedulePreview(container);
  });
}
