// 「生成质量检查」：把模板语义约束（maxChars / maxLines / minChars / preferredLength / required）
// 从"控制 AI 生成"升级为"检查最终排版"，并补充布局类检查：
//   ① 文本溢出  ② 行数·字数限制  ③ 图片留白（分辨率不足）  ④ 图片裁剪异常  ⑤ 表格超出  ⑥ 元素重叠
// 纯前端计算：canvas 测量文本宽度 + 几何计算；不依赖后端。
import type { TemplateDoc, TemplateShape } from '../api.js';
import type { FitResult } from '../api.js'; // 表格适配结果（引擎输出，可选）
import { semanticRoleLabel } from './semantic.js';

export type QualityLevel = 'pass' | 'warn' | 'error';

export interface QualityIssue {
  category: string; // 检查类别（对应 QUALITY_CATEGORIES）
  level: QualityLevel;
  label: string;    // 涉及的元素名（如「标题」「图片位 2」）
  message: string;
}

export interface QualityReport {
  issues: QualityIssue[];
  pass: number;  // 通过类别数
  warn: number;  // 存在警告的类别数
  error: number; // 存在异常的类别数
  hasError: boolean;
  hasWarn: boolean;
}

export const QUALITY_CATEGORIES = ['文本溢出', '行数·字数限制', '图片留白', '图片裁剪', '表格超出', '元素重叠'];

// —— 文本宽度测量（canvas；不可用时回退近似值）——
let _ctx: CanvasRenderingContext2D | null | undefined;
function measurer(): CanvasRenderingContext2D | null {
  if (_ctx === undefined) {
    try { const c = document.createElement('canvas'); _ctx = c.getContext('2d'); }
    catch { _ctx = null; }
  }
  return _ctx;
}
function charWidth(ch: string, px: number, font: string): number {
  const c = measurer();
  if (c) { c.font = px + 'px ' + font; return c.measureText(ch).width; }
  // 回退：CJK 全角 ≈ 1em，ASCII ≈ 0.55em
  return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/.test(ch) ? px : px * 0.55;
}
// 估算 text 在 maxW px 宽度下自动换行后的总行数（\n 强制换行）
export function wrapEstimate(text: string, maxW: number, px: number, font: string): number {
  let total = 0;
  for (const ln of text.split('\n')) {
    if (!ln) { total += 1; continue; }
    let lines = 1, w = 0;
    for (const ch of ln) {
      const cw = charWidth(ch, px, font);
      if (w > 0 && w + cw > maxW) { lines += 1; w = cw; } else w += cw;
    }
    total += lines;
  }
  return Math.max(1, total);
}

const INCH_PX = 96;
const PT_PX = 96 / 72;

// —— 图片尺寸缓存（同一 dataURL 只解析一次）——
const imgSizeCache = new Map<string, { w: number; h: number }>();
export function clearImageSizeCache(): void { imgSizeCache.clear(); }
function loadImageSize(dataUrl: string): Promise<{ w: number; h: number } | null> {
  const hit = imgSizeCache.get(dataUrl);
  if (hit) return Promise.resolve(hit);
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => { const s = { w: img.naturalWidth || 0, h: img.naturalHeight || 0 }; imgSizeCache.set(dataUrl, s); resolve(s); };
    img.onerror = () => resolve(null);
    img.src = dataUrl;
  });
}

function fmtRatio(r: number): string {
  if (!isFinite(r) || r <= 0) return '?';
  return r >= 1 ? r.toFixed(2) + ':1' : '1:' + (1 / r).toFixed(2);
}

// 元素显示名（与工作台内标签一致）
export function shapeLabel(s: TemplateShape, i: number): string {
  const name = s.name ? '「' + s.name + '」' : '';
  if (s.role === 'ai_image') return '图片位 ' + (i + 1) + name;
  if (s.role === 'table') return '表格 ' + (i + 1) + name;
  if (s.role === 'manual_var') return s.varName ? '变量「' + s.varName + '」' : '变量 ' + (i + 1);
  const role = semanticRoleLabel(s.semanticRole);
  return (role || '文本位 ' + (i + 1)) + name;
}

// 表格默认数据（首行表头，其余模板原文）——与预览/后端一致
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

// ① 文本溢出 + ② 行数·字数限制
function checkTextShape(s: TemplateShape, text: string, label: string, issues: QualityIssue[]): void {
  const ts = s.textStyle || {};
  const sizePx = (ts.size || 18) * PT_PX;
  const font = ts.font || 'Microsoft YaHei';
  const b = s.bounds || {};
  const mL = ts.margin?.left ?? 0.1, mR = ts.margin?.right ?? 0.1;
  const mT = ts.margin?.top ?? 0.05, mB = ts.margin?.bottom ?? 0.05;
  const usableW = Math.max(40, (b.width - mL - mR) * INCH_PX);
  const usableH = Math.max(20, (b.height - mT - mB) * INCH_PX);
  const lineH = sizePx * 1.18;
  const availLines = Math.max(1, Math.floor(usableH / lineH));
  const trimmed = (text || '').trim();
  if (!trimmed) {
    return; // 空内容不再报「必填」警告（避免频繁打扰）
  }
  const estLines = wrapEstimate(text, usableW, sizePx, font);
  const autofit = !!ts.autoFit && ts.autoFit !== 'none';
  if (estLines > availLines) {
    const mildAutoFitOverflow = autofit && estLines <= Math.ceil(availLines * 1.35);
    if (!mildAutoFitOverflow) {
      const msg = '约需 ' + estLines + ' 行，框内仅能容纳约 ' + availLines + ' 行，会溢出';
      issues.push({ category: '文本溢出', level: autofit ? 'warn' : 'error', label, message: autofit ? msg + '（模板已设置自动缩小文字）' : msg });
    }
  }
  if (s.maxLines && s.maxLines > 0 && estLines > s.maxLines) {
    issues.push({ category: '行数·字数限制', level: 'warn', label, message: '当前约 ' + estLines + ' 行，模板限定最多 ' + s.maxLines + ' 行' });
  }
  if (s.maxChars && s.maxChars > 0 && text.length > s.maxChars) {
    issues.push({ category: '行数·字数限制', level: 'warn', label, message: '当前 ' + text.length + ' 字，模板限定最多 ' + s.maxChars + ' 字' });
  }
  if (s.minChars && s.minChars > 0 && trimmed.length < s.minChars) {
    issues.push({ category: '行数·字数限制', level: 'warn', label, message: '当前 ' + trimmed.length + ' 字，模板要求至少 ' + s.minChars + ' 字' });
  }
}

// ③ 图片留白（分辨率不足 → 放大模糊/留白） + ④ 图片裁剪异常
async function checkImageShape(
  s: TemplateShape, dataUrl: string, label: string, cropped: boolean, issues: QualityIssue[]
): Promise<void> {
  if (!dataUrl || !dataUrl.trim()) {
    issues.push({ category: '图片留白', level: 'warn', label, message: '未选择图片' });
    return;
  }
  const size = await loadImageSize(dataUrl);
  if (!size || !size.w || !size.h) {
    issues.push({ category: '图片裁剪', level: 'warn', label, message: '无法读取图片尺寸（文件可能损坏）' });
    return;
  }
  const b = s.bounds || {};
  const shapeWpx = (b.width || 0) * INCH_PX;
  const shapeHpx = (b.height || 0) * INCH_PX;
  if (!shapeWpx || !shapeHpx) return;
  const shapeRatio = b.width / b.height;
  const imgRatio = size.w / size.h;
  const ratioDiff = Math.abs(imgRatio - shapeRatio) / shapeRatio;
  // 图片留白：分辨率不足以铺满图片位
  if (size.w < shapeWpx * 0.6 || size.h < shapeHpx * 0.6) {
    issues.push({
      category: '图片留白', level: 'warn', label,
      message: '分辨率偏低（' + size.w + '×' + size.h + '），铺满图片位（约 ' + Math.round(shapeWpx) + '×' + Math.round(shapeHpx) + ' px）会模糊、显得留白'
    });
  }
  // 图片裁剪：比例匹配性 / 裁剪结果异常
  if (cropped) {
    if (ratioDiff > 0.03) {
      issues.push({ category: '图片裁剪', level: 'error', label, message: '裁剪结果比例（' + fmtRatio(imgRatio) + '）与图片位（' + fmtRatio(shapeRatio) + '）不一致，裁剪可能异常' });
    }
  } else if (ratioDiff > 0.5) {
    issues.push({
      category: '图片裁剪', level: 'warn', label,
      message: '图片比例（' + fmtRatio(imgRatio) + '）与图片位（' + fmtRatio(shapeRatio) + '）差异过大，生成时会被拉伸变形、裁切掉大部分内容；建议换图或裁剪'
    });
  } else if (ratioDiff > 0.08) {
    issues.push({
      category: '图片裁剪', level: 'warn', label,
      message: '图片比例（' + fmtRatio(imgRatio) + '）与图片位（' + fmtRatio(shapeRatio) + '）不一致，生成时会拉伸变形（预览按裁剪显示）；建议裁剪到图片位比例'
    });
  }
}

// ⑤ 表格超出：列宽/行高算法与后端 slideBuilder 一致，逐格估算文字是否溢出
function checkTableShape(s: TemplateShape, data: string[][], label: string, issues: QualityIssue[], fit?: FitResult): void {
  // 有 fit 结果（向导已接入排版引擎）：以引擎判定为准（fit 已精确算列宽/行高/缩字/溢出），避免与旧估算重复报
  if (fit) {
    if (fit.overflow) {
      const b = s.bounds || {};
      const extend = Math.max(0, fit.rowHeights.reduce((a, x) => a + x, 0) - (b.height || 0));
      issues.push({
        category: '表格超出', level: 'warn', label,
        message: '表格内容超出：向下延伸 ' + extend.toFixed(2) + ' 英寸（生成时表体会自动撑高）'
      });
    }
    return;
  }
  const b = s.bounds || {};
  const w = b.width, h = b.height;
  if (!w || !h) return;
  const rows = data.length;
  const cols = Math.max(1, ...data.map((r) => (Array.isArray(r) ? r.length : 0)));
  const t = s.table;
  const savedCols = (t && Array.isArray(t.colWidths)) ? t.colWidths : [];
  const savedRows = (t && Array.isArray(t.rowHeights)) ? t.rowHeights : [];
  const sumCol = savedCols.reduce((a, x) => a + (Number(x) || 0), 0);
  const sumRow = savedRows.reduce((a, x) => a + (Number(x) || 0), 0);
  const extraC = Math.max(cols - savedCols.length, 0);
  const oldColBudget = extraC > 0 ? w * 0.75 : w;
  const savedColSum = savedCols.slice(0, cols).reduce((a, x) => a + (Number(x) || 0), 0);
  const colW: number[] = [];
  for (let c = 0; c < cols; c++) {
    if (c < savedCols.length) {
      const sc = Number(savedCols[c]) || 0;
      colW.push(savedColSum > 0 ? (sc / savedColSum) * oldColBudget : oldColBudget / Math.min(cols, savedCols.length));
    } else {
      colW.push(extraC > 0 ? (w - oldColBudget) / extraC : 0);
    }
  }
  if (sumCol <= 0) for (let c = 0; c < cols; c++) colW[c] = w / cols;
  const extraR = Math.max(rows - savedRows.length, 0);
  const oldRowBudget = Math.max(h - extraR * 0.3, 0);
  const savedRowSum = savedRows.slice(0, rows).reduce((a, x) => a + (Number(x) || 0), 0);
  const rowH: number[] = [];
  for (let r = 0; r < rows; r++) {
    if (r < savedRows.length) {
      const sr = Number(savedRows[r]) || 0;
      rowH.push(savedRowSum > 0 ? (sr / savedRowSum) * oldRowBudget : oldRowBudget / Math.min(rows, savedRows.length));
    } else {
      rowH.push(0.3);
    }
  }
  if (sumRow <= 0) for (let r = 0; r < rows; r++) rowH[r] = h / rows;
  const minCol = Math.min(...colW);
  if (minCol < 0.45) {
    issues.push({ category: '表格超出', level: 'warn', label, message: '共 ' + cols + ' 列，最窄列约 ' + minCol.toFixed(2) + ' 英寸，内容会拥挤' });
  }
  const minRow = Math.min(...rowH);
  if (rows > 10 && minRow < 0.28) {
    issues.push({ category: '表格超出', level: 'warn', label, message: '共 ' + rows + ' 行，行高过密（约 ' + minRow.toFixed(2) + ' 英寸/行）' });
  }
  // 单元格文字溢出估算
  const cellAt = (r: number, c: number) => (t ? (t.cells || []).find((x) => x.row === r && x.col === c) : undefined);
  let overflow = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cellText = data[r] && data[r][c] ? data[r][c] : '';
      if (!cellText.trim()) continue;
      const cell = cellAt(r, c);
      const cts = (cell && cell.textStyle) || {};
      const fontPx = (cts.size || 11) * PT_PX;
      const fontFace = cts.eaFont || cts.font || 'Microsoft YaHei';
      const cw = Math.max(20, colW[c] * INCH_PX - 10);
      const ch = Math.max(12, rowH[r] * INCH_PX - 6);
      const lines = wrapEstimate(cellText, cw, fontPx, fontFace);
      if (lines * fontPx * 1.3 > ch) overflow++;
    }
  }
  if (overflow > 0) {
    issues.push({ category: '表格超出', level: 'warn', label, message: '有 ' + overflow + ' 个单元格文字超出，行会被撑高（可能超出页面）' });
  }
}

// ⑥ 元素重叠：内容元素两两检查交集面积（固定装饰/背景不在检查范围）
function checkOverlaps(
  template: TemplateDoc, images: Record<string, string>, texts: Record<string, string>,
  vars: Record<string, string>, tableData: Record<string, string[][]>, issues: QualityIssue[]
): void {
  const content = (template.shapes || []).map((s, i) => ({ s, i })).filter(({ s }) => {
    if (s.role === 'ai_text') return !!((texts[s.id] || '')).trim();
    if (s.role === 'manual_var') return !!((vars[s.id] || '')).trim();
    if (s.role === 'ai_image') return !!((images[s.id] || '')).trim();
    if (s.role === 'table' && s.table) return true;
    if (s.role === 'fixed' && (s.type === 'text' || s.type === 'other') && typeof s.content === 'string' && s.content.trim()) return true;
    return false;
  });
  for (let i = 0; i < content.length; i++) {
    for (let j = i + 1; j < content.length; j++) {
      const a = content[i].s, b2 = content[j].s;
      const ra = a.bounds, rb = b2.bounds;
      if (!ra || !rb || !ra.width || !ra.height || !rb.width || !rb.height) continue;
      const ix = Math.min(ra.left + ra.width, rb.left + rb.width) - Math.max(ra.left, rb.left);
      const iy = Math.min(ra.top + ra.height, rb.top + rb.height) - Math.max(ra.top, rb.top);
      if (ix <= 0.02 || iy <= 0.02) continue;
      const inter = ix * iy;
      const minArea = Math.min(ra.width * ra.height, rb.width * rb.height);
      if (minArea > 0 && inter / minArea >= 0.12) {
        issues.push({
          category: '元素重叠', level: 'warn',
          label: shapeLabel(a, content[i].i) + ' × ' + shapeLabel(b2, content[j].i),
          message: '区域重叠 ' + ix.toFixed(2) + '×' + iy.toFixed(2) + ' 英寸'
        });
      }
    }
  }
}

export async function runQualityChecks(
  template: TemplateDoc,
  images: Record<string, string>,
  texts: Record<string, string>,
  vars: Record<string, string>,
  tableData: Record<string, string[][]>,
  croppedImages?: Set<string>,
  tableFits?: Record<string, FitResult>
): Promise<QualityReport> {
  const issues: QualityIssue[] = [];
  const imgOf = (id: string): string => images[id] || '';
  // ① ② 文本
  template.shapes.forEach((s, i) => {
    if (s.role !== 'ai_text' && s.role !== 'manual_var') return;
    const text = s.role === 'ai_text' ? (texts[s.id] || '') : (vars[s.id] || '');
    checkTextShape(s, text, shapeLabel(s, i), issues);
  });
  // ③ ④ 图片
  const imageShapes = template.shapes.filter((s) => s.role === 'ai_image');
  for (let k = 0; k < imageShapes.length; k++) {
    const s = imageShapes[k];
    const idx = template.shapes.indexOf(s);
    await checkImageShape(s, imgOf(s.id), shapeLabel(s, idx), !!(croppedImages && croppedImages.has(s.id)), issues);
  }
  // ⑤ 表格
  template.shapes.forEach((s, i) => {
    if (s.role !== 'table' || !s.table) return;
    const data = tableData[s.id] || defaultTableData(s);
    checkTableShape(s, data, shapeLabel(s, i), issues, tableFits?.[s.id]);
  });
  // ⑥ 重叠
  checkOverlaps(template, images, texts, vars, tableData, issues);
  // 汇总：每类别取最差级别
  let pass = 0, warn = 0, error = 0;
  for (const cat of QUALITY_CATEGORIES) {
    let worst: QualityLevel = 'pass';
    for (const x of issues) {
      if (x.category !== cat) continue;
      if (x.level === 'error') worst = 'error';
      else if (x.level === 'warn' && worst !== 'error') worst = 'warn';
    }
    if (worst === 'error') error++;
    else if (worst === 'warn') warn++;
    else pass++;
  }
  return { issues, pass, warn, error, hasError: error > 0, hasWarn: warn > 0 };
}
