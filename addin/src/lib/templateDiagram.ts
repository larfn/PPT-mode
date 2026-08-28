// 模板预览结构图生成器（接近真实页面布局的骨架预览）
//
// 设计原则（用户确认）：
//  - 整体按真实布局渲染：背景 + 元素按模板真实位置/字号/颜色/对齐排布（类似向导实时预览）
//  - 图片位：纯色填充 + 反色「图片」
//  - 正文/要点这类「需要表明范围」的长文本：虚线范围框 + 框内灰色占位文字（多行，按真实字号自动换行）
//  - 标题/序号/页码/标签等短元素：不画框，直接渲染占位字（角色名，真实字号/颜色/对齐）
//  - 固定元素：渲染真实内容（不画框）；表格：真实行列网格；线条：直接画线
// 输出与 PPT 页面同比例的 PNG dataURL（供保存为模板预览）。
import { TemplateDoc } from '../api.js';

// 语义角色 → 中文（与语义层 SEMANTIC_ROLES 对齐，精简为 8 类）
const ROLE_LABEL: Record<string, string> = {
  title: '主标题',
  subtitle: '副标题',
  body: '正文',
  seq: '序号',
  date: '日期',
  caption: '图片诠释',
  formula: '公式',
  other: '其他'
};

// 「需要表明范围」的长文本角色：画虚线范围框 + 多行占位文字（仅正文）
const RANGE_ROLES = new Set(['body']);

const PX_PER_INCH = 96;
const FONT_FAMILY = '"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif';

function setFont(ctx: CanvasRenderingContext2D, size: number, bold: boolean, italic = false): void {
  ctx.font = (italic ? 'italic ' : '') + (bold ? 'bold ' : '') + size + 'px ' + FONT_FAMILY;
}

// 真实字号 → px（pt×96/72），超框自动收敛
function fitFontSize(ctx: CanvasRenderingContext2D, text: string, ts: TemplateDoc['shapes'][number]['textStyle'], w: number, h: number, bold: boolean): number {
  const realPt = Number(ts && ts.size) > 0 ? Number(ts && ts.size) : 12;
  let size = Math.round(realPt * 96 / 72);
  size = Math.max(size, 10);
  if (h > 0) size = Math.min(size, h * 0.95);
  setFont(ctx, size, bold, !!(ts && ts.italic));
  const maxW = Math.max(w * 0.95, 10);
  while (size > 9 && ctx.measureText(text).width > maxW) {
    size -= 1;
    setFont(ctx, size, bold, !!(ts && ts.italic));
  }
  return size;
}

// 文字颜色：太浅（接近白底）时加深保证可见
function visibleColor(color: string | undefined, fallback: string): string {
  const c = color || fallback;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(c.trim());
  if (!m) return fallback;
  const n = parseInt(m[1], 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const lum = 0.299 * r + 0.587 * g + 0.114 * b;
  return lum < 150 ? c : fallback; // 太浅的模板文字色（如浅灰页脚）→ 深灰保证可见
}

// 按宽度把一段文字自动换行（返回行数组）
function wrapText(ctx: CanvasRenderingContext2D, text: string, maxW: number): string[] {
  const lines: string[] = [];
  let line = '';
  for (const ch of text) {
    const t = line + ch;
    if (ctx.measureText(t).width > maxW && line) {
      lines.push(line);
      line = ch;
    } else {
      line = t;
    }
    if (lines.length >= 40) break;
  }
  if (line) lines.push(line);
  return lines;
}

// 单行占位文字（短元素：标题/序号/页码等，不画框，真实字号/颜色/对齐）
function drawShortText(ctx: CanvasRenderingContext2D, text: string, x: number, y: number, w: number, h: number, ts: TemplateDoc['shapes'][number]['textStyle']): void {
  if (w < 8 || h < 4 || !text) return;
  const bold = !!(ts && ts.bold);
  const size = fitFontSize(ctx, text, ts, w, h, bold);
  setFont(ctx, size, bold, !!(ts && ts.italic));
  const align = (ts && ts.align) || 'left';
  const valign = (ts && ts.valign) || 'middle';
  let tx = x;
  if (align === 'center') tx = x + w / 2;
  else if (align === 'right') tx = x + w;
  let ty = y + h / 2;
  if (valign === 'top') ty = y + size * 0.72;
  else if (valign === 'bottom') ty = y + h - size * 0.25;
  ctx.fillStyle = visibleColor(ts && ts.color, '#444444');
  ctx.textAlign = align === 'center' ? 'center' : (align === 'right' ? 'right' : 'left');
  ctx.textBaseline = 'middle';
  // 超宽省略
  let shown = text;
  const maxW = Math.max(w * 0.95, 10);
  while (shown.length > 1 && ctx.measureText(shown).width > maxW) shown = shown.slice(0, -1);
  if (shown !== text) shown += '…';
  ctx.fillText(shown, tx, ty);
}

// 正文/要点（需要表明范围）：虚线范围框 + 多行灰色占位文字（按真实字号/对齐）
function drawBodyPlaceholder(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, ts: TemplateDoc['shapes'][number]['textStyle']): void {
  if (w < 10 || h < 6) return;
  const bold = !!(ts && ts.bold);
  const realPt = Number(ts && ts.size) > 0 ? Number(ts && ts.size) : 12;
  const size = Math.round(Math.min(realPt * 96 / 72, h * 0.9));
  setFont(ctx, size, bold, !!(ts && ts.italic));
  // 虚线范围框
  ctx.save();
  ctx.setLineDash([5, 4]);
  ctx.strokeStyle = '#C7CBD1';
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  ctx.restore();
  // 多行占位文字（浅灰，模拟正文排版范围）
  ctx.fillStyle = '#B9BEC5';
  const align = (ts && ts.align) || 'left';
  ctx.textAlign = align === 'center' ? 'center' : (align === 'right' ? 'right' : 'left');
  ctx.textBaseline = 'top';
  const maxW = w - Math.min(10, w * 0.08);
  const para = '这里是正文占位文字，用于预览该区域的文字排版效果。';
  const lines = wrapText(ctx, para.repeat(12), maxW);
  const lineH = size * 1.35;
  const maxLines = Math.max(1, Math.floor((h - size * 0.5) / lineH));
  const shown = lines.slice(0, maxLines);
  const valign = (ts && ts.valign) || 'middle';
  let ly = y + size * 0.5;
  if (valign === 'middle') ly = y + Math.max(size * 0.5, (h - shown.length * lineH) / 2);
  else if (valign === 'bottom') ly = y + h - shown.length * lineH;
  for (const ln of shown) {
    const tx = align === 'center' ? x + w / 2 : (align === 'right' ? x + w - Math.min(10, w * 0.08) : x + Math.min(10, w * 0.08));
    ctx.fillText(ln, tx, ly);
    ly += lineH;
  }
}

// 表格位：真实行列网格 + 中央「表格」
function drawTablePlaceholder(ctx: CanvasRenderingContext2D, s: TemplateDoc['shapes'][number], x: number, y: number, w: number, h: number): void {
  ctx.fillStyle = '#F3F4F6';
  ctx.fillRect(x, y, w, h);
  const t = s.table;
  const rows = (t && Number.isInteger(t.rows) && t.rows > 0) ? t.rows : 3;
  const cols = (t && Number.isInteger(t.cols) && t.cols > 0) ? t.cols : 3;
  ctx.strokeStyle = '#C9CDD3';
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  ctx.beginPath();
  for (let r = 1; r < rows; r++) { const gy = y + (h / rows) * r; ctx.moveTo(x, gy); ctx.lineTo(x + w, gy); }
  for (let c = 1; c < cols; c++) { const gx = x + (w / cols) * c; ctx.moveTo(gx, y); ctx.lineTo(gx, y + h); }
  ctx.stroke();
  drawShortText(ctx, '表格', x, y, w, h, { size: 13, align: 'center' }); // 占位标签居中
}

// 直线
function drawLine(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): void {
  ctx.strokeStyle = '#B8BEC6';
  ctx.lineWidth = 1;
  ctx.beginPath();
  if (w >= h) { ctx.moveTo(x, y + h / 2); ctx.lineTo(x + w, y + h / 2); }
  else { ctx.moveTo(x + w / 2, y); ctx.lineTo(x + w / 2, y + h); }
  ctx.stroke();
}

// 几何形状按保存的 shapeType/填充/线条绘制（椭圆→圆、其余→填充矩形；直线+圆圈点这类图案完整呈现）
function drawGeoShape(ctx: CanvasRenderingContext2D, s: TemplateDoc['shapes'][number], x: number, y: number, w: number, h: number): void {
  const st = s.shapeType || (s.type === 'line' ? 'line' : 'rect');
  if (st === 'line' || s.type === 'line') { drawLine(ctx, x, y, Math.max(w, 1), Math.max(h, 1)); return; }
  const fill = s.fill && s.fill.type === 'Solid' && s.fill.color ? '#' + String(s.fill.color).replace(/^#/, '') : 'transparent';
  const lineC = s.line && s.line.visible && s.line.color ? '#' + String(s.line.color).replace(/^#/, '') : 'transparent';
  ctx.fillStyle = fill;
  ctx.strokeStyle = lineC;
  ctx.lineWidth = 1;
  if (st === 'ellipse' || st === 'oval' || st === 'circle') {
    ctx.beginPath();
    ctx.ellipse(x + w / 2, y + h / 2, Math.max(w / 2, 0.5), Math.max(h / 2, 0.5), 0, 0, Math.PI * 2);
    ctx.fill();
    if (lineC !== 'transparent') ctx.stroke();
  } else {
    ctx.fillRect(x, y, w, h);
    if (lineC !== 'transparent') ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  }
}

export function renderTemplateDiagram(template: TemplateDoc): string {
  const ss = template.slideSize || { width: 10, height: 5.625 };
  const wIn = Number(ss.width) > 0 ? Number(ss.width) : 10;
  const hIn = Number(ss.height) > 0 ? Number(ss.height) : 5.625;
  const W = Math.round(wIn * PX_PER_INCH);
  const H = Math.round(hIn * PX_PER_INCH);
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';

  // 背景：solid 用模板色；picture 画图 cover；其余白
  const bg = template.background;
  if (bg && bg.type === 'solid' && bg.color) {
    ctx.fillStyle = bg.color;
    ctx.fillRect(0, 0, W, H);
  } else {
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, W, H);
  }
  if (bg && bg.type === 'picture' && bg.imageDataUrl) {
    const img = new Image();
    img.src = bg.imageDataUrl;
    if (img.complete && img.naturalWidth) ctx.drawImage(img, 0, 0, W, H);
    else img.onload = () => ctx.drawImage(img, 0, 0, W, H); // 异步：多数情况保存时已加载
  }

  for (const s of template.shapes || []) {
    const b = s.bounds;
    if (!b) continue;
    const x = Number(b.left) * PX_PER_INCH;
    const y = Number(b.top) * PX_PER_INCH;
    const w = Number(b.width) * PX_PER_INCH;
    const h = Number(b.height) * PX_PER_INCH;

    // 零尺寸 / 线：画线
    if (!(w > 0) || !(h > 0)) {
      if (s.type === 'line' || w > 0 || h > 0) drawLine(ctx, x, y, Math.max(w, 1), Math.max(h, 1));
      continue;
    }
    if (s.type === 'line' || Math.min(w, h) / Math.max(w, h) < 0.06) {
      drawLine(ctx, x, y, w, h);
      continue;
    }

    const isPic = s.type === 'picture' || s.role === 'ai_image';
    const isTable = s.type === 'table' || s.role === 'table';

    // 图片位：纯色填充 + 反色「图片」
    if (isPic) {
      ctx.fillStyle = '#D9D9D9';
      ctx.fillRect(x, y, w, h);
      ctx.strokeStyle = '#C6C6C6';
      ctx.lineWidth = 1;
      ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
      drawShortText(ctx, '图片', x, y, w, h, { size: 14, bold: true, align: 'center' }); // 占位标签居中
      continue;
    }
    if (isTable) {
      drawTablePlaceholder(ctx, s, x, y, w, h);
      continue;
    }

    // 固定元素：渲染真实内容（不画框；几何形状按原样绘制）
    if (s.role === 'fixed') {
      const hasTxt = typeof s.content === 'string' && s.content.trim() !== '';
      const isGeo = s.type === 'line' || ((s.type === 'rectangle' || s.type === 'other') && !hasTxt);
      if (isGeo) { drawGeoShape(ctx, s, x, y, w, h); continue; }
      const text = typeof s.content === 'string' ? s.content.trim() : '';
      if (text) drawShortText(ctx, text, x, y, w, h, s.textStyle);
      continue;
    }

    // 变量位（页码/序号）
    if (s.role === 'manual_var') {
      drawShortText(ctx, s.varName || '页码', x, y, w, h, s.textStyle);
      continue;
    }

    // AI 文本位：正文/要点等长文本 → 范围框+占位；短元素 → 角色占位字（不画框）
    if (s.role === 'ai_text') {
      if (s.semanticRole && RANGE_ROLES.has(s.semanticRole)) {
        drawBodyPlaceholder(ctx, x, y, w, h, s.textStyle);
      } else {
        drawShortText(ctx, ROLE_LABEL[s.semanticRole || ''] || '文本', x, y, w, h, s.textStyle);
      }
      continue;
    }

    // 兜底：其他类型 → 淡线框 + 占位
    ctx.strokeStyle = '#D5D9DE';
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    drawShortText(ctx, '元素', x, y, w, h, s.textStyle);
  }
  return canvas.toDataURL('image/png');
}
