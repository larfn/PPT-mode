import { TemplateDoc, TemplateShape } from './api.js';
import { escapeAttr, escapeHtml } from './lib/html.js';
import { expandGrid } from './lib/tableModel.js';
import type { FitCell, FitResult } from './lib/tableModel.js';

const INCH_PX = 96;

// 从保存的表格结构生成默认 2D 数据（首行为表头，其余为模板原文）——与向导内逻辑一致，供预览独立使用
function defaultTableData(s: { table?: { rows?: number; cols?: number; cells?: { row: number; col: number; text?: string }[] } }): string[][] {
  const t = s.table;
  if (!t) return [[]];
  const rows = t.rows || 1, cols = t.cols || 1;
  const out: string[][] = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));
  for (const c of t.cells || []) {
    if (c.row >= 0 && c.row < rows && c.col >= 0 && c.col < cols) out[c.row][c.col] = c.text || '';
  }
  return out;
}

// 几何形状预览：按保存的 shapeType/填充/线条原样绘制（矩形/圆角/椭圆/线条等），
// 与生成端 slideBuilder 渲染一致 —— 直线+圆圈点这类图案在预览里必须完整可见
function renderGeoHtml(s: TemplateShape, style: string): string {
  const fill = (s.fill && s.fill.type === 'Solid' && s.fill.color) ? s.fill.color : '';
  const lineColor = (s.line && s.line.color) || '#888888';
  const st = s.shapeType || (s.type === 'line' ? 'line' : 'rect');
  if (st === 'line') {
    const vertical = (s.bounds.width || 0) < (s.bounds.height || 0) - 0.001;
    const border = vertical
      ? 'width:1px;height:100%;border-left:1px solid ' + lineColor + ';'
      : 'width:100%;height:1px;border-top:1px solid ' + lineColor + ';';
    return `<div style="${style}"><div style="position:relative;left:0;top:0;${border}"></div></div>`;
  }
  const radius = (st === 'ellipse' || st === 'oval' || st === 'circle') ? '50%' : (st === 'roundRect' ? '10px' : '0');
  const bg = fill ? 'background:' + fill + ';' : '';
  const border = 'border:1px solid ' + (s.line && s.line.visible ? lineColor : '#d0d0d0') + ';';
  return `<div style="${style}"><div style="width:100%;height:100%;border-radius:${radius};${bg}${border}"></div></div>`;
}
// images: Record<shapeId, dataURL>（多图片位）；兼容旧调用传单个 string（作用于所有 AI 图片位）
// tableData: Record<shapeId, string[][]>（表格位实时数据）；缺省时用模板保存的原文
// tableFits / tableMerges：表格位排版引擎 fit 结果（可选；不传时表格回退原有 2D 网格渲染）
export function renderPreview(
  template: TemplateDoc,
  images: Record<string, string> | string,
  texts: Record<string, string>,
  vars: Record<string, string>,
  tableData?: Record<string, string[][]>,
  tableFits?: Record<string, FitResult>,
  tableMerges?: Record<string, FitCell[]>
): string {
  const imgOf = (shapeId: string): string => typeof images === 'string' ? images : (images[shapeId] || '');
  const { width, height } = template.slideSize || { width: 13.33, height: 7.5 };
  const canvasW = Math.max(1, Math.round(width * INCH_PX));
  const canvasH = Math.max(1, Math.round(height * INCH_PX));
  const blocks = template.shapes.map((s) => {
    const { left, top, width: w, height: h } = s.bounds;
    const style = `position:absolute;left:${(left * INCH_PX).toFixed(2)}px;top:${(top * INCH_PX).toFixed(2)}px;width:${(w * INCH_PX).toFixed(2)}px;height:${(h * INCH_PX).toFixed(2)}px;`;
    if (s.role === 'ai_image') {
      const d = imgOf(s.id);
      return d ? `<div style="${style}"><img src="${escapeAttr(d)}" style="width:100%;height:100%;object-fit:cover" /></div>` : '';
    }
    if (s.role === 'fixed' && s.type === 'picture') return '';
    // —— 表格位：实时渲染用户编辑/导入的数据（无数据时回退模板原文）——
    if (s.role === 'table' && s.table) {
      const data = (tableData && tableData[s.id]) || defaultTableData(s);
      const fit = (tableFits && tableFits[s.id]) || undefined;
      // 有 fit 结果：按引擎排版渲染（列宽按 colWidths 比例、字号按 fontSize、rowspan/colspan 反映合并、超出红虚线 + 角标）
      if (fit && fit.cells && fit.cells.length) {
        const sumW = fit.colWidths.reduce((a, b) => a + b, 0) || 1;
        const sumH = fit.rowHeights.reduce((a, b) => a + b, 0) || 1;
        const slotH = Math.max(s.bounds.height || sumH, 1e-6);
        // 超出兜底：宽度锁槽宽、高度向下延伸（预览里撑出槽位，让用户看到会延伸多少）
        const heightPct = fit.overflow ? Math.max(100, (sumH / slotH) * 100) : 100;
        const colsHtml = fit.colWidths.map((cw) => `<col style="width:${((cw / sumW) * 100).toFixed(3)}%" />`).join('');
        const grid = expandGrid(fit.cells, fit.rows, fit.cols);
        const trs: string[] = [];
        for (let r = 0; r < fit.rows; r++) {
          let tds = '';
          for (let c = 0; c < fit.cols; c++) {
            const cell = grid[r][c];
            if (!cell || cell.r !== r || cell.c !== c) continue; // 真空位 / 被合并覆盖位跳过
            const span = (cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : '') + (cell.colspan > 1 ? ` colspan="${cell.colspan}"` : '');
            tds += `<td${span} style="padding:1px 3px;overflow:hidden;white-space:nowrap">${escapeHtml(cell.text || '')}</td>`;
          }
          trs.push(`<tr style="height:${((fit.rowHeights[r] / sumH) * 100).toFixed(3)}%">${tds}</tr>`);
        }
        const fontSizePx = (fit.fontSize * 96 / 72).toFixed(1); // px ≈ pt × 1.333（96dpi）
        return `<div style="${style}overflow:visible">
          <div class="wb-pv-tbl-wrap${fit.overflow ? ' wb-pv-tbl-overflow' : ''}" style="position:relative;width:100%;height:${heightPct.toFixed(1)}%">
            <table class="wb-pv-tbl" style="font-size:${fontSizePx}px">
              <colgroup>${colsHtml}</colgroup>
              <tbody>${trs.join('')}</tbody>
            </table>
            ${fit.overflow ? '<span class="wb-pv-overflow-badge">超出</span>' : ''}
          </div>
        </div>`;
      }
      // 回退（旧调用未传 fit）：原有 2D 网格渲染
      const nc = Math.max(1, (data[0] || []).length);
      const cells = data.map((row) => row.map((cell) =>
        `<div style="border-right:1px solid #bbb;border-bottom:1px solid #bbb;padding:1px 3px;overflow:hidden;font-size:10px;line-height:1.3;white-space:nowrap">${escapeHtml(cell || '')}</div>`).join('')).join('');
      return `<div style="${style}overflow:hidden"><div style="display:grid;grid-template-columns:repeat(${nc},1fr);border-top:1px solid #bbb;border-left:1px solid #bbb;width:100%;height:100%">${cells}</div></div>`;
    }
    // —— 几何形状（矩形/椭圆/线条等，含无文本的 fixed 其它形状）：原样绘制，预览与生成一致 ——
    // 仅当该形状不是文字位且没有文本内容时才画图形（矩形里放文字的 ai_text/fixed 仍按文字渲染）
    const isTextSlot = s.role === 'ai_text' || s.role === 'manual_var';
    const hasContentText = typeof s.content === 'string' && s.content.trim() !== '';
    if (!isTextSlot && !hasContentText && (s.type === 'rectangle' || s.type === 'line' || s.type === 'other')) {
      return renderGeoHtml(s, style);
    }
    const text = s.role === 'ai_text' ? (texts[s.id] || '') : s.role === 'manual_var' ? (vars[s.id] || '') : (s.content || '');
    const ts = s.textStyle || {};
    const deco: string[] = [];
    if (ts.underline) deco.push('underline');
    if (ts.strikethrough || ts.doubleStrikethrough) deco.push('line-through');
    const decoStyle = deco.length ? `text-decoration:${deco.join(' ')};` : '';
    // 字号：模板存 pt，屏幕 96dpi 下 1pt ≈ 1.333px；pre-wrap 保留换行、break-word 长英文可换行；
    // overflow:visible 保证超长文本（如自动翻译出的长英文副标题）完整呈现、不被裁剪
    const sizePx = ((Number(ts.size) || 16) * 96 / 72).toFixed(1);
    const tstyle = `font-family:${ts.font || 'Microsoft YaHei'};font-size:${sizePx}px;font-weight:${ts.bold ? 'bold' : 'normal'};color:${ts.color || '#333'};text-align:${ts.align || 'left'};${decoStyle}white-space:pre-wrap;overflow-wrap:break-word;word-break:break-word;line-height:1.25;`;
    return `<div style="${style}"><div style="${tstyle}">${escapeHtml(text)}</div></div>`;
  }).join('');
  const bg = template.background;
  const bgStyle = bg && bg.type === 'solid' && bg.color
    ? `background:${bg.color};`
    : 'background:#fff;';
  const bgImg = bg && bg.type === 'picture' && bg.imageDataUrl
    ? `<img src="${escapeAttr(bg.imageDataUrl)}" style="position:absolute;left:0;top:0;width:100%;height:100%;object-fit:cover" />`
    : '';
  return `<svg viewBox="0 0 ${canvasW} ${canvasH}" preserveAspectRatio="xMinYMin meet" style="display:block;width:100%;height:auto;aspect-ratio:${width}/${height};border:1px solid #ddd;background:#fff;overflow:hidden">
    <foreignObject x="0" y="0" width="${canvasW}" height="${canvasH}">
      <div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${canvasW}px;height:${canvasH}px;${bgStyle}overflow:hidden">${bgImg}${blocks}</div>
    </foreignObject>
  </svg>`;
}
