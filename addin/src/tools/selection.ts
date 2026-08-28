// 小工具 · Office.js 交互层
// 职责：读取当前选中的形状（文本框/表格）、写回文本与表格尺寸。
// 版本兼容：表格逐列宽/行高调整依赖 PowerPointApi 1.9（preview）；不可用时降级为整体缩放。
// 文本写回：先整体替换 text，再逐段（getSubstring）恢复每段原字体/字号，避免强制统一格式。

import { estimateColumnWidths, estimateRowHeights, evenColumnWidths, evenRowHeights } from './tableOps.js';
import {
  normalizeSep, detectSep, restoreSep,
  indentLines, removeEmptyLines, separateParagraphs,
  locateTitleRange, locateBodyRange, splitParagraphs,
} from './textOps.js';

export type SelectedTarget =
  | { kind: 'text'; text: string; sep: string; widthPt: number; heightPt: number }
  | { kind: 'table'; values: string[][]; widthPt: number; heightPt: number }
  | null;

// ---------- 读取选中形状 ----------

// 读取当前选中的第一个形状；文本框返回文本内容，表格返回单元格矩阵。
export async function getSelectedTarget(): Promise<SelectedTarget> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load('items');
    await context.sync();
    const items = shapes.items;
    if (!items || items.length === 0) return null;
    const shape = items[0];
    shape.load(['type', 'width', 'height']);
    await context.sync();

    const kind = shape.type;
    if (kind === PowerPoint.ShapeType.table) {
      try {
        const table = shape.getTable();
        table.load('values');
        await context.sync();
        return {
          kind: 'table',
          values: (table.values || []).map((r: string[]) => r.map((c) => c ?? '')),
          widthPt: shape.width,
          heightPt: shape.height,
        };
      } catch {
        return { kind: 'table', values: [], widthPt: shape.width, heightPt: shape.height };
      }
    }

    // 文本框 / 占位符 / 其他可能带文本的形状：读 textFrame
    try {
      const tf = shape.textFrame;
      if (!tf) return null;
      const tr = tf.textRange;
      tr.load('text');
      await context.sync();
      const raw = (tr.text || '') as string;
      return {
        kind: 'text',
        text: normalizeSep(raw),
        sep: detectSep(raw),
        widthPt: shape.width,
        heightPt: shape.height,
      };
    } catch {
      return null;
    }
  });
}

// 取选中文本（用户也可能只选中文字而非整个形状，作为兜底信息返回）
export async function getSelectedText(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve((result.value as string) || '');
        } else {
          resolve('');
        }
      });
    } catch {
      resolve('');
    }
  });
}

// 段落切分统一用 textOps.splitParagraphs（兼容 \r\n / \r / \n，偏移基于实际文本）

// ---------- 写回文本框（保留每段格式） ----------

interface SegFontVals {
  name?: string; size?: number; bold?: boolean; italic?: boolean;
  color?: string; underline?: 'Single' | 'None' | null; strikethrough?: boolean;
}

// 读取每个段落的字体属性（getSubstring 定位每段 → font.load）
// 空段（len=0）跳过并返回 null 占位，保持索引对齐；调用方需判空。
function loadSegFonts(context: PowerPoint.RequestContext, tr: PowerPoint.TextRange, segs: { start: number; len: number }[]): (PowerPoint.ShapeFont | null)[] {
  return segs.map((s) => {
    if (s.len === 0) return null;
    const sub = tr.getSubstring(s.start, s.len);
    const f = sub.font;
    f.load(['name', 'size', 'bold', 'italic', 'color', 'underline', 'strikethrough']);
    return f;
  });
}

// 将新文本写回选中文本框，并逐段恢复原字体格式（段落数变化时按索引映射，超出的沿用最后一段格式）。
export async function applyTextToSelection(newText: string, sep: string): Promise<void> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load('items');
    await context.sync();
    const items = shapes.items;
    if (!items || items.length === 0) throw new Error('未选中文本框');
    const tr = items[0].textFrame.textRange;

    // 1) 读旧文本每段格式
    tr.load('text');
    await context.sync();
    const oldText = tr.text as string;
    const oldSegs = splitParagraphs(oldText);
    const oldFonts = loadSegFonts(context, tr, oldSegs);
    await context.sync();
    const fontVals: (SegFontVals | null)[] = oldFonts.map((f) => {
      if (!f) return null;
      return {
        name: f.name as string | undefined,
        size: f.size as number | undefined,
        bold: f.bold as boolean | undefined,
        italic: f.italic as boolean | undefined,
        color: f.color as string | undefined,
        underline: f.underline as 'Single' | 'None' | null | undefined,
        strikethrough: f.strikethrough as boolean | undefined,
      };
    });

    // 2) 整体替换文本
    tr.text = restoreSep(newText, sep);
    await context.sync();

    // 3) 逐段恢复格式（新文本段落索引重新计算）
    // 格式映射只按「非空段」编号对应：删除空行/插入空段后，标题格式不会错位到正文。
    const newTextActual = tr.text as string;
    const newSegs = splitParagraphs(newTextActual);
    const fontsToSet = loadSegFonts(context, tr, newSegs);
    await context.sync();

    // 旧格式按非空段收集（保持顺序）
    const nonEmptyOld: SegFontVals[] = fontVals.filter((v): v is SegFontVals => v !== null && Object.values(v).some((x) => x !== undefined && x !== null));
    if (nonEmptyOld.length > 0) {
      let srcIdx = 0;
      for (let i = 0; i < newSegs.length; i++) {
        if (newSegs[i].len === 0) continue; // 空段无需恢复
        const src = nonEmptyOld[Math.min(srcIdx, nonEmptyOld.length - 1)];
        const f = fontsToSet[i];
        srcIdx++;
        if (!src || !f) continue;
        if (src.name) f.name = src.name;
        if (src.size) f.size = src.size;
        if (src.bold !== undefined) f.bold = src.bold;
        if (src.italic !== undefined) f.italic = src.italic;
        if (src.color) f.color = src.color;
        if (src.underline !== undefined && src.underline !== null) f.underline = src.underline;
        if (src.strikethrough !== undefined) f.strikethrough = src.strikethrough;
      }
    }
    await context.sync();
  });
}

// ---------- 表格尺寸 ----------

export interface TableApplyResult {
  ok: boolean;
  degraded?: string; // 降级说明（整体缩放时给出）
}

// 按指定列宽/行高数组写入选中表格；PowerPointApi 1.9 不可用时降级为整体缩放。
export async function applyTableSizes(colWidths: number[], rowHeights: number[]): Promise<TableApplyResult> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load('items');
      await context.sync();
      const items = shapes.items;
      if (!items || items.length === 0) throw new Error('未选中表格');
      const shape = items[0];
      const table = shape.getTable();
      table.columns.load('items');
      table.rows.load('items');
      await context.sync();
      const cols = table.columns.items;
      const rows = table.rows.items;
      for (let i = 0; i < Math.min(cols.length, colWidths.length); i++) {
        cols[i].width = colWidths[i];
      }
      for (let i = 0; i < Math.min(rows.length, rowHeights.length); i++) {
        rows[i].height = rowHeights[i];
      }
      await context.sync();
    });
    return { ok: true };
  } catch (err) {
    // 1.9 不可用（老版本 Office）：降级为整体缩放选中表格 shape
    try {
      const totalW = colWidths.reduce((a, b) => a + b, 0);
      const totalH = rowHeights.reduce((a, b) => a + b, 0);
      if (totalW <= 0 || totalH <= 0) throw new Error('尺寸无效');
      await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load('items');
        await context.sync();
        const items = shapes.items;
        if (!items || items.length === 0) throw new Error('未选中表格');
        const shape = items[0];
        shape.width = totalW;
        shape.height = totalH;
        await context.sync();
      });
      return { ok: true, degraded: '当前 Office 版本不支持逐列调整，已按整体缩放表格' };
    } catch (err2) {
      const msg = err2 instanceof Error ? err2.message : String(err2);
      throw new Error('表格调整失败：' + msg);
    }
  }
}

// ---------- 页面宽度（适配表格用） ----------

// 尝试读取页面宽度（pt）；PageSetup 是 1.10 preview，失败返回 null（调用方用当前表格宽兜底）。
export async function getPageWidthPt(): Promise<number | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const ps = context.presentation.pageSetup;
      ps.load('slideWidth');
      await context.sync();
      return (ps.slideWidth as number) || null;
    });
  } catch {
    return null;
  }
}

// ---------- 高层封装：工具动作 ----------

export interface ToolActionResult {
  ok: boolean;
  message: string;
}

// 文本类动作：读选中 → 变换 → 写回（保留每段格式）。返回处理前后是否变化。
async function runTextAction(transform: (text: string) => string, actionName: string): Promise<ToolActionResult> {
  const target = await getSelectedTarget();
  if (!target) return { ok: false, message: '请先选中一个文本框' };
  if (target.kind !== 'text') return { ok: false, message: '当前选中的是表格，请先选中文本框' };
  const before = target.text;
  const after = transform(before);
  if (before === after) return { ok: true, message: actionName + '：无需处理' };
  await applyTextToSelection(after, target.sep);
  return { ok: true, message: actionName + '完成' };
}

export async function actionRemoveSpaces(): Promise<ToolActionResult> {
  return runTextAction((t) => t.replace(/[\u0020\u00A0\u3000]/g, ''), '去除空格');
}

export async function actionIndentParagraphs(): Promise<ToolActionResult> {
  return runTextAction(indentLines, '段首缩进');
}

export async function actionRemoveEmptyParagraphs(): Promise<ToolActionResult> {
  return runTextAction(removeEmptyLines, '删除空行空段');
}

export async function actionSeparateParagraphs(): Promise<ToolActionResult> {
  return runTextAction(separateParagraphs, '分隔每段');
}

// 选中标题：高亮文本框内标题部分（setSelected），不修改内容。
export async function actionSelectTitle(): Promise<ToolActionResult> {
  return selectTextRange('标题');
}

// 选中正文：高亮文本框内标题之后的正文部分（setSelected），不修改内容。
export async function actionSelectBody(): Promise<ToolActionResult> {
  return selectTextRange('正文');
}

async function selectTextRange(kind: '标题' | '正文'): Promise<ToolActionResult> {
  const target = await getSelectedTarget();
  if (!target) return { ok: false, message: '请先选中一个文本框' };
  if (target.kind !== 'text') return { ok: false, message: '当前选中的是表格，请先选中文本框' };

  try {
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load('items');
      await context.sync();
      const items = shapes.items;
      if (!items || items.length === 0) throw new Error('未选中文本框');
      const tr = items[0].textFrame.textRange;
      tr.load('text');
      await context.sync();
      // 用实际文本（原始分隔符）计算偏移，避免归一化导致 \r\n 长度错位
      const actualText = (tr.text || '') as string;
      const range = kind === '标题' ? locateTitleRange(actualText) : locateBodyRange(actualText);
      if (!range || range.length === 0) {
        return { ok: false, message: '未识别到' + kind + '（文本框内容可能为空或格式特殊）' };
      }
      const sub = tr.getSubstring(range.start, range.length);
      sub.setSelected();
      await context.sync();
      return { ok: true, message: '已选中' + kind + '（' + range.length + ' 字）' };
    });
  } catch (err) {
    return { ok: false, message: '选中' + kind + '失败：' + (err instanceof Error ? err.message : String(err)) };
  }
}

// 表格类动作
async function runTableAction(
  compute: (t: { values: string[][]; widthPt: number; heightPt: number }, pageW: number | null) => { cols: number[]; rows: number[] },
  actionName: string,
): Promise<ToolActionResult> {
  const target = await getSelectedTarget();
  if (!target) return { ok: false, message: '请先选中一个表格' };
  if (target.kind !== 'table') return { ok: false, message: '当前选中的是文本框，请先选中表格' };
  if (!target.values || target.values.length === 0) {
    return { ok: false, message: '未读取到表格内容，请重新选中表格' };
  }
  const pageW = await getPageWidthPt();
  const { cols, rows } = compute(target, pageW);
  if (!cols.length || !rows.length) {
    return { ok: false, message: '未读取到表格行列信息' };
  }
  const res = await applyTableSizes(cols, rows);
  return { ok: true, message: res.degraded ? actionName + '（' + res.degraded + '）' : actionName + '完成' };
}

export async function actionFitTable(): Promise<ToolActionResult> {
  return runTableAction((t, pageW) => {
    const pageAvail = pageW ? Math.max(pageW - 48, t.widthPt * 0.5) : Math.max(t.widthPt, 480);
    const cols = estimateColumnWidths(t.values, t.widthPt, pageAvail);
    const rows = estimateRowHeights(t.values, cols);
    return { cols, rows };
  }, '表格最佳适配');
}

export async function actionEvenTable(): Promise<ToolActionResult> {
  return runTableAction((t) => {
    const nCols = t.values.length > 0 ? Math.max(...t.values.map((r) => r.length)) : 0;
    const nRows = t.values.length;
    return { cols: evenColumnWidths(nCols, t.widthPt), rows: evenRowHeights(nRows, t.heightPt) };
  }, '行列均分');
}