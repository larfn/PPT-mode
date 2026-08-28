// 小工具 · 表格尺寸计算纯函数（无 Office 依赖，可直接单测）
// 单位：磅（pt）。列宽/行高最终由调用方写入 PowerPoint.TableColumn.width / TableRow.height。

export interface TableSizeOptions {
  // 默认正文字号（pt），用于估算字符宽度
  fontSizePt?: number;
  // 单元格左右边距总和（pt）
  cellPaddingX?: number;
  // 单元格上下边距总和（pt）
  cellPaddingY?: number;
  // 单行文本基准行高（pt）
  lineHeightPt?: number;
  // 最小列宽（pt）
  minColumnPt?: number;
}

// ---------- 字符宽度估算 ----------

// 估算一段文本的显示宽度（单位：em，即相对于 fontSizePt 的倍数）。
// 中文/全角 ≈ 1.0em，半角字母数字 ≈ 0.55em，空格 ≈ 0.3em，其余按 0.55em 兜底。
export function estimateTextWidthEm(text: string): number {
  let w = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code === 0x20 || code === 0x3000) w += 0.3; // 半角/全角空格
    else if (
      (code >= 0x4e00 && code <= 0x9fff) ||   // CJK 统一汉字
      (code >= 0x3000 && code <= 0x303f) ||   // CJK 标点
      (code >= 0xff00 && code <= 0xffef) ||   // 全角字符
      (code >= 0x2e80 && code <= 0x2eff) ||   // 部首
      (code >= 0x3400 && code <= 0x4dbf)      // CJK 扩展 A
    ) w += 1.0;
    else w += 0.55;
  }
  return w;
}

// 估算单元格文本所需宽度（pt）= 字符 em 数 × 字号 + 左右边距
export function estimateCellWidthPt(text: string, opts: TableSizeOptions): number {
  const fontSize = opts.fontSizePt ?? 12;
  const padX = opts.cellPaddingX ?? 12;
  return estimateTextWidthEm(text) * fontSize + padX;
}

// ---------- 列宽：按内容最佳适配 ----------

// 输入 values[row][col]，输出每列宽（pt）。
// 策略：内容估算宽按比例分配，总宽收敛到目标总宽（不超出页面可用宽，且不小于当前宽的一半）。
export function estimateColumnWidths(
  values: string[][],
  currentWidthPt: number,
  pageAvailableWidthPt: number,
  opts: TableSizeOptions = {},
): number[] {
  const cols = values.length > 0 ? Math.max(...values.map((r) => r.length)) : 0;
  if (cols === 0) return [];
  const minCol = opts.minColumnPt ?? 40;

  // 每列内容估算宽度
  const contentW: number[] = new Array(cols).fill(0);
  for (const row of values) {
    for (let c = 0; c < cols; c++) {
      const cell = row[c] ?? '';
      contentW[c] = Math.max(contentW[c], estimateCellWidthPt(cell, opts));
    }
  }
  const contentTotal = contentW.reduce((a, b) => a + b, 0);

  // 目标总宽：内容优先，但受页面可用宽与当前宽约束
  const targetTotal = Math.min(
    Math.max(contentTotal, currentWidthPt * 0.5),
    Math.max(pageAvailableWidthPt, minCol * cols),
  );

  // 按内容比例分配目标总宽
  const scale = contentTotal > 0 ? targetTotal / contentTotal : 1;
  let widths = contentW.map((w) => w * scale);

  // 保证最小列宽（超出时从最大列扣减）
  const overflow = widths.reduce((sum, w) => sum + Math.max(0, minCol - w), 0);
  if (overflow > 0) {
    const total = widths.reduce((a, b) => a + b, 0);
    const shrinkable = widths.filter((w) => w > minCol);
    if (shrinkable.length > 0 && total > 0) {
      const excess = Math.min(overflow, shrinkable.reduce((a, w) => a + (w - minCol), 0));
      widths = widths.map((w) => {
        if (w <= minCol) return w;
        const share = (w - minCol) / (shrinkable.reduce((a, x) => a + Math.max(0, x - minCol), 0) || 1);
        return Math.max(minCol, w - excess * share);
      });
    }
  }
  return widths;
}

// ---------- 行高：按内容折行估算 ----------

// 每行高度 = 折行行数 × 行高 + 上下边距；表头行（第 0 行）不低于内容行。
export function estimateRowHeights(
  values: string[][],
  columnWidthsPt: number[],
  opts: TableSizeOptions = {},
): number[] {
  const lineH = opts.lineHeightPt ?? 14;
  const padY = opts.cellPaddingY ?? 8;
  const fontSize = opts.fontSizePt ?? 12;

  return values.map((row, r) => {
    let maxLines = 1;
    for (let c = 0; c < columnWidthsPt.length; c++) {
      const cell = row[c] ?? '';
      const em = estimateTextWidthEm(cell);
      // 可用内容宽度 = 列宽 - 左右边距
      const usable = Math.max(columnWidthsPt[c] - (opts.cellPaddingX ?? 12), fontSize);
      const lines = Math.max(1, Math.ceil((em * fontSize) / usable));
      maxLines = Math.max(maxLines, lines);
    }
    return Math.max(lineH, maxLines * lineH + padY);
  });
}

// ---------- 行列均分 ----------

export function evenColumnWidths(cols: number, totalWidthPt: number): number[] {
  if (cols <= 0) return [];
  const w = totalWidthPt / cols;
  return new Array(cols).fill(w);
}

export function evenRowHeights(rows: number, totalHeightPt: number): number[] {
  if (rows <= 0) return [];
  const h = totalHeightPt / rows;
  return new Array(rows).fill(h);
}
