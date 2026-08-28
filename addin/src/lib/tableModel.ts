// 表格自动排版引擎（阶段 0/2/3 地基）：统一表格数据模型 + 排版引擎（纯函数）
// 纯 TS：顶层不 import 任何 DOM/Office 模块 —— Node 24 可直接 require（type-stripping），浏览器/Vite 亦可直接使用。
// 单位约定：colWidths/rowHeights = 英寸；fontSize = 磅；measure 回调返回文本宽度（磅）。

export interface FitCell {
  r: number;
  c: number;
  rowspan: number;
  colspan: number;
  text: string;
}

export interface FitResult {
  rows: number;
  cols: number;
  colWidths: number[];   // 英寸；fit 成功时总和 = slotWidthIn
  rowHeights: number[];  // 英寸；总和可能 > slotHeightIn（overflow 向下延伸）
  fontSize: number;      // 磅，全表统一，>= fontFloorPt
  cells: FitCell[];      // 与输入一致（主格定义）
  overflow: boolean;     // true = 缩到下限仍放不下 → 宽度锁槽宽、高度向下延伸
  scaleRatio: number;    // 1 = 未缩字；<1 = 等比缩小比例
  baseFontSizePt: number;
}

export interface TableFitOptions {
  slotWidthIn: number;
  slotHeightIn: number;
  baseFontSizePt?: number;   // 默认 14
  fontFloorPt?: number;      // 默认 10
  fontFace?: string;         // 测量用字体（可选）
  cellPaddingXIn?: number;   // 默认 0.1（PPT 默认左右边距 0.1"）
  cellPaddingYIn?: number;   // 默认 0.05
  lineHeightFactor?: number; // 默认 1.4（行高 = 字号 × factor）
  measure?: (text: string, fontSizePt: number, fontFace?: string) => number; // 返回文本宽度（磅）；缺省用 em 估算
}

const PT_PER_IN = 72;

// ---------- 字符分类（与 tableOps.estimateTextWidthEm 一致） ----------
function isCjkCode(code: number): boolean {
  return (code >= 0x4e00 && code <= 0x9fff) ||   // CJK 统一汉字
    (code >= 0x3000 && code <= 0x303f) ||        // CJK 标点
    (code >= 0xff00 && code <= 0xffef) ||        // 全角字符
    (code >= 0x2e80 && code <= 0x2eff) ||        // 部首
    (code >= 0x3400 && code <= 0x4dbf);          // CJK 扩展 A
}

function isCjk(ch: string): boolean {
  return isCjkCode(ch.codePointAt(0) ?? 0);
}

// 缺省度量：em 估算（CJK ≈ 1.0em、半角 ≈ 0.55em、空格 ≈ 0.3em；宽 = em 数 × 字号）
function defaultMeasure(text: string, fontSizePt: number): number {
  let ems = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code === 0x20 || code === 0x3000) ems += 0.3;
    else if (isCjkCode(code)) ems += 1.0;
    else ems += 0.55;
  }
  return ems * fontSizePt;
}

// 把一段文本拆成「不可断单元」：CJK 逐字、Latin 按词（空格断开）、数字/URL 整段不可断。
// 返回单元数组（空格本身不是单元，只是断点）。
function splitUnits(seg: string): string[] {
  const units: string[] = [];
  let cur = '';
  const flush = () => { if (cur) { units.push(cur); cur = ''; } };
  for (const ch of seg) {
    if (ch === ' ' || ch === '\u3000') flush();
    else if (isCjk(ch)) { flush(); units.push(ch); }
    else cur += ch;
  }
  flush();
  return units;
}

// 单段（不含 \n）贪心折行：按不可断单元累积，放不下换行；
// 单元本身超宽（锁宽/溢出模式下列宽小于最长单元）时按可容纳宽度折算多行（内容继续折行）。
function wrapSegment(seg: string, widthPt: number, fontSizePt: number, measure: (t: string, s: number) => number): number {
  if (!seg) return 1; // 空段 = 一个空行
  const units = splitUnits(seg);
  let lines = 0;
  let w = 0;
  for (const u of units) {
    const uw = measure(u, fontSizePt);
    if (uw > widthPt) {
      lines += Math.max(1, Math.ceil(uw / Math.max(widthPt, 1)));
      w = 0;
      continue;
    }
    if (w > 0 && w + uw > widthPt) { lines++; w = uw; }
    else w += uw;
  }
  if (w > 0) lines++;
  return Math.max(1, lines);
}

// 整格折行行数：\n 强制分段，段内按宽折行；至少 1 行
function countWrapLines(text: string, widthPt: number, fontSizePt: number, measure: (t: string, s: number) => number): number {
  let total = 0;
  for (const seg of text.split('\n')) total += wrapSegment(seg, widthPt, fontSizePt, measure);
  return Math.max(1, total);
}

// 每格 min_w / max_w（磅）：min = 各行「最长不可断单元」的最大值；max = 各行单行不折宽的最大值
function cellMinMaxWidthPt(text: string, fontSizePt: number, measure: (t: string, s: number) => number): { min: number; max: number } {
  let min = 0;
  let max = 0;
  for (const seg of text.split('\n')) {
    max = Math.max(max, measure(seg, fontSizePt));
    let longest = 0;
    for (const u of splitUnits(seg)) longest = Math.max(longest, measure(u, fontSizePt));
    min = Math.max(min, longest);
  }
  return { min, max };
}

// 归一化输入格：防御非法值
function normalizeCells(cells: FitCell[]): FitCell[] {
  return (cells || [])
    .filter((c) => c && Number.isFinite(c.r) && Number.isFinite(c.c) && c.r >= 0 && c.c >= 0)
    .map((c) => ({
      r: c.r,
      c: c.c,
      rowspan: Math.max(1, Math.floor(c.rowspan || 1)),
      colspan: Math.max(1, Math.floor(c.colspan || 1)),
      text: String(c.text ?? '')
    }));
}

// ---------- 排版引擎（阶段 2，全部基于「测量」而非猜测） ----------
export function fitTableLayout(cells: FitCell[], opts: TableFitOptions): FitResult {
  const baseFontSizePt = opts.baseFontSizePt ?? 14;
  const fontFloorPt = opts.fontFloorPt ?? 10;
  const cellPaddingXIn = opts.cellPaddingXIn ?? 0.1;
  const cellPaddingYIn = opts.cellPaddingYIn ?? 0.05;
  const lineHeightFactor = opts.lineHeightFactor ?? 1.4;
  const measure = opts.measure || defaultMeasure;
  const measureAt = (t: string, s: number) => measure(t, s, opts.fontFace);
  const padXpt = cellPaddingXIn * PT_PER_IN;
  const padYpt = cellPaddingYIn * PT_PER_IN;

  // 防御：过滤非法格，推导行列（以 cells 实际为准）
  const norm = normalizeCells(cells);
  const rows = norm.length ? Math.max(1, ...norm.map((c) => c.r + c.rowspan)) : 1;
  const cols = norm.length ? Math.max(1, ...norm.map((c) => c.c + c.colspan)) : 1;

  // 每格 min_w / max_w（英寸，含左右 padding）
  const cellWIn = (cell: FitCell, fontSizePt: number) => {
    const { min, max } = cellMinMaxWidthPt(cell.text, fontSizePt, measureAt);
    return { min: (min + padXpt * 2) / PT_PER_IN, max: (max + padXpt * 2) / PT_PER_IN };
  };

  // 槽宽：非法（<=0 / 非数）时视为无约束（用内容总宽）
  const rawSlotW = Number(opts.slotWidthIn);
  let slotW = 0;

  // ① 列最小宽：非合并格贡献 min_w 到自己的列；合并格按 colspan 等分摊到所跨列
  const colMin = new Array<number>(cols).fill(0);
  const colMax = new Array<number>(cols).fill(0); // 各列「想变宽程度」上限（贪心用）
  for (const cell of norm) {
    const { min, max } = cellWIn(cell, baseFontSizePt);
    const k = cell.colspan;
    for (let j = 0; j < k && cell.c + j < cols; j++) {
      colMin[cell.c + j] = Math.max(colMin[cell.c + j], min / k);
      colMax[cell.c + j] = Math.max(colMax[cell.c + j], max / k);
    }
  }
  // ② 迭代归一化（最多 10 轮）：合并格所跨列之和不足该格 min_w → 差额按各列当前最小宽比例补足
  for (let round = 0; round < 10; round++) {
    let changed = false;
    for (const cell of norm) {
      if (cell.colspan <= 1) continue;
      let spannedSum = 0;
      for (let j = 0; j < cell.colspan && cell.c + j < cols; j++) spannedSum += colMin[cell.c + j];
      const need = cellWIn(cell, baseFontSizePt).min;
      if (spannedSum < need - 1e-9) {
        const deficit = need - spannedSum;
        for (let j = 0; j < cell.colspan && cell.c + j < cols; j++) {
          const share = spannedSum > 0
            ? colMin[cell.c + j] / spannedSum
            : 1 / Math.min(cell.colspan, cols - cell.c);
          colMin[cell.c + j] += deficit * share;
        }
        changed = true;
      }
    }
    if (!changed) break;
  }

  let sum = colMin.reduce((a, b) => a + b, 0);
  slotW = (Number.isFinite(rawSlotW) && rawSlotW > 0) ? rawSlotW : sum;
  const slotHeightIn = (Number.isFinite(Number(opts.slotHeightIn)) && Number(opts.slotHeightIn) > 0) ? Number(opts.slotHeightIn) : 0;

  let fontSize = baseFontSizePt;
  let overflow = false;
  let colW: number[];

  if (sum <= slotW + 1e-9) {
    // ③ 总宽适配：剩余空间按「贪心度 = max(0, max_w − col_min)」分配（文字多的列多拿）；全 0 则均分
    const slack = slotW - sum;
    const greed = colMin.map((m, c) => Math.max(0, colMax[c] - m));
    const totalGreed = greed.reduce((a, b) => a + b, 0);
    if (totalGreed > 1e-12) {
      colW = colMin.map((m, c) => m + slack * (greed[c] / totalGreed));
    } else {
      const each = cols > 0 ? slack / cols : 0;
      colW = colMin.map((m) => m + each);
    }
    fontSize = baseFontSizePt;
  } else {
    // ④ 整表等比缩字：ratio = 槽宽 / 总最小宽；字号不低于下限
    const ratio = slotW / sum;
    fontSize = Math.max(fontFloorPt, baseFontSizePt * ratio);
    const colMinAtFont = colMin.map((m) => m * (fontSize / baseFontSizePt));
    const sumAtFont = colMinAtFont.reduce((a, b) => a + b, 0);
    if (sumAtFont <= slotW + 1e-9) {
      // 缩到可用字号后刚好放下（fontSize > 下限 或 恰好等于下限且恰好放得下）
      colW = colMinAtFont;
    } else {
      // ⑤ 超出兜底：缩到下限仍放不下 → 宽度锁槽宽（等比压缩，允许更多折行），高度向下延伸
      const factor = slotW / sumAtFont;
      colW = colMinAtFont.map((m) => m * factor);
      overflow = true;
    }
  }

  // 修正浮点误差：总和精确 = slotW（差额并入最后一列）
  {
    const s = colW.reduce((a, b) => a + b, 0);
    const diff = slotW - s;
    if (Math.abs(diff) > 1e-9 && colW.length) colW[colW.length - 1] += diff;
  }

  // ⑥ 行高：列宽定稿 + 最终字号下逐格折行算行数；
  //    行高[r] = max over 该行各格(行数 × 字号 × factor + 上下 padding)；
  //    rowspan>1 的格：需求高度超出所跨行当前总和的部分加到所跨最后一行（保证所跨行总和 ≥ 该格需求）
  const rowH = new Array<number>(rows).fill(0);
  const cellHeightIn = (cell: FitCell) => {
    let spanW = 0;
    for (let j = 0; j < cell.colspan && cell.c + j < cols; j++) spanW += colW[cell.c + j];
    const availW = Math.max(spanW - cellPaddingXIn * 2, 0.05); // 可用宽 = 列宽 − 左右 padding
    const lines = countWrapLines(cell.text, availW * PT_PER_IN, fontSize, measureAt);
    return (lines * fontSize * lineHeightFactor + padYpt * 2) / PT_PER_IN;
  };
  for (const cell of norm) {
    if (cell.rowspan === 1 && cell.r < rows) {
      rowH[cell.r] = Math.max(rowH[cell.r], cellHeightIn(cell));
    }
  }
  for (const cell of norm) {
    if (cell.rowspan > 1 && cell.r < rows) {
      const last = Math.min(cell.r + cell.rowspan - 1, rows - 1);
      let total = 0;
      for (let i = cell.r; i <= last; i++) total += rowH[i];
      const need = cellHeightIn(cell);
      if (need > total) rowH[last] += need - total;
    }
  }

  // 未溢出但高度实际需要值超过槽高 → 同样视为溢出，高度用实际需要值（向下延伸）
  const totalH = rowH.reduce((a, b) => a + b, 0);
  if (!overflow && slotHeightIn > 0 && totalH > slotHeightIn + 1e-9) overflow = true;

  return {
    rows,
    cols,
    colWidths: colW,
    rowHeights: rowH,
    fontSize,
    cells: norm,
    overflow,
    scaleRatio: fontSize / baseFontSizePt,
    baseFontSizePt
  };
}

// 2D 数组 → FitCell（无合并）
export function cellsFromGrid(grid: string[][]): FitCell[] {
  const cells: FitCell[] = [];
  for (let r = 0; r < (grid || []).length; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      cells.push({ r, c, rowspan: 1, colspan: 1, text: String(row[c] ?? '') });
    }
  }
  return cells;
}

// 展开网格：主格在左上；被覆盖位置返回所属主格（便于渲染区分「被覆盖」与「真空位」），
// 无任何主格覆盖的空位为 null。（注：若按「被覆盖位置 = null」的字面语义，渲染端将无法区分
// 「被覆盖」与「真空位」，合并计划会失真，故被覆盖位保留主格引用，见报告。）
export function expandGrid(cells: FitCell[], rows: number, cols: number): (FitCell | null)[][] {
  const grid: (FitCell | null)[][] = Array.from(
    { length: Math.max(0, rows) },
    () => new Array<FitCell | null>(Math.max(0, cols)).fill(null)
  );
  for (const cell of cells || []) {
    if (!cell || !Number.isFinite(cell.r) || !Number.isFinite(cell.c) || cell.r < 0 || cell.c < 0) continue;
    const rs = Math.max(1, Math.floor(cell.rowspan || 1));
    const cs = Math.max(1, Math.floor(cell.colspan || 1));
    for (let i = 0; i < rs && cell.r + i < rows; i++) {
      for (let j = 0; j < cs && cell.c + j < cols; j++) {
        if (!grid[cell.r + i][cell.c + j]) grid[cell.r + i][cell.c + j] = cell;
      }
    }
  }
  return grid;
}

// 「相同字样一键合并」（阶段 3）：trim 后完全一致、连续 ≥2 个才合并；
// 只合并简单格（rowspan=1 且 colspan=1），空格/已有合并格作为断点；
// 横向 = 逐行扫描相邻列；纵向 = 逐列扫描相邻行；auto = 先横后纵。合并后 r/c 取主格（左上）位置。
export function mergeSameTextCells(cells: FitCell[], dir: 'vertical' | 'horizontal' | 'auto'): FitCell[] {
  let out = normalizeCells(cells);
  if (dir === 'horizontal' || dir === 'auto') out = mergeHorizontal(out);
  if (dir === 'vertical' || dir === 'auto') out = mergeVertical(out);
  return out;
}

function mergeHorizontal(cells: FitCell[]): FitCell[] {
  const rows = cells.length ? Math.max(...cells.map((c) => c.r + 1)) : 0;
  const result: FitCell[] = [];
  for (let r = 0; r < rows; r++) {
    const rowCells = cells.filter((c) => c.r === r).sort((a, b) => a.c - b.c);
    let i = 0;
    while (i < rowCells.length) {
      const cur = rowCells[i];
      const simple = cur.rowspan === 1 && cur.colspan === 1;
      const key = cur.text.trim();
      if (!simple || !key) { result.push(cur); i++; continue; }
      let j = i + 1;
      while (j < rowCells.length) {
        const nxt = rowCells[j];
        if (nxt.rowspan !== 1 || nxt.colspan !== 1 || nxt.text.trim() !== key) break;
        j++;
      }
      if (j - i >= 2) {
        result.push({ ...cur, colspan: j - i });
        i = j;
      } else {
        result.push(cur);
        i++;
      }
    }
  }
  return result;
}

function mergeVertical(cells: FitCell[]): FitCell[] {
  const cols = cells.length ? Math.max(...cells.map((c) => c.c + 1)) : 0;
  const result: FitCell[] = [];
  for (let c = 0; c < cols; c++) {
    const colCells = cells.filter((x) => x.c === c).sort((a, b) => a.r - b.r);
    let i = 0;
    while (i < colCells.length) {
      const cur = colCells[i];
      const simple = cur.rowspan === 1 && cur.colspan === 1;
      const key = cur.text.trim();
      if (!simple || !key) { result.push(cur); i++; continue; }
      let j = i + 1;
      while (j < colCells.length) {
        const nxt = colCells[j];
        if (nxt.rowspan !== 1 || nxt.colspan !== 1 || nxt.text.trim() !== key) break;
        j++;
      }
      if (j - i >= 2) {
        result.push({ ...cur, rowspan: j - i });
        i = j;
      } else {
        result.push(cur);
        i++;
      }
    }
  }
  return result;
}
