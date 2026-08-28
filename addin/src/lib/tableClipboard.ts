// 剪贴板粘贴解析（阶段 0.2：粘贴优先 HTML，colspan/rowspan + 文字，样式全丢）
// 纯 TS、无 DOM 依赖（字符串/正则解析），Node 24 可直接 require 单测。
// 输出统一 FitCell 结构（r/c 从 0 起连续编号），与排版引擎 tableModel.ts 对接。
import type { FitCell } from './tableModel.js';

// ---------- HTML 表格解析 ----------

// 解析 <table>…</table>（含 <thead>/<tbody> 包裹）：
//  - 逐 <tr> 取 <td|th …>…</td|th>，提取 colspan/rowspan（数字，默认 1）
//  - 单元格文字去标签（<br>/<br/> → 换行，其余标签全丢，实体按 XML/HTML 解码）
//  - 丢弃一切样式（class/style/颜色/字体）
//  - 空行（所有单元格 trim 后为空）跳过；健壮：Excel/Word/网页复制出来的 HTML 都长这样
export function parseTableHtml(html: string): FitCell[] {
  const cells: FitCell[] = [];
  const tableMatch = (html || '').match(/<table\b[^>]*>([\s\S]*?)<\/table>/i);
  if (!tableMatch) return [];
  const tableInner = tableMatch[1];
  const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let trM;
  let r = 0;
  const coveredCols: number[] = []; // rowspan 感知：每列还需被占用的剩余行数
  while ((trM = trRe.exec(tableInner))) {
    const rowCells: { colspan: number; rowspan: number; text: string }[] = [];
    const tcRe = /<t[dh]\b([^>]*)>([\s\S]*?)<\/t[dh]>/gi;
    let tcM;
    while ((tcM = tcRe.exec(trM[1]))) {
      const attrs = tcM[1] || '';
      rowCells.push({
        colspan: numAttr(attrs, 'colspan'),
        rowspan: numAttr(attrs, 'rowspan'),
        text: cellText(tcM[2] || '')
      });
    }
    if (!rowCells.length) continue;                 // 无单元格的 tr 跳过
    if (rowCells.every((c) => c.text.trim() === '')) continue; // 空行跳过
    // rowspan 感知的列定位：本行单元格的列号必须跳过「被上行 rowspan 覆盖」的列，
    // 否则后续行的 td 会错位、与合并格重叠（覆盖格文字会被 expandGrid 丢弃）。
    // coveredCols[c] = 该列还需被 rowspan 占用的剩余行数。
    let c = 0;
    for (const rc of rowCells) {
      while (c < coveredCols.length && (coveredCols[c] || 0) > 0) c++; // 跳过被覆盖列
      cells.push({ r, c, rowspan: rc.rowspan, colspan: rc.colspan, text: rc.text });
      if (rc.rowspan > 1) {
        // 存 rowspan（而非 rowspan-1）：行末统一递减一次 = 本行消耗一行，剩余行数留给后续行判断
        for (let j = 0; j < rc.colspan; j++) coveredCols[c + j] = rc.rowspan;
      }
      c += rc.colspan;
    }
    r++;
    for (let i = 0; i < coveredCols.length; i++) if (coveredCols[i] > 0) coveredCols[i]--;
  }
  return cells;
}

// 属性值提取（数字，缺省 1；colspan="2" / colspan=2 均支持）
function numAttr(attrs: string, name: string): number {
  const m = attrs.match(new RegExp(name + '\\s*=\\s*["\']?(\\d+)', 'i'));
  return m ? Math.max(1, Number(m[1]) || 1) : 1;
}

// 单元格文字：<br> → 换行；其余标签全丢（含内联样式标签属性）；实体解码；整体 trim
function cellText(inner: string): string {
  let s = inner;
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = decodeEntities(s);
  return s.trim();
}

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ensp: ' ', emsp: ' ', thinsp: ' ',
  copy: '©', reg: '®', trade: '™', times: '×', divide: '÷', mdash: '—', ndash: '–', hellip: '…',
  bull: '•', middot: '·', laquo: '«', raquo: '»', ldquo: '“', rdquo: '”', lsquo: '‘', rsquo: '’',
  szlig: 'ß', deg: '°', plusmn: '±', sup2: '²', sup3: '³', frac12: '½', frac14: '¼', frac34: '¾'
};

// XML/HTML 实体解码：命名实体表 + 十进制/十六进制数值实体
function decodeEntities(s: string): string {
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, ent: string) => {
    const lower = ent.toLowerCase();
    let cp = NaN;
    if (lower.startsWith('#x')) cp = parseInt(ent.slice(2), 16);
    else if (lower.startsWith('#')) cp = parseInt(ent.slice(1), 10);
    else return NAMED_ENTITIES[ent] ?? m;
    return Number.isFinite(cp) && cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
  });
}

// ---------- CSV / TSV 解析 ----------

// 行分隔：每行按分隔符分列：优先制表符（Excel 复制），其次逗号（引号内逗号不切，
// 双引号转义 "" → "），再次分号（欧洲区域 Excel），最后整行一格（与向导 parseTableClipboard 行为一致）。
// 空行跳过；每格 trim。
export function parseTableCsv(text: string): string[][] {
  const lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const out: string[][] = [];
  for (const line of lines) {
    if (!line.trim()) continue;
    if (line.includes('\t')) {
      out.push(line.split('\t').map((x) => x.trim()));
    } else if (line.includes(',')) {
      out.push(parseCsvLine(line));
    } else if (line.includes(';')) {
      out.push(line.split(';').map((x) => x.trim()));
    } else {
      out.push([line.trim()]);
    }
  }
  return out;
}

// 单行 CSV 解析：引号包裹的字段可含逗号/引号（"" 转义）
function parseCsvLine(line: string): string[] {
  const cells: string[] = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += ch;
    } else if (ch === '"') {
      inQ = true;
    } else if (ch === ',') {
      cells.push(cur.trim());
      cur = '';
    } else cur += ch;
  }
  cells.push(cur.trim());
  return cells;
}
