// 表格自动排版引擎（阶段 2/3）单测：排版引擎 + 粘贴解析 + 一键合并
// Node 24 type-stripping 直接 require addin 的 TS 纯函数模块
const { test } = require('node:test');
const assert = require('node:assert');

const {
  fitTableLayout,
  cellsFromGrid,
  expandGrid,
  mergeSameTextCells,
} = require('../../addin/src/lib/tableModel.ts');

const {
  parseTableHtml,
  parseTableCsv,
} = require('../../addin/src/lib/tableClipboard.ts');

const PT = 1 / 72;

// ---------- fitTableLayout：基础 ----------

test('fitTableLayout 基础：3×2 表放 4in×1.25in 槽，内容不长 → 全表 14pt、ΣcolW=槽宽、行高在槽内', () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['季度', '营收', '利润'],
    ['1', '2', '3']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.equal(fit.fontSize, 14, '未缩字 → 基准字号');
  assert.equal(fit.rows, 2);
  assert.equal(fit.cols, 3);
  const sum = fit.colWidths.reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(sum - 4) < 1e-6, '列宽总和 = 槽宽 4in，实际 ' + sum);
  fit.colWidths.forEach((w) => assert.ok(w > 0, '列宽为正，实际 ' + w));
  const sumH = fit.rowHeights.reduce((a, b) => a + b, 0);
  assert.ok(sumH <= 1.25 + 1e-6, '行高总和 ≤ 槽高 1.25in，实际 ' + sumH);
  assert.equal(fit.overflow, false);
  assert.equal(fit.scaleRatio, 1);
});

// ---------- fitTableLayout：缩字 ----------

test('fitTableLayout 缩字：超长不可断文本 → 字号 <14 且 ≥10，scaleRatio<1，不溢出', () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['标题'],
    ['8888888888888888888888888888888888888888'] // 40 位数字：min_w ≈ 4.48in > 4in
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.ok(fit.fontSize < 14 && fit.fontSize >= 10, '字号缩到 [10,14)，实际 ' + fit.fontSize);
  assert.ok(fit.scaleRatio < 1 && fit.scaleRatio > 0, 'scaleRatio < 1，实际 ' + fit.scaleRatio);
  assert.equal(fit.overflow, false, '缩字即可放下 → 不溢出');
  assert.ok(Math.abs(fit.colWidths.reduce((a, b) => a + b, 0) - 4) < 1e-6, '列宽总和仍 = 槽宽');
});

// ---------- fitTableLayout：超下限（宽度锁槽 + 向下延伸） ----------

test('fitTableLayout 超下限：更长不可断文本 → 字号停 10、overflow=true、ΣcolW=槽宽、ΣrowH>槽高', () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['标题'],
    ['9'.repeat(200)], // 200 位数字：min_w ≈ 21.4in，缩到 10pt 仍远超 4in
    ['行3']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.equal(fit.fontSize, 10, '缩到下限 10pt');
  assert.equal(fit.overflow, true, '宽度锁槽宽 → overflow');
  assert.ok(Math.abs(fit.colWidths.reduce((a, b) => a + b, 0) - 4) < 1e-6, '宽度锁槽宽：ΣcolW 仍 = 槽宽 4in');
  const sumH = fit.rowHeights.reduce((a, b) => a + b, 0);
  assert.ok(sumH > 1.25, '行高向下延伸：ΣrowH > 槽高 1.25in，实际 ' + sumH);
});

// ---------- fitTableLayout：富余分配 ----------

test('fitTableLayout 富余分配：两列内容差异大 → 长文本列拿更多 slack', () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['x', '这是很长的一段文字需要更多宽度']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.ok(Math.abs(fit.colWidths.reduce((a, b) => a + b, 0) - 4) < 1e-6, '总宽 = 槽宽');
  assert.ok(fit.colWidths[1] > fit.colWidths[0] * 2, '长文本列显著更宽：' + fit.colWidths[0] + ' vs ' + fit.colWidths[1]);
  assert.equal(fit.fontSize, 14, '富余时字号不缩');
});

// ---------- fitTableLayout：colspan ----------

test('fitTableLayout colspan：跨 2 列长格 → 所跨列之和 ≥ 该格 min_w', () => {
  const cells = [
    { r: 0, c: 0, rowspan: 1, colspan: 2, text: '12345678901234567890' }, // 20 位数字
    { r: 1, c: 0, rowspan: 1, colspan: 1, text: 'A' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: 'B' }
  ];
  const fit = fitTableLayout(cells, { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.equal(fit.cols, 2);
  // 该格 min_w = (20×0.55em×14pt + 左右 padding 0.2in) 转英寸
  const expectedMin = (20 * 0.55 * 14) * PT + 0.2;
  const spannedSum = fit.colWidths[0] + fit.colWidths[1];
  assert.ok(spannedSum >= expectedMin - 0.01, '所跨列之和 ' + spannedSum + ' ≥ min_w ' + expectedMin);
});

// ---------- fitTableLayout：全表统一字号 ----------

test('fitTableLayout 全表统一字号：FitResult.fontSize 单一值', () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['季度', '营收'],
    ['增长', '18.5%']
  ]), { slotWidthIn: 3, slotHeightIn: 1 });
  assert.equal(typeof fit.fontSize, 'number');
  assert.ok(fit.fontSize >= 10 && fit.fontSize <= 14, '字号在 [10,14] 内，实际 ' + fit.fontSize);
  assert.ok(Number.isFinite(fit.fontSize));
  // 所有主格渲染时使用同一字号（引擎单一 fontSize 字段即保证）
  assert.equal(fit.cells.length, 4);
});

// ---------- mergeSameTextCells ----------

test('mergeSameTextCells 横向：相邻列同文字 ≥2 合并为 colspan', () => {
  const cells = [
    { r: 0, c: 0, rowspan: 1, colspan: 1, text: 'a' },
    { r: 0, c: 1, rowspan: 1, colspan: 1, text: 'a' },
    { r: 0, c: 2, rowspan: 1, colspan: 1, text: 'b' },
    { r: 1, c: 0, rowspan: 1, colspan: 1, text: 'c' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: 'c' },
    { r: 1, c: 2, rowspan: 1, colspan: 1, text: 'c' }
  ];
  const out = mergeSameTextCells(cells, 'horizontal');
  assert.equal(out.length, 3, 'a+a 合成一个、b 单个、c+c+c 合成一个');
  const m0 = out.find((x) => x.r === 0 && x.c === 0);
  assert.equal(m0.colspan, 2);
  assert.equal(m0.text, 'a');
  const m1 = out.find((x) => x.r === 1 && x.c === 0);
  assert.equal(m1.colspan, 3);
  assert.equal(m1.text, 'c');
});

test('mergeSameTextCells 纵向：相邻行同文字 ≥2 合并为 rowspan', () => {
  const cells = [
    { r: 0, c: 0, rowspan: 1, colspan: 1, text: 'x' },
    { r: 1, c: 0, rowspan: 1, colspan: 1, text: 'x' },
    { r: 2, c: 0, rowspan: 1, colspan: 1, text: 'y' },
    { r: 0, c: 1, rowspan: 1, colspan: 1, text: 'z' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: 'w' }
  ];
  const out = mergeSameTextCells(cells, 'vertical');
  assert.equal(out.length, 4, 'x+x 合成一个；z/w 不同不合并');
  const mx = out.find((x) => x.r === 0 && x.c === 0);
  assert.equal(mx.rowspan, 2);
  assert.equal(out.find((x) => x.c === 1 && x.r === 0).rowspan, 1);
});

test('mergeSameTextCells auto：先横后纵（已有合并格作为纵向断点）', () => {
  const cells = [
    { r: 0, c: 0, rowspan: 1, colspan: 1, text: '同' },
    { r: 0, c: 1, rowspan: 1, colspan: 1, text: '同' },
    { r: 1, c: 0, rowspan: 1, colspan: 1, text: '同' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: '同' }
  ];
  const out = mergeSameTextCells(cells, 'auto');
  assert.equal(out.length, 2, '每行合并为 colspan2，纵向跨已合并格不合并');
  out.forEach((c) => assert.equal(c.colspan, 2));
  assert.equal(out[0].r, 0);
  assert.equal(out[1].r, 1);
});

test('mergeSameTextCells：连续 ≥2 才合并（单个相同不合并）；trim 后比较', () => {
  // 单个相同文字不合并
  const c1 = mergeSameTextCells([
    { r: 0, c: 0, rowspan: 1, colspan: 1, text: 'a' },
    { r: 0, c: 1, rowspan: 1, colspan: 1, text: 'b' }
  ], 'horizontal');
  assert.equal(c1.length, 2, '不连续相同 → 不合并');
  // trim 后比较（'  a  ' 与 'a' 视为相同）
  const c2 = mergeSameTextCells([
    { r: 0, c: 0, rowspan: 1, colspan: 1, text: '  a  ' },
    { r: 0, c: 1, rowspan: 1, colspan: 1, text: 'a' },
    { r: 0, c: 2, rowspan: 1, colspan: 1, text: 'b' }
  ], 'horizontal');
  assert.equal(c2.length, 2);
  assert.equal(c2[0].colspan, 2);
  assert.equal(c2[0].text, '  a  ', '保留最左格文字（原样）');
});

// ---------- parseTableCsv ----------

test('parseTableCsv：TSV（Excel 复制）', () => {
  assert.deepEqual(parseTableCsv('A\tB\tC\n1\t2\t3'), [['A', 'B', 'C'], ['1', '2', '3']]);
  assert.deepEqual(parseTableCsv(' 甲 \t 乙 \r\n 1 \t 2 '), [['甲', '乙'], ['1', '2']], '\r\n 与 trim');
  assert.deepEqual(parseTableCsv('只有一列\n'), [['只有一列']], '单列整行一格');
});

test('parseTableCsv：引号包裹字段含逗号', () => {
  const out = parseTableCsv('"项目名,含逗号",数量\n"甲","2"');
  assert.deepEqual(out, [['项目名,含逗号', '数量'], ['甲', '2']]);
  // 转义双引号 ""
  const out2 = parseTableCsv('"他说""好""",x');
  assert.deepEqual(out2, [['他说"好"', 'x']]);
});

// ---------- parseTableHtml ----------

test('parseTableHtml：colspan/rowspan/样式/嵌套标签/<br> 换行', () => {
  const html = '<table border="1" style="width:100%">' +
    '<thead><tr><th colspan="2" style="background:red">季度<b>指标</b></th><th>单位</th></tr></thead>' +
    '<tbody><tr><td>营收</td><td>18.5%</td><td>万元<br/>(万)</td></tr></tbody>' +
    '</table>';
  const cells = parseTableHtml(html);
  assert.equal(cells.length, 5, '表头行 2 格 + 数据行 3 格');
  const c00 = cells.find((c) => c.r === 0 && c.c === 0);
  assert.equal(c00.colspan, 2, 'th colspan=2');
  assert.equal(c00.rowspan, 1);
  assert.equal(c00.text, '季度指标', '嵌套 <b> 标签文字并入、样式丢弃');
  assert.ok(!c00.text.includes('background'), '样式不进文字');
  const c01 = cells.find((c) => c.r === 0 && c.c === 2);
  assert.equal(c01.rowspan, 1, 'th rowspan=1（rowspan 覆盖语义由专门的回归测试覆盖）');
  assert.equal(c01.text, '单位', 'colspan 后单位列正确落在 c=2');
  const c12 = cells.find((c) => c.r === 1 && c.c === 2);
  assert.equal(c12.text, '万元\n(万)', '<br/> → 换行');
});

test('parseTableHtml：实体解码与空行跳过', () => {
  const html = '<table><tr><td>A&amp;B</td><td>x&nbsp;y</td><td>&#169; &#x4E2D;</td></tr></table>';
  const cells = parseTableHtml(html);
  assert.equal(cells.length, 3, '1 行 3 格');
  assert.equal(cells[0].text, 'A&B');
  assert.equal(cells[1].text, 'x y', '&nbsp; → 空格');
  assert.equal(cells[2].text, '© 中', '数值实体（十进制 + 十六进制）');
  // 空行跳过、无表格返回空
  assert.deepEqual(parseTableHtml('<table><tr><td> </td><td></td></tr></table>'), []);
  assert.deepEqual(parseTableHtml('没有表格'), []);
});

// ---------- expandGrid ----------

test('expandGrid：3×3 含 2×2 合并 → 展开网格正确（被覆盖位指向主格，真空位 null）', () => {
  const cells = [
    { r: 0, c: 0, rowspan: 2, colspan: 2, text: 'M' },
    { r: 0, c: 2, rowspan: 1, colspan: 1, text: 'A' },
    { r: 2, c: 0, rowspan: 1, colspan: 1, text: 'B' }
  ];
  const grid = expandGrid(cells, 3, 3);
  assert.equal(grid.length, 3);
  assert.equal(grid[0].length, 3);
  // 主格位置 = 主格本身
  assert.equal(grid[0][0], cells[0]);
  // 被覆盖位置 → 指向所属主格（渲染端据此区分「被覆盖」与「真空位」）
  assert.equal(grid[0][1], cells[0]);
  assert.equal(grid[1][0], cells[0]);
  assert.equal(grid[1][1], cells[0]);
  // 普通格
  assert.equal(grid[0][2], cells[1]);
  assert.equal(grid[2][0], cells[2]);
  // 真空位 = null
  assert.equal(grid[1][2], null);
  assert.equal(grid[2][1], null);
  assert.equal(grid[2][2], null);
});

// 回归：rowspan 感知的列定位 —— 上行 rowspan 覆盖的列，后续行的 td 必须跳过，
// 否则单元格错位与合并格重叠（被覆盖格文字会丢失）
test('parseTableHtml：rowspan 后行跳过被覆盖列（22.3% 落到 c=1 而非 c=0）', () => {
  const html = '<table><thead><tr><th colspan="2">营业收入</th><th>单位</th></tr></thead>' +
    '<tbody><tr><td rowspan="2">华东</td><td>18.5%</td><td>万元</td></tr>' +
    '<tr><td>22.3%</td><td>万元</td></tr></tbody></table>';
  const cells = parseTableHtml(html);
  const find = (t) => cells.find((c) => c.text === t);
  assert.equal(find('营业收入').c, 0);
  assert.equal(find('营业收入').colspan, 2);
  assert.equal(find('华东').r, 1); assert.equal(find('华东').c, 0); assert.equal(find('华东').rowspan, 2);
  assert.equal(find('18.5%').r, 1); assert.equal(find('18.5%').c, 1);
  // rowspan 后一行：c 从 1 开始（跳过 0 列的华东），22.3% 在 (2,1)
  assert.equal(find('22.3%').r, 2); assert.equal(find('22.3%').c, 1);
  const wan = cells.filter((c) => c.text === '万元');
  assert.equal(wan.length, 2, '两处万元（r1c2 与 r2c2）');
  assert.equal(wan[0].r, 1); assert.equal(wan[0].c, 2);
  assert.equal(wan[1].r, 2); assert.equal(wan[1].c, 2);
  // 没有两个格占用同一位置（22.3% 不再与华东的 rowspan 重叠）
  const positions = cells.map((c) => c.r + ':' + c.c);
  assert.equal(new Set(positions).size, positions.length, '无重叠位置');
});

test('parseTableHtml：连续两个 rowspan（相邻列）也正确', () => {
  const html = '<table><tr><td rowspan="2">A</td><td rowspan="2">B</td><td>1</td></tr>' +
    '<tr><td>2</td></tr></table>';
  const cells = parseTableHtml(html);
  const find = (t) => cells.find((c) => c.text === t);
  assert.equal(find('A').c, 0); assert.equal(find('B').c, 1); assert.equal(find('1').c, 2);
  // 第二行：A、B 两列都被覆盖 → 唯一真实格 2 落在 c=2
  const row2 = cells.filter((c) => c.r === 1);
  assert.equal(row2.length, 1);
  assert.equal(row2[0].text, '2');
  assert.equal(row2[0].c, 2);
});

test('parseTableHtml：rowspan + colspan 组合', () => {
  const html = '<table><tr><td colspan="2" rowspan="2">X</td><td>Y</td></tr>' +
    '<tr><td>Z</td></tr></table>';
  const cells = parseTableHtml(html);
  const find = (t) => cells.find((c) => c.text === t);
  assert.equal(find('X').colspan, 2); assert.equal(find('X').rowspan, 2); assert.equal(find('X').c, 0);
  assert.equal(find('Y').c, 2);
  // 第二行：列 0、1 都被 X 覆盖 → Z 落在 c=2
  const row2 = cells.filter((c) => c.r === 1);
  assert.equal(row2.length, 1);
  assert.equal(row2[0].text, 'Z');
  assert.equal(row2[0].c, 2);
});
