// 小工具 · 文本处理与表格尺寸计算单测（Node 24 type-stripping 直接 require .ts）
const { test } = require('node:test');
const assert = require('node:assert');

const {
  removeAllSpaces,
  collapseSpaces,
  indentParagraphs,
  removeEmptyParagraphs,
  textChanged,
  indentLines,
  removeEmptyLines,
  normalizeSep,
  detectSep,
  restoreSep,
  separateParagraphs,
  locateTitleRange,
  locateBodyRange,
  splitSegs,
  splitParagraphs,
} = require('../../addin/src/tools/textOps.ts');

const {
  formatIndexForTargetParagraph,
} = require('../../addin/src/tools/formatMapping.ts');

const {
  estimateTextWidthEm,
  estimateColumnWidths,
  estimateRowHeights,
  evenColumnWidths,
  evenRowHeights,
} = require('../../addin/src/tools/tableOps.ts');

// ---------- textOps：去空格 ----------
test('removeAllSpaces：去除半角/全角/不间断空格，保留换行', () => {
  assert.equal(removeAllSpaces(' 你好 世界 '), '你好世界');
  assert.equal(removeAllSpaces('a b c'), 'abc');
  assert.equal(removeAllSpaces('第一行 内容\n第二行 内容'), '第一行内容\n第二行内容');
  assert.equal(removeAllSpaces('\u3000全角\u3000空格\u00A0'), '全角空格');
});

test('collapseSpaces：连续空格合并为单个，去行首尾', () => {
  assert.equal(collapseSpaces('a   b   c'), 'a b c');
  assert.equal(collapseSpaces('  开头 结尾  '), '开头 结尾');
  assert.equal(collapseSpaces('行1  有空格\n行2  有空格'), '行1 有空格\n行2 有空格');
});

// ---------- textOps：段首缩进 ----------
test('indentParagraphs：每段段首加两个全角空格，已有缩进跳过', () => {
  assert.equal(indentParagraphs('第一段\n第二段'), '\u3000\u3000第一段\n\u3000\u3000第二段');
  assert.equal(indentParagraphs('\u3000\u3000已缩进\n未缩进'), '\u3000\u3000已缩进\n\u3000\u3000未缩进');
  assert.equal(indentParagraphs('有内容\n\n有内容'), '\u3000\u3000有内容\n\n\u3000\u3000有内容');
});

test('indentParagraphs：可改为每段段首加一个半角空格，已有缩进跳过', () => {
  assert.equal(indentParagraphs('第一段\n第二段', ' '), ' 第一段\n 第二段');
  assert.equal(indentParagraphs(' 已缩进\n未缩进', ' '), ' 已缩进\n 未缩进');
});

// ---------- textOps：删空行空段 ----------
test('removeEmptyParagraphs：删除空段落，压缩连续换行', () => {
  assert.equal(removeEmptyParagraphs('第一段\n\n第二段'), '第一段\n第二段');
  assert.equal(removeEmptyParagraphs('\n第一段\n\n\n第二段\n'), '第一段\n第二段');
  assert.equal(removeEmptyParagraphs('只有空白\n   \n内容'), '只有空白\n内容');
});

test('textChanged：判断文本是否变化', () => {
  assert.equal(textChanged('a', 'a'), false);
  assert.equal(textChanged('a', 'b'), true);
});

// ---------- textOps：段落辅助 ----------
test('indentLines / removeEmptyLines：selection 内段落辅助', () => {
  assert.equal(indentLines('第一段\n第二段'), ' 第一段\n 第二段');
  assert.equal(indentLines(' 已缩进\n未缩进'), ' 已缩进\n 未缩进');
  assert.equal(removeEmptyLines('第一段\n\n第二段\n'), '第一段\n第二段');
});

test('formatIndexForTargetParagraph：目标段落更多时沿用最后一段格式，不循环', () => {
  const indexes = [0, 1, 2, 3, 4, 5].map((i) => formatIndexForTargetParagraph(i, 3));
  assert.deepEqual(indexes, [0, 1, 2, 2, 2, 2]);
});

test('normalizeSep / detectSep / restoreSep：\r\n 与 \r 归一化并还原', () => {
  assert.equal(normalizeSep('a\r\nb'), 'a\nb');
  assert.equal(normalizeSep('a\rb'), 'a\nb');
  assert.equal(detectSep('a\rb'), '\r');
  assert.equal(detectSep('a\nb'), '\n');
  assert.equal(restoreSep('a\nb', '\r'), 'a\rb');
  assert.equal(restoreSep('a\nb', '\n'), 'a\nb');
});

// ---------- tableOps：字符宽度估算 ----------
test('estimateTextWidthEm：中文 1em，半角约 0.55em，空格 0.3em', () => {
  assert.ok(estimateTextWidthEm('中文') >= 1.9 && estimateTextWidthEm('中文') <= 2.1);
  assert.ok(estimateTextWidthEm('abc') > 1.6 && estimateTextWidthEm('abc') < 1.7);
  assert.ok(estimateTextWidthEm('a b') > estimateTextWidthEm('ab'));
});

// ---------- tableOps：列宽估算 ----------
test('estimateColumnWidths：内容宽按比例分配且不超页面可用宽', () => {
  const values = [
    ['序号', '项目名称', '说明'],
    ['1', '年度经营分析', '这是一个较长的说明文字'],
    ['2', '预算', '短'],
  ];
  const widths = estimateColumnWidths(values, 600, 900);
  assert.equal(widths.length, 3);
  const total = widths.reduce((a, b) => a + b, 0);
  assert.ok(total <= 900 + 1, '总宽不超页面可用宽，实际 ' + total);
  assert.ok(widths[2] > widths[0], '说明列应比序号列宽');
  widths.forEach((w) => assert.ok(w >= 30, '列宽 ' + w + ' 过小'));
});

test('estimateColumnWidths：空表格返回空数组', () => {
  assert.deepEqual(estimateColumnWidths([], 600, 900), []);
});

// ---------- tableOps：行高估算 ----------
test('estimateRowHeights：行高随内容增长且不低于基准', () => {
  const values = [
    ['表头1', '表头2'],
    ['内容', '这是一段很长的内容需要折行显示在这里'],
  ];
  const widths = [300, 300];
  const heights = estimateRowHeights(values, widths);
  assert.equal(heights.length, 2);
  assert.ok(heights[1] >= heights[0], '长内容行应不低于表头行');
  heights.forEach((h) => assert.ok(h >= 10, '行高 ' + h + ' 过小'));
});

// ---------- tableOps：均分 ----------
test('evenColumnWidths / evenRowHeights：等宽等高', () => {
  assert.deepEqual(evenColumnWidths(4, 800), [200, 200, 200, 200]);
  assert.deepEqual(evenRowHeights(3, 300), [100, 100, 100]);
  assert.deepEqual(evenColumnWidths(0, 800), []);
});
// ---------- textOps：分隔每段 ----------
test('separateParagraphs：段落间插空行，已有空行不重复', () => {
  // 两段 → 中间插一个空段
  assert.equal(separateParagraphs('第一段\n第二段'), '第一段\n\n第二段');
  // 三段连续 → 每段之间一个空行
  assert.equal(separateParagraphs('A\nB\nC'), 'A\n\nB\n\nC');
  // 已有空行 → 不重复插入
  assert.equal(separateParagraphs('第一段\n\n第二段'), '第一段\n\n第二段');
  // 首尾空段不额外补
  assert.equal(separateParagraphs('\n第一段\n'), '第一段');
});

// ---------- textOps：定位标题 / 正文范围（选中语义） ----------
test('locateTitleRange：取首段偏移，长首段截到句末标点', () => {
  // 首段即标题
  assert.deepEqual(locateTitleRange('季度业绩汇报\n这是正文内容。'), { start: 0, length: 6 });
  // 前导空行：偏移包含空行与分隔符
  assert.deepEqual(locateTitleRange('\n\n季度业绩汇报\n正文'), { start: 2, length: 6 });
  // 长首段（>40 字）：截取到第一个句末标点
  const longHead = '这是本季度最详细的一份经营分析报告，覆盖营收、利润、成本与现金流四个维度，重点说明海外市场的增长情况。后面还有更多内容';
  const r = locateTitleRange(longHead);
  assert.equal(r.length, '这是本季度最详细的一份经营分析报告，覆盖营收、利润、成本与现金流四个维度，重点说明海外市场的增长情况。'.length);
  assert.equal(longHead.slice(r.start, r.start + r.length), '这是本季度最详细的一份经营分析报告，覆盖营收、利润、成本与现金流四个维度，重点说明海外市场的增长情况。');
  // 无标点长文本：返回完整首段
  const longNoPunct = '无标点'.repeat(20);
  const r2 = locateTitleRange(longNoPunct);
  assert.deepEqual(r2, { start: 0, length: longNoPunct.length });
  // 空文本
  assert.equal(locateTitleRange(''), null);
});

test('locateTitleRange：\r\n 分隔偏移正确（2 字符分隔）', () => {
  const text = '标题行\r\n正文第一行\r\n正文第二行';
  const r = locateTitleRange(text);
  assert.equal(r.start, 0);
  assert.equal(r.length, 3);
});

test('locateBodyRange：标题之后到末尾', () => {
  const text = '季度业绩汇报\n这是正文第一段。\n这是正文第二段。';
  const r = locateBodyRange(text);
  assert.equal(text.slice(r.start, r.start + r.length), '这是正文第一段。\n这是正文第二段。');
  // 只有标题无正文 → length 0（调用方提示未识别）
  const r2 = locateBodyRange('只有标题');
  assert.equal(r2.length, 0);
});
// ---------- textOps：splitParagraphs（段落切分，兼容所有分隔符） ----------
test('splitParagraphs：\n 分隔切分正确', () => {
  assert.deepEqual(splitParagraphs('a\nb\nc'), [{ start: 0, len: 1 }, { start: 2, len: 1 }, { start: 4, len: 1 }]);
  assert.deepEqual(splitParagraphs('abc'), [{ start: 0, len: 3 }]);
  assert.deepEqual(splitParagraphs(''), [{ start: 0, len: 0 }]);
});

test('splitParagraphs：\r 分隔切分正确（含空段）', () => {
  assert.deepEqual(splitParagraphs('a\r\rb'), [{ start: 0, len: 1 }, { start: 2, len: 0 }, { start: 3, len: 1 }]);
});

test('splitParagraphs：\r\n 分隔偏移正确（2 字符分隔符）', () => {
  assert.deepEqual(splitParagraphs('标题\r\n正文1\r\n正文2'), [
    { start: 0, len: 2 },
    { start: 4, len: 3 },
    { start: 9, len: 3 },
  ]);
});

test('splitSegs：兼容别名（忽略 sep 参数）', () => {
  assert.deepEqual(splitSegs('a\rb', '\r'), [{ start: 0, len: 1 }, { start: 2, len: 1 }]);
  assert.deepEqual(splitSegs('a\nb', '\n'), [{ start: 0, len: 1 }, { start: 2, len: 1 }]);
});
