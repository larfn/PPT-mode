// 前端规则分类器单测（Node 24 type-stripping 直接 require .ts）
const { test } = require('node:test');
const assert = require('node:assert');
const { analyzeShapesByRules, HIGH_CONFIDENCE } = require('../../addin/src/lib/analyze.ts');

const PAGE = { width: 13.33, height: 7.5 };

function shp(over) {
  return {
    id: 'x', type: 'text', name: 'x', source: 'slide',
    hasText: true, text: '内容文字',
    bounds: { left: 1, top: 1, width: 5, height: 1 },
    textStyle: { size: 18 },
    ...over
  };
}
function find(recs, idx) { return recs.find((r) => r.idx === idx); }

test('纯文本模板：最大字号短文本 → title；标题下方次级文本 → subtitle', () => {
  const shapes = [
    shp({ text: '季度业绩汇报', textStyle: { size: 32 }, bounds: { left: 1, top: 0.5, width: 10, height: 1 } }),
    shp({ text: '这是本季度的详细业绩分析，营收增长显著，利润创新高，海外市场表现突出……', textStyle: { size: 14 }, bounds: { left: 1, top: 2, width: 10, height: 3 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedRole, 'ai_text');
  assert.equal(find(recs, 0).recommendedSemanticRole, 'title');
  assert.ok(find(recs, 0).confidence >= 0.7);
  // 14pt 文本位于主标题下方且字号次之 → 按当前 8 类语义模型识别为副标题。
  assert.equal(find(recs, 1).recommendedRole, 'ai_text');
  assert.equal(find(recs, 1).recommendedSemanticRole, 'subtitle');
  assert.equal(find(recs, 1).ruleId, 'R-SUB-UNDER');
  assert.ok(find(recs, 1).confidence >= HIGH_CONFIDENCE);
});

test('图文模板：大图 → ai_image；正文 → body', () => {
  const shapes = [
    shp({ type: 'picture', hasText: false, text: '', bounds: { left: 1, top: 1, width: 6, height: 4 } }),
    shp({ text: '正文段落内容，用于说明图片所展示的主题与关键信息……', textStyle: { size: 14 }, bounds: { left: 8, top: 2, width: 4, height: 3 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedRole, 'ai_image');
  // 右侧 14pt 正文无占位符/结构信号 → 信号不足（设计行为）
  assert.equal(find(recs, 1).ruleId, 'R-FALLBACK');
  assert.ok(find(recs, 1).confidence < 0.8);
});

test('背景大图（>65% 面积）→ fixed；角落小图 → fixed(Logo)；普通中图 → ai_image', () => {
  const shapes = [
    shp({ type: 'picture', hasText: false, text: '', bounds: { left: 0, top: 0, width: 13.3, height: 7.4 } }),
    shp({ type: 'picture', hasText: false, text: '', bounds: { left: 0.2, top: 0.2, width: 0.8, height: 0.6 } }),
    shp({ type: 'picture', hasText: false, text: '', bounds: { left: 2, top: 2, width: 4, height: 3 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedRole, 'fixed', '背景大图');
  assert.equal(find(recs, 1).recommendedRole, 'fixed', '角落小图=Logo');
  assert.equal(find(recs, 2).recommendedRole, 'ai_image');
});

test('页码：底部小数字 → ai_text + seq（高置信）', () => {
  const shapes = [
    shp({ text: '05', textStyle: { size: 10 }, bounds: { left: 12.5, top: 7.1, width: 0.6, height: 0.3 } }),
    shp({ text: '第 3 页', textStyle: { size: 10 }, bounds: { left: 0.5, top: 7.1, width: 0.8, height: 0.3 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedRole, 'ai_text');
  assert.equal(find(recs, 0).recommendedSemanticRole, 'seq');
  assert.ok(find(recs, 0).confidence >= 0.8);
  assert.equal(find(recs, 1).recommendedRole, 'ai_text');
});

test('数据文本：纯数字/百分比 → ai_text + other（等待手动确认）', () => {
  const shapes = [shp({ text: '18.5%', textStyle: { size: 24 }, bounds: { left: 1, top: 1, width: 2, height: 0.8 } })];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedSemanticRole, 'other');
  assert.ok(find(recs, 0).confidence < HIGH_CONFIDENCE);
});

test('垂直重复文本 → body；水平重复文本 → other（等待手动确认）', () => {
  const shapes = [
    shp({ text: '要点一', textStyle: { size: 14 }, bounds: { left: 1, top: 2, width: 4, height: 0.4 } }),
    shp({ text: '要点二', textStyle: { size: 14 }, bounds: { left: 1, top: 2.5, width: 4, height: 0.4 } }),
    shp({ text: '要点三', textStyle: { size: 14 }, bounds: { left: 1, top: 3, width: 4, height: 0.4 } }),
    shp({ text: 'A', textStyle: { size: 14 }, bounds: { left: 2, top: 5, width: 0.6, height: 0.4 } }),
    shp({ text: 'B', textStyle: { size: 14 }, bounds: { left: 3, top: 5, width: 0.6, height: 0.4 } }),
    shp({ text: 'C', textStyle: { size: 14 }, bounds: { left: 4, top: 5, width: 0.6, height: 0.4 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  for (const i of [0, 1, 2]) assert.equal(find(recs, i).recommendedSemanticRole, 'body', 'idx ' + i);
  for (const i of [3, 4, 5]) assert.equal(find(recs, i).recommendedSemanticRole, 'other', 'idx ' + i);
});

test('几何图形/线条 → fixed（高置信）；空文本 → 低置信；0 尺寸 → 低置信', () => {
  const shapes = [
    shp({ type: 'rectangle', hasText: false, text: '', bounds: { left: 1, top: 1, width: 4, height: 2 } }),
    shp({ type: 'line', hasText: false, text: '', bounds: { left: 1, top: 3, width: 4, height: 0 } }),
    shp({ type: 'text', hasText: true, text: '   ', bounds: { left: 1, top: 4, width: 2, height: 0.5 } }),
    shp({ type: 'text', text: 'x', bounds: { left: 1, top: 5, width: 0, height: 0.5 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(find(recs, 0).recommendedRole, 'fixed');
  assert.ok(find(recs, 0).confidence >= 0.7);
  assert.equal(find(recs, 1).recommendedRole, 'fixed');
  assert.ok(find(recs, 2).confidence < HIGH_CONFIDENCE, '空文本置信度应低于自动接受阈值');
  assert.ok(find(recs, 3).confidence < HIGH_CONFIDENCE, '0 尺寸置信度应低');
});

test('表格 → 表格位（table）高置信；chart/SmartArt 等 → fixed 低置信，不崩溃', () => {
  // 表格：已支持表格位（结构/样式/尺寸保留，生成时可填数据）→ 高置信推荐 table
  const tableRecs = analyzeShapesByRules([shp({ type: 'table', hasText: false, text: '', bounds: { left: 1, top: 1, width: 5, height: 3 } })], PAGE);
  assert.equal(tableRecs[0].recommendedRole, 'table');
  assert.ok(tableRecs[0].confidence >= HIGH_CONFIDENCE, '表格置信度高');
  assert.equal(tableRecs[0].isTableSlot, true);
  // 其余复杂对象：仍 fixed 低置信
  for (const type of ['chart', 'smartArt', 'media']) {
    const recs = analyzeShapesByRules([shp({ type, hasText: false, text: '', bounds: { left: 1, top: 1, width: 5, height: 3 } })], PAGE);
    assert.equal(recs[0].recommendedRole, 'fixed', type);
    assert.ok(recs[0].confidence < HIGH_CONFIDENCE, type + ' 置信度低');
  }
});

test('版式/母版元素 → 默认 fixed', () => {
  const shapes = [
    shp({ text: '页眉标题', textStyle: { size: 12 }, source: 'master', bounds: { left: 0.5, top: 0.2, width: 4, height: 0.4 } })
  ];
  const recs = analyzeShapesByRules(shapes, PAGE);
  assert.equal(recs[0].recommendedRole, 'fixed');
});

test('高置信度阈值常量 = 0.8，且建议带 constraints/prompt', () => {
  assert.equal(HIGH_CONFIDENCE, 0.8);
  const shapes = [shp({ text: '年度总结', textStyle: { size: 36 }, bounds: { left: 1, top: 0.5, width: 8, height: 1 } })];
  const recs = analyzeShapesByRules(shapes, PAGE);
  const r = recs[0];
  assert.ok(r.suggestedPrompt && r.suggestedPrompt.length > 0, '应有建议提示词');
  assert.ok(r.suggestedConstraints && r.suggestedConstraints.maxChars > 0, '应有长度约束');
  assert.ok(r.reason.length > 0, '应有原因');
});
