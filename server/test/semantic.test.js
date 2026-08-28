// 模板语义层单元测试：归一化 / 约束应用 / 提示词格式化
const { test } = require('node:test');
const assert = require('node:assert');
const {
  SEMANTIC_ROLES,
  normalizeShapeSemantics,
  normalizeTemplate,
  shapeSemantics,
  formatSemanticConstraints,
  applyTextConstraints
} = require('../src/semantic.js');

test('SEMANTIC_ROLES covers the designed roles', () => {
  for (const r of ['title', 'subtitle', 'body', 'seq', 'date', 'caption', 'formula', 'other']) {
    assert.ok(SEMANTIC_ROLES.has(r), 'missing role ' + r);
  }
});

test('normalizeShapeSemantics keeps valid semantic fields and drops invalid ones', () => {
  const s = normalizeShapeSemantics({
    id: 'a', role: 'ai_text',
    semanticRole: 'title',
    contentType: ' 数字指标 ',
    required: true,
    maxChars: 50, maxLines: 2, minChars: 5, preferredLength: 30,
    generationInstruction: '  用一句话概括  '
  });
  assert.equal(s.semanticRole, 'title');
  assert.equal(s.contentType, '数字指标');
  assert.equal(s.required, true);
  assert.equal(s.maxChars, 50);
  assert.equal(s.maxLines, 2);
  assert.equal(s.minChars, 5);
  assert.equal(s.preferredLength, 30);
  assert.equal(s.generationInstruction, '用一句话概括');
});

test('normalizeShapeSemantics drops invalid enum / numbers / required values', () => {
  const s = normalizeShapeSemantics({
    id: 'b', role: 'ai_text',
    semanticRole: 'banana',            // 非法/旧枚举 → 归为「不指定」
    maxChars: -5,                      // 负数 → 删除
    maxLines: 1.5,                     // 非整数 → 删除
    minChars: 'abc',                   // 非数字 → 删除
    preferredLength: Number.NaN,       // NaN → 删除
    required: 1,                       // 非布尔 → 删除
    contentType: '   ',                // 空串 → 删除
    generationInstruction: 42          // 非字符串 → 删除
  });
  assert.equal(s.semanticRole, 'other', 'invalid enum → 不指定');
  assert.equal(s.maxChars, undefined);
  assert.equal(s.maxLines, undefined);
  assert.equal(s.minChars, undefined);
  assert.equal(s.preferredLength, undefined);
  assert.equal(s.required, undefined);
  assert.equal(s.contentType, undefined);
  assert.equal(s.generationInstruction, undefined);
  assert.deepEqual(Object.keys(s).sort(), ['id', 'role', 'semanticRole']);
});

test('normalizeShapeSemantics keeps translate / translateSource (auto subtitle translation)', () => {
  const s = normalizeShapeSemantics({
    id: 't1', role: 'ai_text',
    semanticRole: 'subtitle',
    translate: true,
    translateSource: 'shp-title',
    translate: 'false' // 字符串布尔 → false
  });
  assert.equal(s.translate, false);
  assert.equal(s.translateSource, 'shp-title');
  const s2 = normalizeShapeSemantics({ id: 't2', role: 'ai_text', translate: true, translateSource: 'theme' });
  assert.equal(s2.translate, true);
  assert.equal(s2.translateSource, 'theme');
  // 非法值丢弃：非布尔 translate、非字符串 translateSource
  const s3 = normalizeShapeSemantics({ id: 't3', role: 'ai_text', translate: 1, translateSource: 42 });
  assert.equal(s3.translate, undefined);
  assert.equal(s3.translateSource, undefined);
  // 空串来源删除
  const s4 = normalizeShapeSemantics({ id: 't4', role: 'ai_text', translate: true, translateSource: '   ' });
  assert.equal(s4.translateSource, undefined);
});

test('normalizeShapeSemantics: maxChars=0 is preserved as "no limit" (never means empty text)', () => {
  const s = normalizeShapeSemantics({ id: 'c', role: 'ai_text', semanticRole: 'body', maxChars: 0, maxLines: 0, required: false });
  assert.equal(s.maxChars, 0);
  assert.equal(s.maxLines, 0);
  assert.equal(s.required, false);
});

test('normalizeShapeSemantics coerces numeric strings and boolean strings', () => {
  const s = normalizeShapeSemantics({ id: 'd', role: 'ai_text', semanticRole: 'title', maxChars: '30', required: 'false' });
  assert.equal(s.maxChars, 30);
  assert.equal(s.required, false);
});

test('normalizeTemplate normalizes every shape and tolerates missing shapes', () => {
  const t = normalizeTemplate({
    schemaVersion: 1,
    shapes: [
      { id: 'x', role: 'ai_text', semanticRole: 'nope', maxChars: -1 },
      { id: 'y', role: 'fixed', content: 'hi' },
      null
    ]
  });
  assert.equal(t.shapes[0].semanticRole, 'other', 'invalid enum → 不指定');
  assert.equal(t.shapes[0].maxChars, undefined);
  assert.equal(t.shapes[1].semanticRole, undefined, 'fixed elements stay empty');
  assert.equal(t.shapes[2], null);
  assert.equal(normalizeTemplate(undefined), undefined);
  assert.equal(normalizeTemplate({}).shapes, undefined);
});

test('shapeSemantics returns only valid fields ({} for non-shape input)', () => {
  assert.deepEqual(shapeSemantics(null), {});
  assert.deepEqual(shapeSemantics({ id: 'x', role: 'ai_text' }), {});
  const s = shapeSemantics({ id: 'x', role: 'ai_text', semanticRole: 'body', maxChars: 0, required: true });
  assert.deepEqual(s, { semanticRole: 'body', maxChars: 0, required: true });
});

test('formatSemanticConstraints builds a Chinese constraint line; 0/missing fields are omitted', () => {
  const line = formatSemanticConstraints({
    semanticRole: 'title', contentType: '公司名', required: true,
    maxChars: 20, maxLines: 1, minChars: 2, preferredLength: 12,
    generationInstruction: '用一句话概括'
  });
  assert.ok(line.includes('语义角色：title'));
  assert.ok(line.includes('内容类型：公司名'));
  assert.ok(line.includes('最多 20 个字符'));
  assert.ok(line.includes('最多 1 行'));
  assert.ok(line.includes('至少 2 个字符'));
  assert.ok(line.includes('建议长度约 12 个字符'));
  assert.ok(line.includes('必填'));
  assert.ok(line.includes('用一句话概括'));
  // 0 = 不限制，不出现在约束里
  const zero = formatSemanticConstraints({ semanticRole: 'body', maxChars: 0, maxLines: 0 });
  assert.ok(!zero.includes('最多'), '0 limit must not be rendered');
  assert.equal(formatSemanticConstraints(null), '');
  assert.equal(formatSemanticConstraints({}), '');
});

test('applyTextConstraints truncates over-long text by maxChars and maxLines', () => {
  const template = {
    shapes: [
      { id: 'a', role: 'ai_text', semanticRole: 'title', maxChars: 5, maxLines: 1 },
      { id: 'b', role: 'ai_text', maxLines: 1 },
      { id: 'c', role: 'ai_text', maxChars: 3 }
    ]
  };
  const r = applyTextConstraints(template, { a: '1234567890', b: 'x\ny\nz', c: 'hello' });
  assert.equal(r.texts.a, '12345', 'truncated to maxChars');
  assert.equal(r.texts.b, 'x', 'truncated to maxLines');
  assert.equal(r.texts.c, 'hel', 'truncated to maxChars');
  assert.equal(r.warnings.length, 3);
});

test('applyTextConstraints: maxChars=0 / missing means no limit; unknown ids untouched; minChars warns', () => {
  const template = {
    shapes: [
      { id: 'a', role: 'ai_text', maxChars: 0 },
      { id: 'b', role: 'ai_text', minChars: 10 }
    ]
  };
  const r = applyTextConstraints(template, { a: 'any length is fine', b: 'short', z: 'not in template' });
  assert.equal(r.texts.a, 'any length is fine');
  assert.equal(r.texts.b, 'short', 'minChars never fabricates content');
  assert.equal(r.texts.z, 'not in template');
  assert.deepEqual(r.warnings, ['b：内容 5 字符，少于模板要求下限 10 字符']);
});

test('applyTextConstraints does not mutate the caller object and handles non-string values', () => {
  const template = { shapes: [{ id: 'a', role: 'ai_text', maxChars: 2 }] };
  const input = { a: 'abcdef' };
  const r = applyTextConstraints(template, input);
  assert.equal(input.a, 'abcdef', 'input object must not be mutated');
  assert.equal(r.texts.a, 'ab');
  const r2 = applyTextConstraints(template, { a: 42 });
  assert.equal(r2.texts.a, 42);
  assert.deepEqual(r2.warnings, []);
  const r3 = applyTextConstraints(template, undefined);
  assert.deepEqual(r3, { texts: {}, warnings: [] });
});

test('applyTextConstraints keeps newlines inside maxChars counting (chars, not bytes)', () => {
  const template = { shapes: [{ id: 'a', role: 'ai_text', maxChars: 4 }] };
  const r = applyTextConstraints(template, { a: '中文\n内容' });
  assert.equal(r.texts.a, '中文\n内', 'maxChars counts chars, newline included');
});
