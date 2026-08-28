// AI 自动模板分析（后端增强层）测试：严格 JSON 校验 / 失败回退 / HTTP
const { test, mock } = require('node:test');
const assert = require('node:assert');
const { createApp } = require('../src/index.js');
const { validateRecommendation, extractJson, analyzeWithAI, VALID_ROLES } = require('../src/analyze.js');

const IDS = new Set(['0', '1', '2']);

test('validateRecommendation accepts valid entry and drops invalid ones', () => {
  const ok = validateRecommendation({
    shapeId: '0', recommendedRole: 'ai_text', recommendedSemanticRole: 'title',
    confidence: 0.9, reason: '最大字号', suggestedPrompt: '写标题',
    suggestedConstraints: { maxChars: 30, maxLines: 1, required: true, contentType: '标题' }
  }, IDS);
  assert.equal(ok.shapeId, '0');
  assert.equal(ok.recommendedRole, 'ai_text');
  assert.equal(ok.suggestedConstraints.maxChars, 30);
  assert.equal(ok.suggestedConstraints.required, true);
  assert.equal(ok.suggestedConstraints.contentType, '标题');
  assert.equal(ok.source, 'ai');
  // 不存在的 shapeId
  assert.equal(validateRecommendation({ shapeId: '999', recommendedRole: 'ai_text', confidence: 0.9 }, IDS), null);
  // 非法角色
  assert.equal(validateRecommendation({ shapeId: '0', recommendedRole: 'banana', confidence: 0.9 }, IDS), null);
  // 非法语义角色
  assert.equal(validateRecommendation({ shapeId: '0', recommendedRole: 'ai_text', recommendedSemanticRole: 'nope', confidence: 0.9 }, IDS), null);
  // confidence 非数字 / 越界
  assert.equal(validateRecommendation({ shapeId: '0', recommendedRole: 'ai_text', confidence: 'high' }, IDS), null);
  assert.equal(validateRecommendation({ shapeId: '0', recommendedRole: 'ai_text', confidence: 1.5 }, IDS), null);
  assert.equal(validateRecommendation({ shapeId: '0', recommendedRole: 'ai_text', confidence: -0.1 }, IDS), null);
  // constraints 非法字段被忽略 / 非法数值被丢弃
  const c = validateRecommendation({ shapeId: '0', recommendedRole: 'fixed', confidence: 0.5, suggestedConstraints: { maxChars: -5, evil: 'x', maxLines: 'two', preferredLength: 12 } }, IDS);
  assert.deepEqual(c.suggestedConstraints, { preferredLength: 12 });
  // 非对象
  assert.equal(validateRecommendation('str', IDS), null);
  assert.equal(validateRecommendation(null, IDS), null);
});

test('extractJson tolerates code fences and surrounding text', () => {
  assert.deepEqual(extractJson('[{"a":1}]'), [{ a: 1 }]);
  assert.deepEqual(extractJson('```json\n[{"a":1}]\n```'), [{ a: 1 }]);
  assert.deepEqual(extractJson('好的，结果如下：\n[{"a":1}]\n希望有帮助'), [{ a: 1 }]);
  assert.equal(extractJson('不是 JSON'), null);
  assert.equal(extractJson('{broken'), null);
  assert.equal(extractJson(null), null);
  assert.equal(extractJson(''), null);
});

test('analyzeWithAI: not configured → error (frontend falls back to rules)', async () => {
  const r = await analyzeWithAI({ apiKey: '' }, [{ shapeId: '0' }]);
  assert.equal(r.ok, false);
  assert.equal(r.error, 'text ai not configured');
  const r2 = await analyzeWithAI(null, []);
  assert.equal(r2.ok, false);
});

test('analyzeWithAI: malformed JSON from AI → error (fallback)', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, json: async () => ({ choices: [{ message: { content: '这不是JSON' } }] })
  }));
  const r = await analyzeWithAI({ apiKey: 'k', baseUrl: 'https://x.com', model: 'm' }, [{ shapeId: '0' }]);
  assert.equal(r.ok, false);
  assert.ok(r.error.includes('JSON') || r.error.includes('数组'));
  mock.restoreAll();
});

test('analyzeWithAI: AI timeout / network failure → error (fallback)', async () => {
  mock.method(global, 'fetch', async () => { const e = new Error('network'); throw e; });
  const r = await analyzeWithAI({ apiKey: 'k', baseUrl: 'https://x.com', model: 'm' }, [{ shapeId: '0' }]);
  assert.equal(r.ok, false);
  assert.ok(r.error.includes('分析失败'));
  mock.restoreAll();
});

test('analyzeWithAI: partial valid results are kept; unknown shapeId dropped', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, json: async () => ({
      choices: [{ message: { content: JSON.stringify([
        { shapeId: '0', recommendedRole: 'ai_text', recommendedSemanticRole: 'title', confidence: 0.9, reason: 'ok' },
        { shapeId: '999', recommendedRole: 'ai_text', confidence: 0.9, reason: 'bad id' },
        { shapeId: '1', recommendedRole: 'invalid_role', confidence: 0.9, reason: 'bad role' },
        { shapeId: '2', recommendedRole: 'fixed', confidence: 0.4, reason: 'ok2', suggestedConstraints: { maxChars: 10 } }
      ]) } }]
    })
  }));
  const r = await analyzeWithAI({ apiKey: 'k', baseUrl: 'https://x.com', model: 'm' }, [{ shapeId: '0' }, { shapeId: '1' }, { shapeId: '2' }]);
  assert.equal(r.ok, true);
  assert.equal(r.recommendations.length, 2, '非法条目被丢弃，合法保留');
  assert.equal(r.recommendations[0].shapeId, '0');
  assert.equal(r.recommendations[0].recommendedSemanticRole, 'title');
  assert.equal(r.recommendations[1].shapeId, '2');
  mock.restoreAll();
});

test('analyzeWithAI: all entries invalid → failure (fallback to rules)', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, json: async () => ({ choices: [{ message: { content: '[{"shapeId":"0","recommendedRole":"nope","confidence":9}]' } }] })
  }));
  const r = await analyzeWithAI({ apiKey: 'k', baseUrl: 'https://x.com', model: 'm' }, [{ shapeId: '0' }]);
  assert.equal(r.ok, false);
  mock.restoreAll();
});

test('HTTP POST /api/analyze: no shapes → 400; AI 失败 → ok:false fallback（不依赖真实配置）', async () => {
  // 无论真实 config 是否配置了 API Key，AI 返回非 JSON → 必然 ok:false
  const realFetch = global.fetch.bind(global);
  mock.method(global, 'fetch', async () => ({
    ok: true, json: async () => ({ choices: [{ message: { content: '这完全不是 JSON' } }] })
  }));
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const bad = await realFetch(base + '/api/analyze', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({}) });
    assert.equal(bad.status, 400);
    // 默认配置无 text.apiKey → 返回 ok:false（前端回退规则）
    const res = await realFetch(base + '/api/analyze', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ shapes: [{ shapeId: '0', text: '很长很长的文本'.repeat(50), fontSize: 32 }] })
    });
    const body = await res.json();
    assert.equal(body.ok, false, 'AI 失败必须 ok:false（前端回退规则）');
  } finally { server.close(); mock.restoreAll(); }
});

test('VALID_ROLES covers the designed roles', () => {
  for (const r of ['ai_image', 'ai_text', 'manual_var', 'fixed']) assert.ok(VALID_ROLES.has(r));
});
