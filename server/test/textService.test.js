
const { test, mock } = require('node:test');
const assert = require('node:assert');
const { generateText } = require('../src/textService.js');

test('generateText calls chat/completions and returns content', async () => {
  mock.method(global, 'fetch', async (url, opts) => {
    assert.equal(url, 'https://api.deepseek.com/chat/completions');
    const body = JSON.parse(opts.body);
    assert.equal(body.model, 'deepseek-chat');
    const headers = new Map(Object.entries(opts.headers));
    assert.equal(headers.get('Authorization'), 'Bearer sk-test');
    return {
      ok: true,
      json: async () => ({ choices: [{ message: { content: '生成的内容' } }] })
    };
  });
  const text = await generateText({ baseUrl: 'https://api.deepseek.com', apiKey: 'sk-test', model: 'deepseek-chat', systemPrompt: '你是PPT文案助手', userPrompt: '写标题' });
  assert.equal(text, '生成的内容');
  mock.restoreAll();
});

test('generateText throws on non-ok response', async () => {
  mock.method(global, 'fetch', async () => ({ ok: false, status: 401, text: async () => 'unauthorized' }));
  await assert.rejects(
    generateText({ baseUrl: 'https://x.com', apiKey: 'k', model: 'm', systemPrompt: '', userPrompt: 'hi' }),
    /401/
  );
  mock.restoreAll();
});

// —— 模板语义层：约束传递 ——

test('generateText with constraints bakes them into system prompt, protocol stays OpenAI-compatible', async () => {
  let captured = null;
  mock.method(global, 'fetch', async (url, opts) => {
    captured = { url, body: JSON.parse(opts.body) };
    return { ok: true, json: async () => ({ choices: [{ message: { content: '标题内容' } }] }) };
  });
  const text = await generateText({
    baseUrl: 'https://api.deepseek.com', apiKey: 'sk-test', model: 'deepseek-chat',
    systemPrompt: '你是PPT文案助手', userPrompt: '写标题',
    constraints: { semanticRole: 'title', maxChars: 20, maxLines: 1, required: true, generationInstruction: '用一句话概括' }
  });
  assert.equal(text, '标题内容');
  // /chat/completions 协议不变：只发标准字段
  assert.ok(captured.url.endsWith('/chat/completions'));
  assert.deepEqual(Object.keys(captured.body).sort(), ['messages', 'model', 'temperature']);
  // 约束拼进 system 提示词
  const sys = captured.body.messages.find((m) => m.role === 'system').content;
  assert.ok(sys.startsWith('你是PPT文案助手'), 'original system prompt preserved');
  assert.ok(sys.includes('内容约束：语义角色：title；最多 20 个字符；最多 1 行；该位置为必填，不能留空；用一句话概括'));
  mock.restoreAll();
});

test('generateText without constraints keeps system prompt byte-identical', async () => {
  let captured = null;
  mock.method(global, 'fetch', async (url, opts) => {
    captured = JSON.parse(opts.body);
    return { ok: true, json: async () => ({ choices: [{ message: { content: 'ok' } }] }) };
  });
  await generateText({ baseUrl: 'https://x.com', apiKey: 'k', model: 'm', systemPrompt: '原样', userPrompt: 'hi' });
  assert.equal(captured.messages[0].content, '原样');
  mock.restoreAll();
});

test('generateText ignores empty / junk constraints', async () => {
  let captured = null;
  mock.method(global, 'fetch', async (url, opts) => {
    captured = JSON.parse(opts.body);
    return { ok: true, json: async () => ({ choices: [{ message: { content: 'ok' } }] }) };
  });
  await generateText({ baseUrl: 'https://x.com', apiKey: 'k', model: 'm', systemPrompt: '', userPrompt: 'hi', constraints: {} });
  assert.equal(captured.messages[0].content, '');
  await generateText({ baseUrl: 'https://x.com', apiKey: 'k', model: 'm', systemPrompt: '', userPrompt: 'hi', constraints: { semanticRole: 'title', maxChars: 0 } });
  assert.equal(captured.messages[0].content, '内容约束：语义角色：title');
  mock.restoreAll();
});
