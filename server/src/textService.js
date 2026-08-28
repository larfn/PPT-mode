const { formatSemanticConstraints } = require('./semantic.js');
const { cleanAiText } = require('./cleanText.js');

// constraints: 模板语义层的可选约束（semanticRole/maxChars/maxLines/minChars/preferredLength/
// generationInstruction 等）。不改动 OpenAI 兼容 /chat/completions 协议：
// 请求体仍是 model/temperature/messages 三个标准字段，约束只拼进 system 提示词。
async function generateText({ baseUrl, apiKey, model, systemPrompt = '', userPrompt, temperature = 0.8, constraints, clean }) {
  const url = `${baseUrl.replace(/\/$/, '')}/chat/completions`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      temperature,
      messages: [
        { role: 'system', content: buildSystemPrompt(systemPrompt, constraints) },
        { role: 'user', content: userPrompt }
      ]
    })
  });
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(`text api ${res.status}: ${detail.slice(0, 200)}`);
  }
  const data = await res.json();
  const text = data.choices?.[0]?.message?.content || '';
  // AI 输出净化：所有返回统一清洗（Markdown 痕迹 / 异常换行 / 超长文本）。
  //  - 始终剥离 Markdown 语法、规整换行（对任何调用安全，含表格 JSON）
  //  - clean.plain=true（前端「纯文本」模式）额外剥离列表编号/符号
  //  - 长度上限取「输出模式限字数」与「模板语义 maxChars/maxLines」中更严格者（0/缺失 = 不限制）
  const c = clean && typeof clean === 'object' ? clean : {};
  const charCap = Math.min(c.maxChars || Infinity, constraints?.maxChars || Infinity);
  const lineCap = Math.min(c.maxLines || Infinity, constraints?.maxLines || Infinity);
  return cleanAiText(text, {
    plain: c.plain === true,
    maxChars: Number.isFinite(charCap) && charCap > 0 ? charCap : 0,
    maxLines: Number.isFinite(lineCap) && lineCap > 0 ? lineCap : 0
  });
}

// 旧模板/无约束时 systemPrompt 原样返回，行为与之前完全一致
function buildSystemPrompt(systemPrompt, constraints) {
  const block = formatSemanticConstraints(constraints || null);
  if (!block) return systemPrompt;
  return (systemPrompt ? systemPrompt + '\n' : '') + '内容约束：' + block;
}

module.exports = { generateText, buildSystemPrompt };
