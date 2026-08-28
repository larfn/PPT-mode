// AI 自动模板分析：可选增强层（严格 JSON 输出校验 + 失败完全回退规则分类）
//
// 定位：规则分类器（前端 lib/analyze.ts）是可靠基础；本模块仅在配置了文本 AI 时被调用。
// 任何失败（未配置/超时/网络/非法 JSON/shapeId 不存在/非法枚举/confidence 非数字/
// constraints 非法字段）都返回 { ok:false, error }，调用方（前端）回退到规则结果，
// 绝不让模板保存/分析流程失败。
'use strict';
const { generateText } = require('./textService.js');
const { SEMANTIC_ROLES } = require('./semantic.js');

const VALID_ROLES = new Set(['ai_image', 'ai_text', 'manual_var', 'fixed', 'table']);
const NUM_CONSTRAINTS = new Set(['maxChars', 'maxLines', 'minChars', 'preferredLength']);

const SYSTEM_PROMPT = [
  '你是 PPT 模板结构分析助手。根据每个形状的特征（名称/类型/字号/加粗/文本/位置/尺寸/来源/占位符类型/对齐/颜色/行数），推荐最合适的角色与语义角色。',
  '优先级：若输入包含 phType（占位符类型 title/ctrTitle/subTitle/body/pic/tbl/sldNum/dt/ftr/hdr/chart/dgm/media/clipArt/obj/content 等），优先尊重官方语义：title→主标题、subTitle→副标题、pic→图片、tbl→表格、sldNum→序号、dt→日期、ftr/hdr→固定、body→正文；仅当占位符语义明显与实际内容不符时再按特征调整。',
  '只输出一个 JSON 数组，不要输出任何其他文字、解释、Markdown 或代码围栏。',
  '数组元素格式：{"shapeId":"0","recommendedRole":"ai_text","recommendedSemanticRole":"title","confidence":0.9,"reason":"占位符类型=title 且位于页面上部","suggestedConstraints":{"maxChars":30,"maxLines":1}}',
  'reason 请简述命中的关键信号（占位符类型/字号/位置/对齐/颜色/文本特征），帮助用户理解判断依据。',
  'recommendedRole 只能是 ai_image / ai_text / manual_var / fixed / table 之一（table 仅用于表格形状）。',
  'recommendedSemanticRole 只能是 title/subtitle/body/seq/date/caption/formula/other 之一，或省略。',
  'confidence 必须是 0 到 1 之间的数字。suggestedConstraints 只能包含 maxChars/maxLines/minChars/preferredLength（非负整数）/required（布尔）/contentType/generationInstruction（字符串）。',
  '只输出识别结果，绝不输出 suggestedPrompt：提示词由内置规则统一决定（主标题→围绕主题输入主标题，副标题→围绕主题输入副标题，正文→围绕主题输入正文，其余角色无提示词），AI 识别只增强识别效果、不改动提示词。',
  '每个输入形状都必须输出一个元素，shapeId 使用输入的 shapeId。'
].join('\n');

// 从模型输出中提取 JSON（容忍可能的代码围栏/前后缀）
function extractJson(text) {
  if (!text || typeof text !== 'string') return null;
  let t = text.trim();
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fence) t = fence[1].trim();
  // 找第一个 [ 到最后一个 ]
  const start = t.indexOf('[');
  const end = t.lastIndexOf(']');
  if (start < 0 || end <= start) return null;
  try { return JSON.parse(t.slice(start, end + 1)); } catch { return null; }
}

// 校验单条推荐；非法（shapeId 不存在/枚举/数字/constraints）返回 null
function validateRecommendation(raw, validShapeIds) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null;
  const shapeId = String(raw.shapeId ?? raw.idx ?? '');
  if (!validShapeIds.has(shapeId)) return null;              // 不存在的 shapeId
  if (!VALID_ROLES.has(raw.recommendedRole)) return null;    // 非法角色
  if (raw.recommendedSemanticRole !== undefined && raw.recommendedSemanticRole !== null &&
      !SEMANTIC_ROLES.has(raw.recommendedSemanticRole)) return null; // 非法语义角色
  const confidence = Number(raw.confidence);
  if (!Number.isFinite(confidence) || confidence < 0 || confidence > 1) return null; // confidence 非数字
  const constraints = {};
  if (raw.suggestedConstraints && typeof raw.suggestedConstraints === 'object' && !Array.isArray(raw.suggestedConstraints)) {
    for (const [k, v] of Object.entries(raw.suggestedConstraints)) {
      if (NUM_CONSTRAINTS.has(k)) {
        const n = Number(v);
        if (Number.isFinite(n) && Number.isInteger(n) && n >= 0) constraints[k] = n;
      } else if (k === 'required') {
        if (typeof v === 'boolean') constraints[k] = v;
      } else if ((k === 'contentType' || k === 'generationInstruction') && typeof v === 'string' && v.trim()) {
        constraints[k] = v.trim().slice(0, 200);
      }
      // 未知字段忽略（不报错）
    }
  }
  return {
    shapeId,
    recommendedRole: raw.recommendedRole,
    recommendedSemanticRole: typeof raw.recommendedSemanticRole === 'string' ? raw.recommendedSemanticRole : undefined,
    confidence,
    reason: typeof raw.reason === 'string' ? raw.reason.slice(0, 300) : '',
    suggestedPrompt: typeof raw.suggestedPrompt === 'string' ? raw.suggestedPrompt.slice(0, 500) : undefined,
    suggestedConstraints: Object.keys(constraints).length ? constraints : undefined,
    source: 'ai'
  };
}

// 主入口：调 AI 分析 → 严格校验 → 返回 { ok, recommendations } 或 { ok:false, error }
// cfg: { baseUrl, apiKey, model }（来自 config.text）；shapes: 摘要数组（含 shapeId）
async function analyzeWithAI(cfg, shapes) {
  if (!cfg || !cfg.apiKey) return { ok: false, error: 'text ai not configured' };
  if (!Array.isArray(shapes) || !shapes.length) return { ok: false, error: 'empty shapes' };
  const validShapeIds = new Set(shapes.map((s) => String(s.shapeId ?? s.idx ?? '')));
  try {
    const userPrompt = JSON.stringify(shapes);
    const text = await generateText({
      baseUrl: cfg.baseUrl, apiKey: cfg.apiKey, model: cfg.model,
      systemPrompt: SYSTEM_PROMPT, userPrompt, temperature: 0.2
    });
    const parsed = extractJson(text);
    if (!parsed) return { ok: false, error: 'AI 未返回合法 JSON' };
    if (!Array.isArray(parsed)) return { ok: false, error: 'AI 返回的不是数组' };
    // 部分合法：合法条目保留，非法条目跳过；全部非法 → 视为失败（调用方回退规则）
    const recommendations = [];
    for (const raw of parsed) {
      const r = validateRecommendation(raw, validShapeIds);
      if (r) recommendations.push(r);
    }
    if (!recommendations.length) return { ok: false, error: 'AI 返回的条目全部校验失败' };
    return { ok: true, recommendations };
  } catch (e) {
    return { ok: false, error: 'AI 分析失败：' + ((e && e.message) || String(e)) };
  }
}

module.exports = { analyzeWithAI, validateRecommendation, extractJson, VALID_ROLES };
