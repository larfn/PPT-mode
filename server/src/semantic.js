// 模板语义层（Semantic Template Layer）
// 给 TemplateShape 增加可选语义字段，让模板不仅知道元素是 ai_text / ai_image / var / fixed，
// 还知道元素在页面中的语义用途（标题/正文/结论…），为 AI 生成、模板检索、布局检查提供结构化依据。
//
// 兼容性原则：
//  - 所有语义字段都是可选的；旧模板没有这些字段时按 undefined 处理（默认行为）。
//  - schemaVersion 保持 additive 演进：读取端从不因 schemaVersion 拒绝模板。
//  - 数字约束为 0 或缺失 = 不限制（maxChars: 0 绝不表示"空文本"）。
//  - 非法枚举归为「不指定(other)」；非法数字一律丢弃，绝不写入模板。

'use strict';

// 语义角色枚举（保持简单、可扩展；不在列表中的值视为非法）
// 语义角色枚举（精简为 8 类；旧模板中的旧角色值（bullet/label/data/conclusion/page_number/tag 等）
// 不在列表中，归一化时一律归为 'other'（不指定），由用户重新归类。
const SEMANTIC_ROLES = new Set([
  'title', 'subtitle', 'body', 'seq', 'date', 'caption', 'formula', 'other'
]);

// 数字约束字段：非负整数；0 = 不限制
const SEMANTIC_NUMBER_FIELDS = ['maxChars', 'maxLines', 'minChars', 'preferredLength'];
// 字符串语义字段
const SEMANTIC_STRING_FIELDS = ['contentType', 'generationInstruction', 'translateSource'];
// 布尔语义字段（'true'/'false' 字符串兼容）
const SEMANTIC_BOOL_FIELDS = ['required', 'translate'];

function isValidRole(v) {
  return typeof v === 'string' && SEMANTIC_ROLES.has(v);
}

// 校验并清洗单个形状的语义字段（幂等）：
//  - semanticRole 必须是合法枚举，否则删除
//  - maxChars/maxLines/minChars/preferredLength 必须是非负整数（数字字符串自动转），否则删除；0 保留（=不限制）
//  - required 只接受布尔（'true'/'false' 字符串兼容），否则删除
//  - contentType / generationInstruction 去空白，空串删除
// 不强制任何元素拥有语义字段：fixed / image / variable 等不适合的元素保持为空。
function normalizeShapeSemantics(s) {
  if (!s || typeof s !== 'object') return s;
  const out = { ...s };

  if (out.semanticRole !== undefined) {
    // 非法/旧角色值 → 归为「不指定」（不删除：语义字段保持可读，用户可重新归类）
    if (!isValidRole(out.semanticRole)) out.semanticRole = 'other';
  }
  for (const f of SEMANTIC_NUMBER_FIELDS) {
    const v = out[f];
    if (v === undefined || v === null || v === '') { delete out[f]; continue; }
    const n = typeof v === 'number' ? v : Number(v);
    if (Number.isFinite(n) && Number.isInteger(n) && n >= 0) out[f] = n;
    else delete out[f];
  }
  for (const f of SEMANTIC_BOOL_FIELDS) {
    const v = out[f];
    if (v === undefined || v === null) { delete out[f]; continue; }
    if (v === true || v === 'true') out[f] = true;
    else if (v === false || v === 'false') out[f] = false;
    else delete out[f];
  }
  for (const f of SEMANTIC_STRING_FIELDS) {
    const v = out[f];
    if (v === undefined || v === null) { delete out[f]; continue; }
    if (typeof v === 'string') {
      const t = v.trim();
      if (t) out[f] = t; else delete out[f];
    } else {
      delete out[f];
    }
  }
  return out;
}

// 归一化整个模板的语义字段（读取与保存时都调用，保证 JSON 往返稳定）
function normalizeTemplate(template) {
  if (!template || typeof template !== 'object') return template;
  const out = { ...template };
  if (Array.isArray(out.shapes)) {
    out.shapes = out.shapes.map(normalizeShapeSemantics);
  }
  return out;
}

// 取出形状的语义约束对象（只含已清洗字段；无约束时返回 {}）
function shapeSemantics(s) {
  if (!s || typeof s !== 'object') return {};
  const out = {};
  if (isValidRole(s.semanticRole)) out.semanticRole = s.semanticRole;
  if (typeof s.contentType === 'string' && s.contentType.trim()) out.contentType = s.contentType.trim();
  if (s.required === true || s.required === false) out.required = s.required;
  for (const f of SEMANTIC_NUMBER_FIELDS) {
    if (typeof s[f] === 'number' && Number.isInteger(s[f]) && s[f] >= 0) out[f] = s[f];
  }
  if (typeof s.generationInstruction === 'string' && s.generationInstruction.trim()) {
    out.generationInstruction = s.generationInstruction.trim();
  }
  return out;
}

// 把语义约束格式化成给 AI 的中文约束描述（用于拼进 system 提示词）。
// 接受形状对象或纯约束对象；0 / 缺失的约束不输出（避免"最多 0 个字符"这类误导）。
function formatSemanticConstraints(s) {
  if (!s || typeof s !== 'object') return '';
  const parts = [];
  if (isValidRole(s.semanticRole)) parts.push('语义角色：' + s.semanticRole);
  if (typeof s.contentType === 'string' && s.contentType.trim()) parts.push('内容类型：' + s.contentType.trim());
  const n = (v) => (typeof v === 'number' && Number.isInteger(v) && v > 0 ? v : 0);
  const maxChars = n(s.maxChars);
  const maxLines = n(s.maxLines);
  const minChars = n(s.minChars);
  const preferred = n(s.preferredLength);
  if (maxChars) parts.push('最多 ' + maxChars + ' 个字符');
  if (maxLines) parts.push('最多 ' + maxLines + ' 行');
  if (minChars) parts.push('至少 ' + minChars + ' 个字符');
  if (preferred) parts.push('建议长度约 ' + preferred + ' 个字符');
  if (s.required === true) parts.push('该位置为必填，不能留空');
  if (typeof s.generationInstruction === 'string' && s.generationInstruction.trim()) {
    parts.push(s.generationInstruction.trim());
  }
  return parts.join('；');
}

// 按模板语义约束修正外部传入的文本（MCP generate_slide / AI 填内容场景）：
//  - maxChars / maxLines 超限 → 截断（0 或缺失 = 不限制）
//  - minChars 不满足 → 只报告警告（不编造内容）
// 返回 { texts, warnings }；传入的 texts 原对象不被修改。
function applyTextConstraints(template, texts) {
  const out = { ...(texts || {}) };
  const warnings = [];
  for (const s of (template && template.shapes) || []) {
    if (!s || !s.id || !(s.id in out)) continue;
    const raw = out[s.id];
    if (typeof raw !== 'string') continue;
    let t = raw;
    if (s.maxChars > 0 && t.length > s.maxChars) t = t.slice(0, s.maxChars);
    if (s.maxLines > 0) {
      const lines = t.split('\n');
      if (lines.length > s.maxLines) t = lines.slice(0, s.maxLines).join('\n');
    }
    if (t !== raw) {
      warnings.push(s.id + '：超出语义约束已截断（' + raw.length + ' 字符 → ' + t.length + ' 字符）');
    } else if (s.minChars > 0 && t.length < s.minChars) {
      warnings.push(s.id + '：内容 ' + t.length + ' 字符，少于模板要求下限 ' + s.minChars + ' 字符');
    }
    out[s.id] = t;
  }
  return { texts: out, warnings };
}

module.exports = {
  SEMANTIC_ROLES,
  normalizeShapeSemantics,
  normalizeTemplate,
  shapeSemantics,
  formatSemanticConstraints,
  applyTextConstraints
};
