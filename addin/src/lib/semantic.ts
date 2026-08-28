// 模板语义层：前端侧的语义角色定义与约束工具。
// 与后端 server/src/semantic.js 保持一致（枚举集合、0=不限制语义、约束描述格式）。
import type { TemplateShape } from '../api.js';

export interface SemanticRoleDef {
  value: string;
  label: string;
  hint: string;
}

// 语义角色枚举（保持简单、可扩展；不在列表中的值视为非法，保存时后端会丢弃）
// 语义角色枚举（精简为 8 类；不在列表中的值视为非法，保存时后端会归为「不指定」）
// 主标题/副标题/正文：AI 文本位，带提示词（围绕主题输入主标题/副标题/正文）；
// 序号/日期/图片诠释/公式：内容位，无提示词；不指定：以上都不符合。
export const SEMANTIC_ROLES: SemanticRoleDef[] = [
  { value: 'title', label: '主标题', hint: '页面主标题' },
  { value: 'subtitle', label: '副标题', hint: '标题下的补充说明' },
  { value: 'body', label: '正文', hint: '段落正文（含要点列表）' },
  { value: 'seq', label: '序号', hint: '编号 01~06、页码等短编号' },
  { value: 'date', label: '日期', hint: '日期，生成时自动取当日' },
  { value: 'caption', label: '图片诠释', hint: '图片下方的说明文字' },
  { value: 'formula', label: '公式', hint: '数学公式（生成时转为专业型）' },
  { value: 'other', label: '其他', hint: '不指定（以上都不符合）' }
];

const ROLE_LABEL_MAP: Record<string, string> = Object.fromEntries(SEMANTIC_ROLES.map((r) => [r.value, r.label]));

export function semanticRoleLabel(role?: string): string {
  return (role && ROLE_LABEL_MAP[role]) || '';
}

export function isValidSemanticRole(v: string): boolean {
  return !!ROLE_LABEL_MAP[v];
}

// 元素读取/展示顺序稳定化：同一来源组内按「上到下、左到右」排序（Office.js 返回 z-order，
// 可能出现竖向 1、3、2 的乱序）；跨来源保持 页面 → 版式 → 母版 分组。
export function sortShapesByPosition<T extends { bounds?: { top?: number; left?: number }; source?: string }>(shapes: T[]): T[] {
  const sourceRank: Record<string, number> = { slide: 0, layout: 1, master: 2 };
  return [...shapes].sort((a, b) => {
    const ra = sourceRank[a.source || 'slide'] ?? 0;
    const rb = sourceRank[b.source || 'slide'] ?? 0;
    if (ra !== rb) return ra - rb;
    const da = a.bounds?.top ?? 0, db = b.bounds?.top ?? 0;
    if (Math.abs(da - db) > 0.001) return da - db;
    return (a.bounds?.left ?? 0) - (b.bounds?.left ?? 0);
  });
}

// 根据字号给出合理的默认语义角色（仅用于 AI 文本位；用户可改/可不指定）
export function defaultSemanticRole(size?: number): string {
  const sz = Number(size) || 0;
  if (sz >= 28) return 'title';
  if (sz >= 20) return 'subtitle';
  return 'body';
}

// 生成一行的约束摘要，如「标题 · 最多50字 · 最多2行 · 必填」，用于向导/保存界面展示
export function constraintSummary(s: Pick<TemplateShape, 'semanticRole' | 'maxChars' | 'maxLines' | 'minChars' | 'preferredLength' | 'required' | 'generationInstruction'>): string {
  const parts: string[] = [];
  const label = semanticRoleLabel(s.semanticRole);
  if (label) parts.push(label);
  if (s.maxChars && s.maxChars > 0) parts.push('最多' + s.maxChars + '字');
  if (s.maxLines && s.maxLines > 0) parts.push('最多' + s.maxLines + '行');
  if (s.minChars && s.minChars > 0) parts.push('至少' + s.minChars + '字');
  if (s.preferredLength && s.preferredLength > 0) parts.push('建议约' + s.preferredLength + '字');
  if (s.required) parts.push('必填');
  return parts.join(' · ');
}

// 把模板形状的语义字段取成可传给 /api/text/generate 的 constraints 对象（无字段时为 undefined）
export function shapeConstraints(s: TemplateShape): {
  semanticRole?: string; contentType?: string; required?: boolean;
  maxChars?: number; maxLines?: number; minChars?: number; preferredLength?: number;
  generationInstruction?: string;
} | undefined {
  const has = s.semanticRole || s.contentType || s.required ||
    (s.maxChars ?? 0) > 0 || (s.maxLines ?? 0) > 0 || (s.minChars ?? 0) > 0 ||
    (s.preferredLength ?? 0) > 0 || s.generationInstruction;
  if (!has) return undefined;
  return {
    semanticRole: s.semanticRole || undefined,
    contentType: s.contentType || undefined,
    required: s.required || undefined,
    maxChars: s.maxChars || undefined,
    maxLines: s.maxLines || undefined,
    minChars: s.minChars || undefined,
    preferredLength: s.preferredLength || undefined,
    generationInstruction: s.generationInstruction || undefined
  };
}
