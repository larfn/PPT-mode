// 输出模式：AI 生成文本的「风格开关」（整段 / 分点 / 精简 / 限字数）。
// 状态存 localStorage（记住上次选择，默认整段，最安全的防污染状态）。
// 模式会同时变成：
//   ① 附加给 AI 的提示词指令（buildOutputInstruction）
//   ② 传给后端 /api/text/generate 的清洗规则（outputCleanFlags → clean 字段）
// 用户从未手动切换（touched=false）时，模板里标注为「要点/列表」的文本位按分点处理
// （模板设计意图优先）；手动切换过后全局模式优先。

export interface OutputMode {
  plain: boolean;    // 整段：一段连续正文，不用 Markdown / 列表符号
  bullets: boolean;  // 分点：每个要点单独一行，保留编号
  condense: boolean; // 精简：删冗余、句子短
  maxChars: number;  // 0 = 不限
  touched?: boolean; // 用户是否手动切换过
}

const KEY = 'pptai.outputMode.v1';
const DEFAULT_MODE: OutputMode = { plain: true, bullets: false, condense: false, maxChars: 0, touched: false };

// 限字数下拉的可选值（0 = 不限）
export const LIMIT_CHOICES = [0, 50, 100, 150, 200, 300];

export function loadOutputMode(): OutputMode {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return { ...DEFAULT_MODE };
    const p = JSON.parse(raw) as Partial<OutputMode>;
    const maxChars = Number(p.maxChars);
    return {
      plain: p.plain !== false,
      bullets: p.bullets === true,
      condense: p.condense === true,
      maxChars: Number.isFinite(maxChars) && maxChars > 0 ? maxChars : 0,
      touched: p.touched === true
    };
  } catch {
    return { ...DEFAULT_MODE };
  }
}

export function saveOutputMode(m: OutputMode): void {
  try { localStorage.setItem(KEY, JSON.stringify(m)); } catch { /* 忽略 */ }
}

// 附加到 system 提示词的输出风格指令（追加在现有 system 之后）
export function buildOutputInstruction(m: OutputMode): string {
  const parts: string[] = [];
  if (m.bullets) {
    parts.push('用分点列表输出：每个要点单独一行，行首使用「1. 」「2. 」或「· 」编号，要点数量 3~6 个，不要使用 Markdown 标记，不要输出解释性文字。');
  } else {
    parts.push('只输出一段连续的整段正文：不要使用任何 Markdown 标记（如 ###、**、反引号、列表符号、编号），不要分点、不要输出解释性文字，直接给出正文内容。');
  }
  if (m.condense) parts.push('内容尽量精简：删除冗余修饰，只保留核心信息，句子尽量短。');
  if (m.maxChars > 0) parts.push('全文总长度不超过 ' + m.maxChars + ' 个字符。');
  return parts.join('');
}

// 传给后端 /api/text/generate 的清洗规则（clean 字段；无规则时返回 undefined）
// keepList=true：不清列表编号（表格 AI 等场景，避免洗掉单元格里合法的「1. 步骤」）
export function outputCleanFlags(m: OutputMode, opts: { keepList?: boolean } = {}):
  { plain?: boolean; maxChars?: number } | undefined {
  const out: { plain?: boolean; maxChars?: number } = {};
  if (!m.bullets && !opts.keepList) out.plain = true;
  if (m.maxChars > 0) out.maxChars = m.maxChars;
  return out.plain !== undefined || out.maxChars !== undefined ? out : undefined;
}

// 解析某文本位实际生效的模式与清洗规则
export function resolveSlotMode(mode: OutputMode, slotRole?: string): {
  instruction: string;
  clean?: { plain?: boolean; maxChars?: number };
} {
  let m = mode;
  let keepList = false;
  // 注：语义角色已精简为 8 类（要点 bullet 已并入正文移除），所有文字位统一使用全局输出模式；
  // 表格位通过 opts.keepList 自行保留列表编号。
  return { instruction: buildOutputInstruction(m), clean: outputCleanFlags(m, { keepList }) };
}

// 模式摘要（供 UI 显示当前状态）
export function outputModeSummary(m: OutputMode): string {
  const parts: string[] = [];
  if (m.bullets) parts.push('分点');
  else parts.push('整段');
  if (m.condense) parts.push('精简');
  if (m.maxChars > 0) parts.push('≤' + m.maxChars + '字');
  return parts.join(' + ');
}
