// 公式位工具：把用户粘贴/输入的公式文本规范化为 LaTeX（本地兜底，不依赖 AI）。
// 说明：向导里公式位的「AI 生成」会优先给出完整 LaTeX；未配 AI 或用户直接输入时，
// 用本模块做常见 Unicode 数学符号/上下标的规范化，交给后端转成 OMML（专业型公式）。

// 今日日期 → YYYY/MM/DD（如 2026/05/15）
export function todayStr(): string {
  const d = new Date();
  const p = (n: number): string => String(n).padStart(2, '0');
  return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate());
}

// Unicode 数学符号 → LaTeX（按优先级替换，长符号在前）
const SYMBOL_MAP: Array<[RegExp, string]> = [
  [/\times/g, '\\times'],
  [/\div/g, '\\div'],
  [/\pm/g, '\\pm'],
  [/\cdot/g, '\\cdot'],
  [/\le|≤/g, '\\leq'],
  [/\ge|≥/g, '\\geq'],
  [/\approx|≈/g, '\\approx'],
  [/\ne|≠/g, '\\neq'],
  [/\infty|∞/g, '\\infty'],
  [/\Delta|Δ/g, '\\Delta'],
  [/\nabla|∇/g, '\\nabla'],
  [/\partial|∂/g, '\\partial'],
  [/\sqrt|√/g, '\\sqrt{}'],
  [/×/g, '\\times'],
  [/÷/g, '\\div'],
  [/±/g, '\\pm'],
  [/⋅/g, '\\cdot'],
  [/π/g, '\\pi'],
  [/𝜋/g, '\\pi'],
  [/∑/g, '\\sum'],
  [/Σ/g, '\\Sigma'],
  [/∫/g, '\\int'],
  [/∈/g, '\\in'],
  [/∉/g, '\\notin'],
  [/≤/g, '\\leq'],
  [/≥/g, '\\geq'],
  [/≈/g, '\\approx'],
  [/≠/g, '\\neq'],
  [/∞/g, '\\infty']
];

// 上下标 Unicode 数字
const SUP_DIGITS: Record<string, string> = { '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9' };
const SUB_DIGITS: Record<string, string> = { '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4', '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9' };

// 把「_abc」/「^abc」这类紧跟字母/数字的上下标改写成带花括号的 LaTeX 形式
function braceRuns(s: string): string {
  return s
    .replace(/_([A-Za-z0-9τπ]{2,})/g, '_{$1}')
    .replace(/\^([A-Za-z0-9τπ]{2,})/g, '^{$1}');
}

// 公式文本 → LaTeX（本地规范化；只处理常见情况，复杂公式请用 AI 生成完整 LaTeX）
export function localTextToLatex(text: string): string {
  let s = String(text || '').trim();
  if (!s) return '';
  // 1) Unicode 上下标数字
  for (const [u, d] of Object.entries(SUP_DIGITS)) s = s.split(u).join('^{' + d + '}');
  for (const [u, d] of Object.entries(SUB_DIGITS)) s = s.split(u).join('_{' + d + '}');
  // 2) 常见数学符号
  for (const [re, rep] of SYMBOL_MAP) s = s.replace(re, rep);
  // 3) 上下标 runs 加花括号
  s = braceRuns(s);
  // 4) 最外层整体分式：形如 (a)/(b) → \frac{a}{b}（只处理整串恰为一个分式的情况）
  const frac = s.match(/^\(([^()]+)\)\/\(([^()]+)\)$/);
  if (frac) s = '\\frac{' + frac[1].trim() + '}{' + frac[2].trim() + '}';
  return s;
}
