// 模板自动分析：分层规则引擎（第三阶段：精简为 8 类语义角色 + 启发式决策表）
// 输入：ShapeInfo[]（含 phType 占位符类型）+ 页面尺寸 → 输出每个形状的推荐
// 语义角色（8 类）：主标题 title / 副标题 subtitle / 正文 body / 序号 seq /
//                  日期 date / 图片诠释 caption / 公式 formula / 不指定 other
// 提示词约定：主标题 →「围绕主题输入主标题」；副标题 →「围绕主题输入副标题」；正文 →「围绕主题输入正文」；
//            序号/日期/图片诠释/公式/不指定 一律不带提示词（AI 识别仅增强识别，不改动提示词）。
// 分层：L0 官方语义(phType) → L1 结构(source/页眉页脚) → L2 几何(位置/面积/宽高比)
//      → L3 视觉(字号/加粗/对齐/颜色) → L4 文本(正则/长度/行数) → L5 重复结构
//      → L6 上下文(与图片相邻) → L7 兜底(信号不足)
// 置信度口径：≥0.85 官方语义/强信号；0.8–0.85 强启发式（自动应用阈值 HIGH_CONFIDENCE=0.8）；
//             0.5–0.7 弱信号（仅建议，不自动应用）；<0.5 不推荐（显示「信号不足」）
// 约定：用户已手动设置的 role 绝不覆盖（调用方保护）；规则永不抛错（失败即低置信推荐）。
import type { ShapeInfo } from '../office/readSlide.js';

export interface ShapeRecommendation {
  idx: number;
  recommendedRole: string;      // ai_image | ai_text | manual_var | fixed | table
  recommendedSemanticRole?: string;
  contentType?: string;
  confidence: number;           // 0-1
  reason: string;               // 信号明细（如「标题：占位符类型=title；页面上部；最大字号 32pt」）
  suggestedPrompt?: string;
  suggestedConstraints?: {
    maxChars?: number; maxLines?: number; minChars?: number; preferredLength?: number; required?: boolean;
  };
  source: 'rule' | 'ai';
  ruleId?: string;              // 命中规则编号（如 R-TITLE-01），便于调试与调优
  isFixed?: boolean;
  isImageSlot?: boolean;
  isVarSlot?: boolean;
  isAiTextSlot?: boolean;
  isTableSlot?: boolean;
}

// 自动应用阈值：≥0.8 才自动应用（用户决策：较原 0.7 更保守）
export const HIGH_CONFIDENCE = 0.8;

// 复杂/不可重建对象：无法可靠自动标注 → 建议固定元素
const COMPLEX_TYPES = new Set(['chart', 'smartArt', 'media', 'ole', 'contentApp', 'ink', 'model3D', 'graphic', 'diagram']);
// phType 明确为不可重建/装饰类占位符
const COMPLEX_PH_TYPES = new Set(['chart', 'dgm', 'media', 'clipArt', 'orgChart']);
// phType 明确为页眉页脚类（固定，不参与生成）
const CHROME_PH_TYPES = new Set(['dt', 'ftr', 'hdr']);
// phType 文本语义类
const TITLE_PH = new Set(['title', 'ctrTitle', 'vertTitle']);
const SUBTITLE_PH = new Set(['subTitle']);
const BODY_PH = new Set(['body', 'content', 'vertBody', 'obj', 'vertContent', 'vertObj']);

// —— 文本特征正则 ——
const DATA_RE = /^[0-9][0-9,.%￥$¥€+\- ]*$/;
const PAGE_NUM_RE = /^\d{1,4}$|^\d{1,4}\s*[\/／]\s*\d{1,4}$|^第\s*\d+\s*页$/;
const DATE_RE = /^(20\d{2}|19\d{2})[年\/\-.](\d{1,2})([月\/\-.](\d{1,2})日?)?$|^\d{4}-\d{1,2}-\d{1,2}$|^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$|^[0-9]{4}\s*年\s*[0-9]{1,2}\s*月$/;
const URL_RE = /^(https?:\/\/|www\.)[^\s]+|[\w.-]+\.(com|cn|net|org|io|gov|edu)(\/[^\s]*)?$/i;
// 序号：1~2 位阿拉伯数字、带序号点/顿号的编号、中文数字编号（如 一、 1. (二)）
const SEQ_RE = /^\d{1,2}$|^\d{1,2}\s*[.、．)）]\s*$|^[（(]?[一二三四五六七八九十百]{1,3}[)）]?[.、．]?$/;
// 公式：数学符号 / 上下标 / 分式 / LaTeX 命令特征（需同时含字母数字，避免把纯符号装饰当公式）
const FORMULA_RE = /[π𝜋Π∑Σ∫√∛∜±×÷≤≥≈≠∞∆Δ∇∂∈∉]|_[A-Za-z0-9{τπ𝜏}]|\^[0-9A-Za-z{τπ𝜏}]|\\frac|\\sqrt|\\sum|\\int|\\times|\\pi|\\tau/;
const CONCLUSION_WORDS = /结论|总结|综上|总而言之|因此|所以|key\s*takeaway|summary|conclusion/i;
const COMPANY_RE = /(?:有限公司|集团|股份|公司|©|®|™|版权所有|confidential|内部资料)/;
const LOGO_NAME_RE = /logo|标志|标识|商标|brand|公司图标|分隔|装饰|线条|花纹|ornament|divider/i;
const SECTION_WORDS = /^(第?[一二三四五六七八九十百0-9]+[章节部分]|[0-9]{1,2}\s*[.．、]\s*|\d{2,4}[-–]\d{2,4}|(?:概述|背景|介绍|总结|目录|contents|agenda|overview|intro|background|summary))$/i;
const UPPER_RE = /[A-Z]{3,}/;

interface Ctx {
  pageW: number; pageH: number; pageArea: number;
}

function rec(idx: number, role: string, confidence: number, reason: string, ruleId: string, extra: Partial<ShapeRecommendation> = {}): ShapeRecommendation {
  return {
    idx, recommendedRole: role, confidence, reason, source: 'rule', ruleId, ...extra,
    isFixed: role === 'fixed',
    isImageSlot: role === 'ai_image',
    isVarSlot: role === 'manual_var',
    isAiTextSlot: role === 'ai_text',
    isTableSlot: role === 'table'
  };
}

function promptFor(role: string, sem?: string): string | undefined {
  if (role === 'ai_image') return '请描述你需要的图片（主题、风格、构图）';
  if (role === 'manual_var') return '填写变量内容';
  // 提示词约定：主标题 → 围绕主题输入主标题；副标题 → 围绕主题输入副标题；正文 → 围绕主题输入正文；
  // 序号/日期/图片诠释/公式/不指定 一律不带提示词（AI 识别仅增强识别，不改动提示词）
  switch (sem) {
    case 'title': return '围绕主题输入主标题';
    case 'subtitle': return '围绕主题输入副标题';
    case 'body': return '围绕主题输入正文';
    default: return undefined;
  }
}

// 颜色是否为「强调色」：非纯黑白灰（Office.js color 可能是 #RRGGBB 或 rgb(r,g,b) 或主题色名）
function isAccentColor(color?: string): boolean {
  if (!color) return false;
  const c = String(color).trim();
  const low = c.toLowerCase();
  if (low === 'white' || low === 'black' || low === 'gray' || low === 'grey') return false;
  if (low.includes('gray') || low.includes('grey')) return false;
  let r = 0, g = 0, b = 0;
  const hex = c.match(/^#?([0-9a-fA-F]{6})$/);
  if (hex) {
    r = parseInt(hex[1].slice(0, 2), 16); g = parseInt(hex[1].slice(2, 4), 16); b = parseInt(hex[1].slice(4, 6), 16);
  } else {
    const rgb = c.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgb) { r = Number(rgb[1]); g = Number(rgb[2]); b = Number(rgb[3]); } else return false;
  }
  return Math.abs(r - g) > 24 || Math.abs(g - b) > 24 || Math.abs(r - b) > 24;
}

// 检测重复结构：同列（垂直排列）+ 字号相近 → 列表(bullet)；同行（水平排列）+ 字号相近 → 条目(data)。
function detectRepeats(shapes: ShapeInfo[]): { vertical: Set<number>; horizontal: Set<number> } {
  const vertical = new Set<number>();
  const horizontal = new Set<number>();
  const texts = shapes
    .map((s, idx) => ({ idx, s, size: s.textStyle?.size || 0 }))
    .filter((x) => x.s.hasText && (x.s.text || '').trim());
  const nearSize = (a: number, b: number) => Math.abs(a - b) <= 6;
  for (const a of texts) {
    let group = 0;
    for (const b of texts) {
      if (a.idx === b.idx) continue;
      const sameCol = Math.abs((a.s.bounds?.left || 0) - (b.s.bounds?.left || 0)) < Math.max((a.s.bounds?.width || 0) * 0.4, 0.3);
      const overlap = Math.abs((a.s.bounds?.top || 0) - (b.s.bounds?.top || 0)) < Math.max((a.s.bounds?.height || 0) * 0.3, 0.05);
      if (sameCol && nearSize(a.size, b.size) && !overlap) group++;
    }
    if (group >= 2) vertical.add(a.idx);
  }
  for (const a of texts) {
    let group = 0;
    for (const b of texts) {
      if (a.idx === b.idx) continue;
      const sameRow = Math.abs((a.s.bounds?.top || 0) - (b.s.bounds?.top || 0)) < Math.max((a.s.bounds?.height || 0) * 0.4, 0.3);
      const overlap = Math.abs((a.s.bounds?.left || 0) - (b.s.bounds?.left || 0)) < Math.max((a.s.bounds?.width || 0) * 0.3, 0.05);
      if (sameRow && nearSize(a.size, b.size) && !overlap) group++;
    }
    if (group >= 2) horizontal.add(a.idx);
  }
  return { vertical, horizontal };
}

// 判断图片是否位于角落（Logo 特征）
function inCorner(s: ShapeInfo, ctx: Ctx): boolean {
  const { left, top, width, height } = s.bounds || { left: 0, top: 0, width: 0, height: 0 };
  const cx = left + width / 2;
  const cy = top + height / 2;
  const leftZone = cx < ctx.pageW * 0.18;
  const rightZone = cx > ctx.pageW * 0.82;
  const topZone = cy < ctx.pageH * 0.15;
  const bottomZone = cy > ctx.pageH * 0.85;
  return (topZone && (leftZone || rightZone)) || (bottomZone && (leftZone || rightZone));
}

/**
 * 分层规则分类主入口：返回每个形状的推荐（低置信=信号不足，绝不抛错）。
 */
export function analyzeShapesByRules(shapes: ShapeInfo[], slideSize: { width: number; height: number }): ShapeRecommendation[] {
  const pageW = slideSize.width || 13.33;
  const pageH = slideSize.height || 7.5;
  const pageArea = pageW * pageH;
  const ctx: Ctx = { pageW, pageH, pageArea };
  const { vertical, horizontal } = detectRepeats(shapes);
  const out: ShapeRecommendation[] = [];
  const pictureShapes: { idx: number; s: ShapeInfo }[] = [];
  const maxSizeAll = Math.max(...shapes.map((x) => x.textStyle?.size || 0), 0);

  shapes.forEach((s, idx) => {
    const b = s.bounds || { left: 0, top: 0, width: 0, height: 0 };
    const w = b.width || 0, h = b.height || 0;
    const area = w * h;
    const areaRatio = pageArea > 0 ? area / pageArea : 0;
    const text = (s.text || '').trim();
    const textLen = text.length;
    const fontSize = s.textStyle?.size || 0;
    const bold = s.textStyle?.bold === true;
    const align = s.textStyle?.align || '';
    const color = s.textStyle?.color;
    const lines = text ? text.split(/\r?\n/).length : 0;
    const type = s.type || '';
    const source = s.source || 'slide';
    const phType = s.phType || '';
    const cy = b.top + h / 2;
    const topBand = cy < ctx.pageH * 0.3;          // 页面上部 30%（标题带）
    const isFooter = cy > ctx.pageH * 0.85;         // 底部 15%
    const isHeaderEdge = cy < ctx.pageH * 0.12 && (b.left < ctx.pageW * 0.25 || b.left > ctx.pageW * 0.7);
    const chromeZone = isFooter || isHeaderEdge;    // 页眉页脚区
    const accentVisual = bold || align === 'center' || isAccentColor(color);
    const upperText = UPPER_RE.test(text) && text === text.toUpperCase();

    // ========== 表格 / 复杂对象 / 几何 ==========
    // 1) 表格：结构检测 + phType=tbl
    if (type === 'table') {
      const conf = phType === 'tbl' ? 0.95 : 0.9;
      out.push(rec(idx, 'table', conf, '表格：建议作为「表格位」保存（保留结构/样式/尺寸，生成时可填数据）', 'R-TABLE-01'));
      return;
    }
    // 2) 复杂/不可重建对象
    if (COMPLEX_TYPES.has(type)) {
      out.push(rec(idx, 'fixed', 0.6, '复杂对象（图表/SmartArt 等）建议固定元素，暂无法可靠自动标注', 'R-FIX-COMPLEX'));
      return;
    }
    // 3) 0 尺寸 / 空 bounds → 不推荐（无法判断）
    if (w <= 0 || h <= 0) {
      out.push(rec(idx, 'fixed', 0.3, '零尺寸元素，无法可靠判断', 'R-FIX-ZERO'));
      return;
    }
    // 4) 图片
    if (type === 'picture') {
      pictureShapes.push({ idx, s });
      if (areaRatio >= 0.65) {
        out.push(rec(idx, 'fixed', 0.8, '大面积图片（占页面 65% 以上）更像背景图，建议固定', 'R-IMG-BG', { contentType: '背景图' }));
        return;
      }
      if (areaRatio <= 0.08 && inCorner(s, ctx)) {
        out.push(rec(idx, 'fixed', 0.85, '角落小图（符合 Logo/页眉标记特征），建议固定', 'R-IMG-LOGO'));
        return;
      }
      const whr = h > 0 ? w / h : 0;
      if (areaRatio <= 0.3 && (whr > 4 || (w > 0 && h / w > 4))) {
        out.push(rec(idx, 'fixed', 0.8, '窄长图片（宽高比超过 4:1），更像分隔线/装饰条，建议固定', 'R-IMG-DECO'));
        return;
      }
      if (LOGO_NAME_RE.test(s.name || '')) {
        out.push(rec(idx, 'fixed', 0.8, '图片名称含 Logo/装饰/分隔特征，建议固定', 'R-IMG-NAME'));
        return;
      }
      // 图标类小图（icon）：最大边 ≤ 0.6 英寸 → 固定（装饰图标/按钮图标原样保留；大图不默认固定，作图片位参与生成）
      const iconMaxDim = Math.max(w, h);
      if (iconMaxDim > 0 && iconMaxDim <= 0.6) {
        out.push(rec(idx, 'fixed', 0.85, '图标类小图（' + w.toFixed(2) + '×' + h.toFixed(2) + '），装饰图标原样保留', 'R-IMG-ICON'));
        return;
      }
      const conf = phType === 'pic' ? 0.9 : 0.85;
      out.push(rec(idx, 'ai_image', conf, '图片位：生成时由 AI 图片替换', 'R-IMG-GEN', { suggestedPrompt: '请描述你需要的图片（主题、风格、构图）' }));
      return;
    }
    // 5) 几何图形 / 线条（无文本）
    if ((type === 'rectangle' || type === 'line') && !s.hasText) {
      out.push(rec(idx, 'fixed', 0.9, '几何图形/线条为固定元素', 'R-FIX-GEO'));
      return;
    }
    // 6) 无文本其它 → 低置信不推荐
    if (!s.hasText || !text) {
      out.push(rec(idx, 'fixed', 0.4, '空文本元素，无法可靠判断', 'R-FIX-NOTEXT'));
      return;
    }

    // ========== 文本元素 ==========
    // 7a) L0 官方语义：phType 直接定角色（最高优先级）
    if (phType) {
      if (TITLE_PH.has(phType)) {
        out.push(rec(idx, 'ai_text', 0.95, '主标题：占位符类型=' + phType + '（页面主标题）', 'R-TITLE-PH', {
          recommendedSemanticRole: 'title', contentType: '主标题',
          suggestedPrompt: '围绕主题输入主标题', suggestedConstraints: { maxChars: 30, maxLines: 1 }
        }));
        return;
      }
      if (SUBTITLE_PH.has(phType)) {
        out.push(rec(idx, 'ai_text', 0.9, '副标题：占位符类型=' + phType, 'R-SUB-PH', {
          recommendedSemanticRole: 'subtitle', contentType: '副标题',
          suggestedPrompt: '围绕主题输入副标题', suggestedConstraints: { maxChars: 50, maxLines: 2 }
        }));
        return;
      }
      if (phType === 'sldNum') {
        out.push(rec(idx, 'ai_text', 0.9, '序号：占位符类型=sldNum（幻灯片编号）', 'R-PGNUM-PH', {
          recommendedSemanticRole: 'seq', contentType: '序号'
        }));
        return;
      }
      if (CHROME_PH_TYPES.has(phType)) {
        out.push(rec(idx, 'fixed', 0.85, '页眉页脚占位符（' + phType + '），建议固定', 'R-CHROME-PH'));
        return;
      }
      if (COMPLEX_PH_TYPES.has(phType)) {
        out.push(rec(idx, 'fixed', 0.7, '复杂占位符（' + phType + '），建议固定', 'R-FIX-PH'));
        return;
      }
      // body 类占位符：按文本特征细分（序号优先，再列表/数据/正文）
      if (BODY_PH.has(phType)) {
        // 序号优先：正文占位符里的 01/02/1./(一) 等短编号 → 序号位（不能先归「不指定」）
        if (textLen <= 6 && SEQ_RE.test(text.trim())) {
          out.push(rec(idx, 'ai_text', 0.8, '序号：正文占位符内的短编号（' + text.trim() + '）', 'R-SEQ-PH', {
            recommendedSemanticRole: 'seq', contentType: '序号'
          }));
          return;
        }
        if (textLen <= 20 && DATA_RE.test(text)) {
          out.push(rec(idx, 'ai_text', 0.6, '短数字内容：占位符类型=' + phType + '，建议归为「不指定」或手动选择', 'R-DATA-PH', {
            recommendedSemanticRole: 'other'
          }));
          return;
        }
        const bodyShortLines = text.split(/\r?\n/).every((ln) => ln.trim().length <= 30);
        if ((vertical.has(idx) || lines >= 2) && bodyShortLines) {
          out.push(rec(idx, 'ai_text', 0.85, '正文（要点列表）：占位符类型=' + phType + ' 且多行垂直排列', 'R-BULLET-PH', {
            recommendedSemanticRole: 'body', contentType: '正文',
            suggestedPrompt: '围绕主题输入正文', suggestedConstraints: { maxLines: Math.max(3, lines + 1), maxChars: Math.min(300, Math.max(80, textLen * 2)) }
          }));
          return;
        }
        const conf = textLen > 40 || lines >= 3 ? 0.85 : 0.8;
        out.push(rec(idx, 'ai_text', conf, '正文：占位符类型=' + phType, 'R-BODY-PH', {
          recommendedSemanticRole: 'body', contentType: '正文',
          suggestedPrompt: '围绕主题输入正文',
          suggestedConstraints: { maxChars: Math.min(300, Math.max(80, Math.ceil(textLen * 1.5))), maxLines: Math.max(3, lines + 1) }
        }));
        return;
      }
    }

    // 7b) L1：版式/母版元素（source=layout/master）→ 一律固定（主题装饰层，不参与生成；
    // 用户可在保存页手动改回文字/图片位。日期/公司/网址等页眉页脚内容同样固定）
    if (source !== 'slide') {
      out.push(rec(idx, 'fixed', 0.85, '母版/版式元素：主题装饰层，建议固定（如需参与生成请手动改为文字/图片位）', 'R-FIX-MASTER'));
      return;
    }

    // 7c) 页码：底部/顶部角小数字（页面级）
    if (chromeZone && PAGE_NUM_RE.test(text) && textLen <= 8) {
      out.push(rec(idx, 'ai_text', 0.9, '序号：位于页眉/页脚边缘 + 编号样式', 'R-PGNUM', {
        recommendedSemanticRole: 'seq', contentType: '序号'
      }));
      return;
    }
    // 7d) 页眉页脚/日期/公司/网址 → 固定（避免生成时被 AI 改写）
    if (chromeZone && fontSize <= 14 && textLen <= 40 && (DATE_RE.test(text) || URL_RE.test(text) || COMPANY_RE.test(text))) {
      if (DATE_RE.test(text)) {
        out.push(rec(idx, 'ai_text', 0.85, '日期：页眉/页脚中的日期文本，生成时自动取当日', 'R-CHROME-DATE', {
          recommendedSemanticRole: 'date', contentType: '日期'
        }));
      } else {
        out.push(rec(idx, 'fixed', 0.85, '页眉/页脚内容（公司/网址），建议固定', 'R-CHROME'));
      }
      return;
    }
    // 7d2) 日期：任意位置的日期文本 → 日期位（生成时自动取当日）
    if (DATE_RE.test(text) && textLen <= 30) {
      out.push(rec(idx, 'ai_text', 0.9, '日期：文本符合日期格式，生成时自动取当日', 'R-DATE', {
        recommendedSemanticRole: 'date', contentType: '日期'
      }));
      return;
    }
    // 7d3) 公式：含数学符号/上下标/LaTeX 特征 → 公式位（生成时转专业型）
    if (textLen >= 2 && textLen <= 300 && FORMULA_RE.test(text)) {
      out.push(rec(idx, 'ai_text', 0.85, '公式：文本含数学符号/上下标/分式特征，生成时转为专业型', 'R-FORMULA', {
        recommendedSemanticRole: 'formula', contentType: '公式'
      }));
      return;
    }
    // 7d4) 序号：独立短编号（非页眉页脚区）→ 序号位
    if (textLen <= 6 && SEQ_RE.test(text.trim()) && !(chromeZone && fontSize <= 14)) {
      out.push(rec(idx, 'ai_text', 0.8, '序号：短编号文本（' + text.trim() + '）', 'R-SEQ', {
        recommendedSemanticRole: 'seq', contentType: '序号'
      }));
      return;
    }
    // 7e) 数据：短数字/百分比/货币
    if (textLen <= 20 && DATA_RE.test(text)) {
      const strong = /[,.%￥$¥€]/.test(text) || /\d{4,}/.test(text);
      out.push(rec(idx, 'ai_text', strong ? 0.6 : 0.55, '短数字/百分比/货币文本，建议归为「不指定」或手动选择', strong ? 'R-DATA-STRONG' : 'R-DATA', {
        recommendedSemanticRole: 'other'
      }));
      return;
    }

    // 7g) 主标题：页面最大字号（相对，不设 24pt 硬门槛）+ 短文本 + 位置/视觉信号
    // （放在重复结构判定之前：竖排的「子系统设计思路」这类节标题若恰是最大字号，必须先判为标题，不能先被列表规则打成正文）
    if (fontSize >= 16 && fontSize >= maxSizeAll - 1 && textLen <= 60) {
      const secondMax = Math.max(...shapes.filter((x, i) => i !== idx).map((x) => x.textStyle?.size || 0), 0);
      if (topBand && accentVisual) {
        out.push(rec(idx, 'ai_text', 0.85, '主标题：页面上部 + 最大字号 ' + fontSize + 'pt + 视觉突出（居中/加粗/彩色）', 'R-TITLE-STRONG', {
          recommendedSemanticRole: 'title', contentType: '主标题',
          suggestedPrompt: '围绕主题输入主标题', suggestedConstraints: { maxChars: 30, maxLines: 1 }
        }));
      } else if (topBand) {
        const conf = (fontSize - secondMax) >= 4 ? 0.85 : 0.8;
        out.push(rec(idx, 'ai_text', conf, '主标题：页面上部 + 页面最大字号 ' + fontSize + 'pt', 'R-TITLE', {
          recommendedSemanticRole: 'title', contentType: '主标题',
          suggestedPrompt: '围绕主题输入主标题', suggestedConstraints: { maxChars: 30, maxLines: 1 }
        }));
      } else {
        // 页面最大字号 + 短文本 = 标题信号很强（即使不在页面上部；数字/日期已被前面的序号/数据规则拦截）
        out.push(rec(idx, 'ai_text', 0.8, '主标题：页面最大字号 ' + fontSize + 'pt（短文本）', 'R-TITLE-MID', {
          recommendedSemanticRole: 'title', contentType: '主标题',
          suggestedPrompt: '围绕主题输入主标题', suggestedConstraints: { maxChars: 30, maxLines: 1 }
        }));
      }
      return;
    }
    // 7h) 副标题：第二大字号（与最大字号差距 ≤8pt 且短）或 紧贴标题下方
    const titleIdx = shapes.findIndex((x) => (x.textStyle?.size || 0) >= 16 && (x.textStyle?.size || 0) >= maxSizeAll - 1);
    const underTitle = titleIdx >= 0 && titleIdx !== idx && Math.abs(b.top - ((shapes[titleIdx].bounds?.top || 0) + (shapes[titleIdx].bounds?.height || 0))) < ctx.pageH * 0.12;
    if ((fontSize >= 16 && fontSize < maxSizeAll && textLen <= 50 && (maxSizeAll - fontSize) <= 8) || (fontSize >= 14 && underTitle && textLen <= 50)) {
      out.push(rec(idx, 'ai_text', 0.8, '副标题：字号仅次于主标题或紧贴其下方', 'R-SUB-UNDER', {
        recommendedSemanticRole: 'subtitle', contentType: '副标题',
        suggestedPrompt: '围绕主题输入副标题', suggestedConstraints: { maxChars: 50, maxLines: 2 }
      }));
      return;
    }
    if (align === 'center' && fontSize >= 18 && textLen <= 40) {
      out.push(rec(idx, 'ai_text', 0.7, '居中且字号较大、文本短，像副标题/引题', 'R-SUB-CENTER', {
        recommendedSemanticRole: 'subtitle', contentType: '副标题',
        suggestedPrompt: '围绕主题输入副标题', suggestedConstraints: { maxChars: 50, maxLines: 2 }
      }));
      return;
    }
    // 7f) 重复结构（在标题/副标题判定之后）：垂直 → 正文（要点列表）；水平 → 不指定。
    // 数字列兜底：竖排里的短编号/短数字在 7d4/7e 已拦截过，这里再兜底一次
    if (vertical.has(idx)) {
      if (textLen <= 6 && (SEQ_RE.test(text.trim()) || DATA_RE.test(text.trim()))) {
        out.push(rec(idx, 'ai_text', 0.75, '序号：数字列中的短编号', 'R-SEQ-V', {
          recommendedSemanticRole: 'seq', contentType: '序号'
        }));
        return;
      }
      const shortLines = text.split(/\r?\n/).every((ln) => ln.trim().length <= 30);
      if (shortLines) {
        out.push(rec(idx, 'ai_text', 0.85, '正文（要点列表）：与多个文本垂直重复排列且为短行列表', 'R-BULLET', {
          recommendedSemanticRole: 'body', contentType: '正文',
          suggestedPrompt: '围绕主题输入正文', suggestedConstraints: { maxLines: Math.max(3, lines + 1), maxChars: Math.min(300, Math.max(80, textLen * 2)) }
        }));
      } else {
        out.push(rec(idx, 'ai_text', 0.7, '与多个文本垂直重复排列但行文较长，更像多段正文', 'R-BULLET-LONG', {
          recommendedSemanticRole: 'body', contentType: '正文',
          suggestedPrompt: '围绕主题输入正文',
          suggestedConstraints: { maxChars: Math.min(300, Math.max(80, Math.ceil(textLen * 1.5))), maxLines: Math.max(3, lines + 1) }
        }));
      }
      return;
    }
    if (horizontal.has(idx)) {
      out.push(rec(idx, 'ai_text', 0.6, '与多个文本水平重复排列，像数据/条目，建议归为「不指定」或手动选择', 'R-DATA-H', {
        recommendedSemanticRole: 'other'
      }));
      return;
    }
    // 7i) 正文：小字号长文本（多行必须平均行较长，短文本多行不判正文）
    if (fontSize <= 18 && (textLen > 40 || (lines >= 3 && textLen / lines > 10))) {
      out.push(rec(idx, 'ai_text', 0.8, '正文：小字号长文本（' + textLen + ' 字/' + lines + ' 行）', 'R-BODY', {
        recommendedSemanticRole: 'body', contentType: '正文',
        suggestedPrompt: '围绕主题输入正文',
        suggestedConstraints: { maxChars: Math.min(300, Math.max(80, Math.ceil(textLen * 1.5))), maxLines: Math.max(3, lines + 1) }
      }));
      return;
    }
    // 7j) 图注：小字号短文本紧邻图片
    const nearPicture = pictureShapes.some(({ s: ps }) => {
      const pb = ps.bounds || { top: 0, height: 0, left: 0, width: 0 };
      const below = Math.abs(b.top - (pb.top + pb.height)) < ctx.pageH * 0.1;
      const above = Math.abs(pb.top - (b.top + h)) < ctx.pageH * 0.1;
      return below || above;
    });
    if (fontSize <= 14 && textLen <= 40 && nearPicture) {
      const belowPic = pictureShapes.some(({ s: ps }) => {
        const pb = ps.bounds || { top: 0, height: 0, left: 0, width: 0 };
        return Math.abs(b.top - (pb.top + pb.height)) < ctx.pageH * 0.1;
      });
      out.push(rec(idx, 'ai_text', belowPic ? 0.75 : 0.65, '图片诠释：小字号短文本' + (belowPic ? '紧邻图片下方' : '位于图片上方'), 'R-CAPTION', {
        recommendedSemanticRole: 'caption', contentType: '图片诠释'
      }));
      return;
    }

    // 7k) 结论：结论关键词 或 加粗+底部
    if (CONCLUSION_WORDS.test(text)) {
      out.push(rec(idx, 'ai_text', 0.75, '正文（结论句）：文本含结论关键词', 'R-CONCL-KW', {
        recommendedSemanticRole: 'body', contentType: '正文',
        suggestedPrompt: '围绕主题输入正文', suggestedConstraints: { maxChars: 80, maxLines: 2 }
      }));
      return;
    }
    if (bold && fontSize >= 18 && fontSize <= 28 && cy > ctx.pageH * 0.6) {
      out.push(rec(idx, 'ai_text', 0.6, '正文（结论句）：加粗文本位于页面下部', 'R-CONCL', {
        recommendedSemanticRole: 'body', contentType: '正文',
        suggestedPrompt: '围绕主题输入正文', suggestedConstraints: { maxChars: 80, maxLines: 2 }
      }));
      return;
    }
    // 7l) 标签/眉题/徽章：短文本 + 视觉/结构特征
    if (textLen <= 10 && fontSize <= 18) {
      const strongLabel = accentVisual || upperText || SECTION_WORDS.test(text) || /[:：]$/.test(text);
      if (strongLabel) {
        out.push(rec(idx, 'ai_text', 0.6, '短文本 + 视觉/结构特征（' + (upperText ? '全大写' : accentVisual ? '加粗/居中/彩色' : '章节词/冒号') + '），建议归为「不指定」', 'R-LABEL', {
          recommendedSemanticRole: 'other'
        }));
      } else {
        out.push(rec(idx, 'ai_text', 0.55, '短文本小字号，像标签/眉题，建议归为「不指定」', 'R-TAG', {
          recommendedSemanticRole: 'other'
        }));
      }
      return;
    }
    // 7m) 兜底：信号不足 → 不自动应用，显示信号不足
    out.push(rec(idx, 'ai_text', 0.45, '信号不足：特征不明显，建议归为「不指定」或手动选择角色', 'R-FALLBACK', {
      recommendedSemanticRole: 'other', contentType: undefined,
      suggestedPrompt: undefined, suggestedConstraints: undefined
    }));
  });

  return out;
}
