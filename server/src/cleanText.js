// AI 输出净化器：把 LLM 返回的「文章」洗成适合 PPT 的「贴片」文本。
// 所有 /api/text/generate 的返回都经过这里（textService.generateText 统一调用），
// 因此向导逐段生成 / 翻译 / 表格 JSON / 套版生成本页文字 全部自动生效。
//
// 规则分层（保证幂等、对任何调用安全）：
//  - 始终生效：剥离 Markdown 语法痕迹（###、**、`、代码块围栏、引用、分隔线、图片/链接）、规整换行
//  - plain=true（纯文本模式）：额外剥离列表编号/符号（1. / 1、 / - / * / + / · / •）
//  - maxChars / maxLines > 0：超限截断（0 或缺失 = 不限制）
// 注意：表格 AI 返回的是 JSON 数组，调用方不应传 plain/maxChars，否则会洗掉合法的
// 「1. 步骤」类单元格内容、甚至截断 JSON；本函数对 JSON 字符串本身是安全的（不做结构破坏）。

'use strict';

/**
 * @param {string} text AI 原始返回
 * @param {{ plain?: boolean; maxChars?: number; maxLines?: number }} [opts]
 * @returns {string} 清洗后的纯文本
 */
function cleanAiText(text, opts = {}) {
  if (typeof text !== 'string') return text;
  let t = text;

  // 1) 代码块围栏（```lang ... ```）：去掉围栏，保留内部内容，交给后续逐行处理
  t = t.replace(/```[\s\S]*?```/g, (m) =>
    m.replace(/^```[^\n]*\n?/, '').replace(/\n?```\s*$/, ''));

  // 2) 逐行剥离行级 Markdown 语法
  const lines = t.split('\n');
  const out = [];
  for (const raw of lines) {
    let line = raw.replace(/\r$/, '');
    // 残余代码块围栏行（未闭合等）
    if (/^\s*```[^\n]*$/.test(line)) line = '';
    // 标题：### 标题 → 标题
    line = line.replace(/^\s{0,3}#{1,6}\s+/, '');
    // 引用：> xxx → xxx
    line = line.replace(/^\s{0,3}>\s?/, '');
    // 分隔线：--- / *** / ___ 整行 → 空行
    if (/^\s*([-*_])\s*(\1\s*){2,}$/.test(line)) line = '';
    // 图片：![alt](url) → 删除；链接：[text](url) → text
    line = line.replace(/!\[[^\]]*\]\([^)]*\)/g, '');
    line = line.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1');
    // 行内代码 `x` → x
    line = line.replace(/`([^`]*)`/g, '$1');
    // 粗斜体 ***x*** → x；粗体 **x** / __x__ → x；删除线 ~~x~~ → x
    line = line.replace(/\*\*\*([^*]+)\*\*\*/g, '$1');
    line = line.replace(/\*\*([^*]+)\*\*/g, '$1');
    line = line.replace(/__([^_]+)__/g, '$1');
    line = line.replace(/~~([^~]+)~~/g, '$1');
    // 斜体 *x* / _x_（要求成对且边界非字母数字/下划线，避免误伤 5*3、snake_case）
    line = line.replace(/(^|[^*\w])\*([^*\n]+)\*(?!\*)/g, '$1$2');
    line = line.replace(/(^|[^_\w])_([^_\n]+)_(?!_)/g, '$1$2');
    // 纯文本模式：剥离列表编号/符号（1. / 1、 / 1) / - / * / + / · / •）
    if (opts.plain === true) {
      line = line.replace(/^\s{0,3}(?:[-*+•·]|\d{1,3}[.)、])\s+/, '');
    }
    out.push(line);
  }
  t = out.join('\n');

  // 3) 规整：每行去首尾空白、连续 3+ 空行压成 2、去掉首尾空行
  t = t
    .split('\n')
    .map((l) => l.trim())
    .join('\n');
  t = t.replace(/\n{3,}/g, '\n\n').replace(/^\n+|\n+$/g, '');

  // 4) 控制字符 + 换行符统一（与 slideBuilder.normalizeBreaks 保持一致）
  t = t
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');

  // 5) 长度约束（0 / 缺失 = 不限制）
  if (opts.maxLines && opts.maxLines > 0) {
    const ls = t.split('\n');
    if (ls.length > opts.maxLines) t = ls.slice(0, opts.maxLines).join('\n');
  }
  if (opts.maxChars && opts.maxChars > 0 && t.length > opts.maxChars) {
    t = t.slice(0, opts.maxChars) + '…';
  }
  return t.trim();
}

module.exports = { cleanAiText };
