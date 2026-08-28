// 公式位 → PowerPoint「专业型」数学（OMML）转换。
// 管线：LaTeX → tex-to-mml → MathML → mathml2omml → OMML（<m:oMath> 片段）。
// 两个库均为纯 JS、无 jsdom，可在任意路径（含中文路径）运行；已在 server 依赖中。
// 说明：向导侧负责把用户输入的公式转成 LaTeX（AI 结果优先，本地 Unicode 规范化兜底）；
//       本模块只做 LaTeX → OMML 与 PPTX 注入。

'use strict';

const texToMml = require('tex-to-mml');
const mathml2omml = require('mathml2omml');

const OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// LaTeX → OMML 字符串（<m:oMath xmlns:m=... xmlns:w=...>…</m:oMath>）。失败抛错。
function latexToOmml(latex) {
  const tex = String(latex || '').trim();
  if (!tex) throw new Error('公式为空');
  let mml;
  try {
    mml = texToMml(tex);
  } catch (e) {
    throw new Error('LaTeX 解析失败：' + (e && e.message ? e.message : String(e)));
  }
  if (!mml || typeof mml !== 'string' || mml.indexOf('<math') < 0) {
    throw new Error('公式无法转换为 MathML');
  }
  let omml;
  // 静默：mathml2omml 对不支持节点会向 console 打印噪音（如 "Type not supported: merror"），临时静默
  const saveErr = console.error, saveLog = console.log;
  console.error = () => {}; console.log = () => {};
  try {
    const fn = mathml2omml.convertMathML2OMML || mathml2omml.default || (typeof mathml2omml === 'function' ? mathml2omml : Object.values(mathml2omml)[0]);
    omml = String(fn(mml));
  } catch (e) {
    throw new Error('MathML → OMML 转换失败：' + (e && e.message ? e.message : String(e)));
  } finally {
    console.error = saveErr; console.log = saveLog;
  }
  if (!omml || omml.indexOf('<m:') < 0) throw new Error('OMML 转换结果为空');
  return omml;
}

// 本地 Unicode 公式文本 → LaTeX（兜底；向导侧也有一份等价实现，这里是 MCP/AI 直接生成场景的保险）。
function localTextToLatex(text) {
  let s = String(text || '').trim();
  if (!s) return '';
  const sup = ['⁰0','¹1','²2','³3','⁴4','⁵5','⁶6','⁷7','⁸8','⁹9'];
  const sub = ['₀0','₁1','₂2','₃3','₄4','₅5','₆6','₇7','₈8','₉9'];
  for (const p of sup) s = s.split(p[0]).join('^{' + p[1] + '}');
  for (const p of sub) s = s.split(p[0]).join('_{' + p[1] + '}');
  const map = [['×','\\times'],['÷','\\div'],['±','\\pm'],['⋅','\\cdot'],['π','\\pi'],['𝜋','\\pi'],['∑','\\sum'],['∫','\\int'],['≤','\\leq'],['≥','\\geq'],['≈','\\approx'],['≠','\\neq'],['∞','\\infty'],['Δ','\\Delta'],['√','\\sqrt{}']];
  for (const [u, rep] of map) s = s.split(u).join(rep);
  s = s.replace(/_([A-Za-z0-9τπ]{2,})/g, '_{$1}').replace(/\^([A-Za-z0-9τπ]{2,})/g, '^{$1}');
  const frac = s.match(/^\(([^()]+)\)\/\(([^()]+)\)$/);
  if (frac) s = '\\frac{' + frac[1].trim() + '}{' + frac[2].trim() + '}';
  return s;
}

// 把一页 slide XML 中的公式标记（<a:t>@@OMATH:<shapeId>@@</a:t>）替换为 OMML 专业型公式。
// formulas: [{ shapeId, latex, raw }]；转换成功 → <m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>；
// 失败 → 回退为普通文本（raw 优先，其次 latex），保证页面不出现标记残留。
function applyMathOmmlXml(xml, formulas) {
  if (!formulas || !formulas.length) return xml;
  let out = xml;
  let changed = false;
  for (const f of formulas) {
    const marker = '@@OMATH:' + f.shapeId + '@@';
    const esc = marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const tRe = new RegExp('(<a:t[^>]*>)' + esc + '(<\/a:t>)');
    const m = out.match(tRe);
    if (!m) continue;
    const tStart = m.index;
    const tEnd = tStart + m[0].length;
    // 找到包含它的 <a:p> 边界：取该 a:t 之前最近的 <a:p> 或 <a:p ...>（注意排除 <a:pPr，它同样以 <a:p 开头）
    const beforeText = out.slice(0, tStart);
    const pBare = beforeText.lastIndexOf('<a:p>');
    const pAttr = beforeText.lastIndexOf('<a:p ');
    const pStart = Math.max(pBare, pAttr);
    const pEnd = out.indexOf('</a:p>', tEnd);
    if (pStart < 0 || pEnd < 0) continue;
    let para = out.slice(pStart, pEnd + 6);
    // 保留 a:pPr（对齐等），其余（a:r/a:br/a:endParaRPr）全部替换为数学内容
    let pPr = '';
    const pPrM = para.match(/<a:pPr[^>]*>.*?<\/a:pPr>/s) || para.match(/<a:pPr[^>]*\/>/);
    if (pPrM) pPr = pPrM[0];
    let inner;
    try {
      const omml = latexToOmml(f.latex || f.raw || '');
      // 提取 <m:oMath>…</m:oMath> 主体（含其自带命名空间声明，可整体独立嵌入）
      const body = omml.match(/<m:oMath[^>]*>.*<\/m:oMath>/s);
      inner = '<m:oMathPara>' + (body ? body[0] : omml) + '<\/m:oMathPara>';
    } catch (e) {
      const plain = (f.raw && String(f.raw).trim()) || String(f.latex || '').trim() || '';
      inner = '<a:r><a:rPr lang="zh-CN" altLang="en-US" dirty="0"/><a:t xml:space="preserve">' + plain.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</a:t></a:r>';
    }
    const newPara = '<a:p>' + pPr + inner + '</a:p>';
    out = out.slice(0, pStart) + newPara + out.slice(pEnd + 6);
    changed = true;
  }
  if (changed) {
    // 确保 p:sld 根元素声明 m 命名空间（oMathPara 里的 m: 前缀才能解析）
    if (out.indexOf('xmlns:m="' + OMML_NS + '"') < 0) {
      out = out.replace(/<p:sld[^>]*>/, (head) => {
        if (head.indexOf('xmlns:m=') >= 0) return head;
        return head.replace(/>$/, ' xmlns:m="' + OMML_NS + '" xmlns:w="' + WORD_NS + '">');
      });
    }
  }
  return out;
}

module.exports = { latexToOmml, localTextToLatex, applyMathOmmlXml };