// 公式位 → OMML 转换与 PPTX 注入测试
const { test } = require('node:test');
const assert = require('node:assert');
const { latexToOmml, localTextToLatex, applyMathOmmlXml } = require('../src/mathOmml.js');

test('latexToOmml converts a LaTeX formula to OMML with fraction/subscript/superscript', () => {
  const omml = latexToOmml('T_{allow}=\\frac{\\pi[\\tau]d^3}{16K_{\\tau}\\times10^3}');
  assert.ok(omml.includes('<m:oMath'), 'has oMath root');
  assert.ok(omml.includes('<m:f>'), 'has fraction');
  assert.ok(omml.includes('<m:sSub>'), 'has subscript');
  assert.ok(omml.includes('<m:sSup>'), 'has superscript');
  assert.ok(omml.includes('xmlns:m='), 'has math namespace');
});

test('localTextToLatex normalizes unicode math paste', () => {
  const out = localTextToLatex('𝑇_allow=(𝜋[𝜏] 𝑑^3)/(16𝐾_𝜏×10^3)');
  assert.ok(out.includes('\\times'), '× → \\times');
  assert.ok(out.includes('\\pi'), '𝜋 → \\pi');
  assert.ok(out.includes('^3'), 'exponent kept');
  assert.ok(out.includes('_{allow}'), 'grouped subscript');
});

test('applyMathOmmlXml replaces the marker paragraph with oMathPara and adds m namespace', () => {
  const xml = '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:pPr marL="0" indent="0"/><a:r><a:rPr lang="en-US"/><a:t>@@OMATH:shp0@@</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
  const out = applyMathOmmlXml(xml, [{ shapeId: 'shp0', latex: 'x=\\frac{a}{b}', raw: 'x=a/b' }]);
  assert.ok(!out.includes('@@OMATH:shp0@@'), 'marker removed');
  assert.ok(out.includes('<m:oMathPara>'), 'oMathPara injected');
  assert.ok(out.includes('<m:f>'), 'fraction present');
  assert.ok(out.includes('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'), 'm namespace on p:sld');
  assert.ok(out.includes('</a:p>'), 'paragraph closed');
});

test('applyMathOmmlXml falls back to plain text when conversion is impossible', () => {
  const xml = '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>@@OMATH:shp1@@</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
  const out = applyMathOmmlXml(xml, [{ shapeId: 'shp1', latex: '', raw: '' }]);
  assert.ok(!out.includes('@@OMATH:shp1@@'), 'marker removed on failure');
  assert.ok(out.includes('<a:t xml:space="preserve"></a:t>'), 'plain empty text fallback');
});