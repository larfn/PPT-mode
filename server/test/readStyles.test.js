const { test } = require('node:test');
const assert = require('node:assert');
const JSZip = require('jszip');
const {
  parseShapeBlock, parseSlideShapesDetailed, aggregateTextStyle, mergeTextStyle, readStyles
} = require('../src/readStyles.js');

function shapeXml({ ph, pPr, runs }) {
  const txBody = `<p:txBody><a:bodyPr anchor="t"/><a:lstStyle/><a:p>${pPr || ''}${runs}<a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr/><p:nvPr>${ph ? `<p:ph type="${ph}"/>` : ''}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>${txBody}</p:sp>`;
}

test('parseShapeBlock reads underline/size/font/color/align', () => {
  const xml = shapeXml({
    pPr: '<a:pPr algn="ctr"/>',
    runs: '<a:r><a:rPr lang="en-US" sz="2400" b="1" u="sng" dirty="0"><a:solidFill><a:srgbClr val="4B3959"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Hi</a:t></a:r>'
  });
  const block = parseShapeBlock(xml);
  assert.strictEqual(block.name, 'T1');
  const style = aggregateTextStyle(block.paragraphs);
  assert.deepStrictEqual(style, {
    align: 'center', underline: true, font: 'Arial', size: 24, bold: true, color: '#4B3959'
  });
});

test('strike/superscript/subscript are detected', () => {
  const xml = shapeXml({
    runs: [
      '<a:r><a:rPr strike="dblStrike" baseline="-25000"/><a:t>a</a:t></a:r>',
      '<a:r><a:rPr baseline="30000"/><a:t>b</a:t></a:r>'
    ].join('')
  });
  const style = aggregateTextStyle(parseShapeBlock(xml).paragraphs);
  assert.strictEqual(style.doubleStrikethrough, true);
  assert.strictEqual(style.superscript, true);
  assert.strictEqual(style.subscript, true);
});

test('mergeTextStyle prefers live Office.js values, fills gaps from XML', () => {
  const xml = { align: 'center', underline: true, font: 'Arial', size: 24, bold: true, color: '#112233' };
  assert.deepStrictEqual(mergeTextStyle({}, xml), { align: 'center', underline: true, font: 'Arial', size: 24, bold: true, color: '#112233' });
  assert.strictEqual(mergeTextStyle({ align: 'center', underline: false }, xml).align, 'center');
  assert.strictEqual(mergeTextStyle({ align: 'center' }, xml).align, 'center');
  // 字体：XML 值（含 latin）优先于 Office.js（Office.js 只能拿 latin，可能不符显示字体）
  assert.strictEqual(mergeTextStyle({ font: 'Keep' }, xml).font, 'Arial');
  assert.strictEqual(mergeTextStyle({ align: 'left' }, xml).align, 'center');
  assert.strictEqual(mergeTextStyle(undefined, null), undefined);
});

test('mergeTextStyle: XML ea font overrides Office.js latin font (Chinese display font)', () => {
  const xml = { font: '微软雅黑', eaFont: '微软雅黑' };
  assert.strictEqual(mergeTextStyle({ font: 'Calibri' }, xml).font, '微软雅黑');
  assert.strictEqual(mergeTextStyle({}, xml).font, '微软雅黑');
});

test('aggregateTextStyle: font=latin + eaFont=ea separately; latin-only falls back', () => {
  // latin=Arial + ea=微软雅黑 → font=Arial（英文）、eaFont=微软雅黑（中文）
  const xml1 = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>';
  const b1 = parseShapeBlock(xml1);
  assert.strictEqual(aggregateTextStyle(b1.paragraphs).font, 'Arial');
  assert.strictEqual(aggregateTextStyle(b1.paragraphs).eaFont, '微软雅黑');
  // 只有 latin → font=latin，无 eaFont
  const xml2 = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>';
  const b2 = parseShapeBlock(xml2);
  assert.strictEqual(aggregateTextStyle(b2.paragraphs).font, 'Arial');
  assert.strictEqual(aggregateTextStyle(b2.paragraphs).eaFont, undefined);
  // 只有 ea → font 取 ea 兜底（旧模板单字体字段兼容）
  const xml3 = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:ea typeface="宋体"/></a:rPr><a:t>中</a:t></a:r></a:p></p:txBody></p:sp>';
  const b3 = parseShapeBlock(xml3);
  assert.strictEqual(aggregateTextStyle(b3.paragraphs).font, '宋体');
  assert.strictEqual(aggregateTextStyle(b3.paragraphs).eaFont, '宋体');
});

test('readStyles end-to-end: slide placeholder inherits center from layout, underline from slide', async () => {
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr u="sng" sz="2800"/><a:t>Title</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Layout"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapeXml({ ph: 'ctrTitle', pPr: '<a:pPr algn="ctr"/>', runs: '<a:r><a:rPr/><a:t/></a:r>' })}</p:spTree></p:cSld></p:sldLayout>`;
  const masterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld></p:sldMaster>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slideXml);
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml);

  const zipBase64 = await zip.generateAsync({ type: 'base64' });
  const shapes = [{ name: 'T1', bounds: { left: 1, top: 1, width: 4, height: 1 }, textStyle: { align: 'left', size: 28 } }];
  const { styles } = await readStyles({ zipBase64, slideIndex: 1, shapes });
  assert.ok(styles[0]);
  assert.strictEqual(styles[0].align, 'center');
  assert.strictEqual(styles[0].underline, true);
  assert.strictEqual(styles[0].size, 28);
});

test('parseShapeBlock reads lstStyle default alignment (layout placeholder style)', () => {
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="PH"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="10"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"/></a:lstStyle><a:p><a:r><a:rPr/><a:t>ALPHA</a:t></a:r></a:p></p:txBody></p:sp>`;
  const block = parseShapeBlock(xml);
  assert.strictEqual(block.lstAlign, 'center');
  assert.strictEqual(block.anchor, 'middle');
  const style = aggregateTextStyle(block.paragraphs, block.lstAlign);
  assert.strictEqual(style.align, 'center');
});

test('aggregateTextStyle returns align for run-less placeholder with lstAlign', () => {
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="PH"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="10"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"/></a:lstStyle><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>`;
  const block = parseShapeBlock(xml);
  const style = aggregateTextStyle(block.paragraphs, block.lstAlign);
  assert.strictEqual(style.align, 'center');
});

test('readStyles matches slide by position when shapes have no name', async () => {
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapeXml({ ph: 'ctrTitle', pPr: '<a:pPr algn="ctr"/>', runs: '<a:r><a:rPr u="sng" sz="2800"/><a:t>Title</a:t></a:r>' })}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Layout"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapeXml({ ph: 'ctrTitle', pPr: '<a:pPr algn="ctr"/>', runs: '<a:r><a:rPr/><a:t/></a:r>' })}</p:spTree></p:cSld></p:sldLayout>`;
  const masterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld></p:sldMaster>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slideXml);
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml);
  const zipBase64 = await zip.generateAsync({ type: 'base64' });
  // 无 name（模拟旧模板重载后再次保存），仅靠位置匹配
  const shapes = [{ name: '', bounds: { left: 1, top: 1, width: 4, height: 1 }, textStyle: { align: 'left', size: 28 } }];
  const { styles } = await readStyles({ zipBase64, slideIndex: 1, shapes });
  assert.ok(styles[0]);
  assert.strictEqual(styles[0].align, 'center');
  assert.strictEqual(styles[0].underline, true);
});

test('readStyles inherits align from layout lstStyle and valign from layout anchor', async () => {
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr u="sng" sz="2800"/><a:t>Title</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Layout"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="PH"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"/></a:lstStyle><a:p><a:r><a:rPr/><a:t/></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`;
  const masterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"/></p:titleStyle><p:bodyStyle><a:lvl1pPr algn="l"/></p:bodyStyle></p:txStyles></p:sldMaster>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slideXml);
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml);
  const zipBase64 = await zip.generateAsync({ type: 'base64' });
  const shapes = [{ name: 'T1', bounds: { left: 1, top: 1, width: 4, height: 1 }, textStyle: { align: 'left', size: 28, valign: 'top' } }];
  const { styles } = await readStyles({ zipBase64, slideIndex: 1, shapes });
  assert.ok(styles[0]);
  // 页面无显式对齐 → 继承 layout lstStyle algn=ctr
  assert.strictEqual(styles[0].align, 'center');
  // 页面无显式垂直对齐 → 继承 layout bodyPr anchor=ctr（Office.js 的 top 被 XML 修正为 middle）
  assert.strictEqual(styles[0].valign, 'middle');
  assert.strictEqual(styles[0].underline, true);
});
test('parseShapeBlock reads picture style (geometry / soft edge / crop)', () => {
  const picXml = '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Pic 1"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="1000" t="0" r="2000" b="500"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' +
    '<a:effectLst><a:softEdge rad="50800"/></a:effectLst></p:spPr></p:pic>';
  const block = parseShapeBlock(picXml);
  assert.strictEqual(block.imgPrst, 'roundRect');
  assert.strictEqual(block.imgSoftEdgeEmu, 50800);
  assert.deepEqual(block.imgSrcRect, { l: 1000, t: 0, r: 2000, b: 500 });
});

test('mergeTextStyle carries eaFont through for generation', () => {
  const xml = { font: 'Calibri', eaFont: '微软雅黑' };
  const merged = mergeTextStyle({ font: 'Calibri' }, xml);
  assert.strictEqual(merged.font, 'Calibri', 'latin stays in font');
  assert.strictEqual(merged.eaFont, '微软雅黑', 'ea carried separately');
});

test('parseShapeBlock reads full picture style: spPrXml (except xfrm) + blip effects', () => {
  const picXml = '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Pic 1"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2" cstate="print"><a:alphaModFix amt="70000"/></a:blip>' +
    '<a:srcRect l="1000" t="0" r="2000" b="500"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' +
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>' +
    '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>' +
    '<a:softEdge rad="50800"/></a:effectLst>' +
    '<a:scene3d><a:camera prst="perspectiveContrastingRightFacing"/></a:scene3d>' +
    '<a:sp3d><a:bevelT w="50800" h="25400"/></a:sp3d>' +
    '</p:spPr></p:pic>';
  const block = parseShapeBlock(picXml);
  // 完整 spPr（无 xfrm）
  assert.ok(block.imgSpPrXml, 'spPrXml should be extracted');
  assert.ok(!block.imgSpPrXml.includes('<a:xfrm'), 'xfrm must be excluded');
  assert.ok(block.imgSpPrXml.includes('prst="roundRect"'), 'geometry kept');
  assert.ok(block.imgSpPrXml.includes('outerShdw'), 'shadow kept');
  assert.ok(block.imgSpPrXml.includes('<a:softEdge rad="50800"/>'), 'soft edge kept');
  assert.ok(block.imgSpPrXml.includes('<a:scene3d'), '3d scene kept');
  assert.ok(block.imgSpPrXml.includes('<a:sp3d'), '3d shape kept');
  assert.ok(block.imgSpPrXml.includes('<a:ln w="19050"'), 'border kept');
  // blip 效果：属性去 r:embed、子元素保留
  assert.strictEqual(block.imgBlipAttrs, 'cstate="print"');
  assert.ok(block.imgBlipKids.includes('alphaModFix'), 'blip kids kept');
  assert.deepEqual(block.imgSrcRect, { l: 1000, t: 0, r: 2000, b: 500 });
});

test('readStyles returns imageStyles for picture shapes (e2e)', async () => {
  const slideXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Pic 1"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="1000" t="0" r="2000" b="500"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' +
    '<a:effectLst><a:softEdge rad="50800"/></a:effectLst></p:spPr></p:pic>' +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slideXml);
  const zipBase64 = await zip.generateAsync({ type: 'base64' });
  const shapes = [{ name: 'Pic 1', type: 'picture', bounds: { left: 1, top: 1, width: 4, height: 1 } }];
  const { styles, imageStyles } = await readStyles({ zipBase64, slideIndex: 1, shapes });
  assert.ok(imageStyles[0], 'imageStyles should be returned');
  assert.strictEqual(imageStyles[0].shape, 'roundRect');
  assert.strictEqual(imageStyles[0].softEdgeEmu, 50800);
  assert.deepEqual(imageStyles[0].srcRect, { l: 1000, t: 0, r: 2000, b: 500 });
  // 非图片形状 → imageStyles 为 null
  assert.strictEqual(styles[0], undefined);
  const shapes2 = [{ name: 'Pic 1', bounds: { left: 1, top: 1, width: 4, height: 1 } }];
  const r2 = await readStyles({ zipBase64, slideIndex: 1, shapes: shapes2 });
  assert.strictEqual(r2.imageStyles[0], null, 'no imageStyle without type=picture');
});

test('readStyles inherits placeholder font from layout defRPr and master txStyles (e2e)', async () => {
  // 页面占位符 run 无任何显式字体；layout 占位符 lstStyle lvl1 defRPr 定义 latin=Calibri + ea=微软雅黑
  const slideXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2800"/><a:t>标题</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const layoutXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld name="Layout"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:latin typeface="Calibri"/><a:ea typeface="微软雅黑"/></a:defRPr></a:lvl1pPr></a:lstStyle>' +
    '<a:p><a:r><a:rPr/><a:t/></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sldLayout>';
  const masterXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld name="Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '</p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slideXml);
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml);
  const zipBase64 = await zip.generateAsync({ type: 'base64' });
  // Office.js 只拿到 latin（Calibri），页面 run 无字体 → 应继承 layout 的 ea=微软雅黑
  const shapes = [{ name: 'Title 1', bounds: { left: 1, top: 1, width: 4, height: 1 }, textStyle: { font: 'Calibri', size: 28 } }];
  const { styles } = await readStyles({ zipBase64, slideIndex: 1, shapes });
  assert.ok(styles[0]);
  // layout 的 latin=Calibri 进 font，ea=微软雅黑 单独进 eaFont
  assert.strictEqual(styles[0].font, 'Calibri');
  assert.strictEqual(styles[0].eaFont, '微软雅黑');
  assert.strictEqual(styles[0].align, 'center', 'align inherited from layout lstStyle');
});
