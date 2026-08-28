const { test } = require('node:test');
const assert = require('node:assert');
const JSZip = require('jszip');
const { buildSlideBase64 } = require('../src/slideBuilder.js');

const MINIMAL_TEMPLATE = {
  schemaVersion: 1,
  name: 't',
  slideSize: { width: 13.33, height: 7.5 },
  shapes: [
    { id: 's1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { font: 'Arial', size: 18, color: '#FF0000', align: 'Center' } },
    { id: 's2', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 3, width: 4, height: 3 } }
  ]
};

async function unzip(base64) {
  return JSZip.loadAsync(Buffer.from(base64, 'base64'));
}

test('buildSlideBase64 returns a valid pptx with text and image', async () => {
  const base64 = await buildSlideBase64({
    template: MINIMAL_TEMPLATE,
    imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=',
    texts: { s1: '你好' },
    vars: {}
  });
  assert.equal(typeof base64, 'string');
  assert.ok(base64.length > 0);

  const zip = await unzip(base64);
  assert.ok(zip.file('[Content_Types].xml'), 'missing [Content_Types].xml');
  assert.ok(zip.file('ppt/slides/slide1.xml'), 'missing slide1.xml');

  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('<a:t>你好</a:t>'), 'slide1.xml should contain the ai_text value');
  assert.ok(slideXml.includes('<a:blip'), 'slide1.xml should reference an image');

  const mediaFiles = Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/'));
  assert.ok(mediaFiles.length > 0, 'ppt/media should contain an image file');
});

test('buildSlideBase64 with images record: each ai_image shape gets its own picture (multi-image)', async () => {
  const tpl = {
    schemaVersion: 1, name: 'multi', slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 'imgA', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 1, width: 4, height: 3 } },
      { id: 'imgB', type: 'picture', role: 'ai_image', bounds: { left: 7, top: 1, width: 4, height: 3 } }
    ]
  };
  const base64 = await buildSlideBase64({
    template: tpl,
    images: { imgA: 'data:image/png;base64,AAAA', imgB: 'data:image/png;base64,BBBB' },
    texts: {}, vars: {}
  });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  // 两个 p:pic 块、两张媒体文件（每图独立引用）
  const pics = (slideXml.match(/<p:pic>/g) || []).length;
  assert.equal(pics, 2, 'should render 2 pictures for 2 ai_image slots');
  const media = Object.keys(zip.files).filter((n) => n.startsWith('ppt/media/') && !n.endsWith('/'));
  assert.equal(media.length, 2, 'should embed 2 distinct media files');
  // 两张图内容不同（AA vs BB）
  const a = await zip.file(media[0]).async('uint8array');
  const b = await zip.file(media[1]).async('uint8array');
  assert.notDeepEqual(Array.from(a), Array.from(b), 'two slots should hold different image data');
  // 兼容旧调用：imageDataUrl 单值 → 所有图片位共用同一张图（内容一致）
  const base642 = await buildSlideBase64({ template: tpl, imageDataUrl: 'data:image/png;base64,CCCC', texts: {}, vars: {} });
  const zip2 = await unzip(base642);
  const media2 = Object.keys(zip2.files).filter((n) => n.startsWith('ppt/media/') && !n.endsWith('/'));
  const a2 = await zip2.file(media2[0]).async('uint8array');
  const b2 = await zip2.file(media2[1]).async('uint8array');
  assert.deepEqual(Array.from(a2), Array.from(b2), 'legacy single imageDataUrl: all slots share identical image');
});

test('buildSlideBase64 writes template picture background as real slide background', async () => {
  const template = {
    schemaVersion: 1,
    name: 'bg',
    slideSize: { width: 13.33, height: 7.5 },
    background: { type: 'picture', imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=' },
    shapes: []
  };
  const base64 = await buildSlideBase64({ template, texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');

  assert.ok(slideXml.includes('<p:bg>'), 'picture background should be written into p:bg');
  assert.equal((slideXml.match(/<p:pic>/g) || []).length, 0, 'background must not be a normal picture shape');
});

test('buildSlideBase64 writes template solid background as real slide background', async () => {
  const template = {
    schemaVersion: 1,
    name: 'bg',
    slideSize: { width: 13.33, height: 7.5 },
    background: { type: 'solid', color: '#123456' },
    shapes: []
  };
  const base64 = await buildSlideBase64({ template, texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');

  assert.ok(slideXml.includes('<p:bg>'), 'solid background should be written into p:bg');
  assert.ok(slideXml.includes('val="123456"'), 'solid background color should be preserved');
  assert.equal((slideXml.match(/<p:sp>/g) || []).length, 0, 'background must not be a normal shape');
});

test('align values map to pptx algn attributes', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    shapes: [
      { id: 'c', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { align: 'Center' } },
      { id: 'l', type: 'text', role: 'ai_text', bounds: { left: 1, top: 2, width: 5, height: 1 }, textStyle: { align: 'left' } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: null, texts: { c: 'a', l: 'b' }, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('algn="ctr"'), 'Center should map to algn="ctr"');
  assert.ok(slideXml.includes('algn="l"'), 'left should map to algn="l"');
});

test('underline, strike and superscript are written into slide xml', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    shapes: [
      { id: 'u1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { underline: true, strikethrough: true } },
      { id: 'u2', type: 'text', role: 'ai_text', bounds: { left: 1, top: 2, width: 5, height: 1 }, textStyle: { doubleStrikethrough: true, superscript: true } },
      { id: 'u3', type: 'text', role: 'ai_text', bounds: { left: 1, top: 3, width: 5, height: 1 }, textStyle: { align: 'Distributed' } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: null, texts: { u1: 'a', u2: 'b', u3: 'c' }, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('u="sng"'), 'underline should be written');
  assert.ok(slideXml.includes('strike="sngStrike"'), 'single strike should be written');
  assert.ok(slideXml.includes('strike="dblStrike"'), 'double strike should be written');
  assert.ok(!slideXml.includes('strike="none"') && !slideXml.includes('strike="single"') && !slideXml.includes('strike="double"'), 'invalid strike enum values must never be written (PowerPoint rejects the file)');
  assert.ok(slideXml.includes('baseline="30000"'), 'superscript should be written');
  assert.ok(slideXml.includes('algn="just"'), 'Distributed should map to justify');
});

test('control characters in text are stripped so build never fails', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    shapes: [
      { id: 'x1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: {} }
    ]
  };
  const base64 = await buildSlideBase64({
    template,
    imageDataUrl: null,
    texts: { x1: '正常\u0000\u0001\u000B\u000C\u001F文字\uFFFF' },
    vars: {}
  });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('<a:t>正常文字</a:t>'), 'invalid xml chars should be stripped');
});

test('zero-height line shape gets a minimum thickness so PowerPoint import never hangs', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    slideSize: { width: 10, height: 5.625 },
    shapes: [
      { id: 'l1', type: 'line', role: 'fixed', shapeType: 'line', bounds: { left: 1, top: 1, width: 5, height: 0 }, line: { color: '000000', weight: 1, visible: true } },
      { id: 'l2', type: 'line', role: 'fixed', shapeType: 'line', bounds: { left: 1, top: 2, width: 0, height: 2 }, line: { color: '000000', weight: 1, visible: true } },
      { id: 't1', type: 'text', role: 'fixed', content: 'x', bounds: { left: 1, top: 3, width: 0, height: 0 }, textStyle: {} }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: null, texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  // 所有内容形状（排除 grpSpPr 组容器）的 ext 都不能为 0
  const blocks = slideXml.match(/<p:(?:cxnSp|sp)>[\s\S]*?<\/p:(?:cxnSp|sp)>/g) || [];
  assert.ok(blocks.length >= 3, 'expected 3 content shapes');
  for (const b of blocks) {
    const ext = b.match(/<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"/);
    assert.ok(ext, 'shape should have ext');
    assert.ok(Number(ext[1]) > 0 && Number(ext[2]) > 0, 'shape ext must be non-zero, got ' + ext[1] + 'x' + ext[2]);
  }
});

test('manual_var text value is written into the slide', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    shapes: [
      { id: 'v1', type: 'text', role: 'manual_var', bounds: { left: 1, top: 1, width: 5, height: 1 }, varName: 'name' }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: null, texts: {}, vars: { v1: '变量值' } });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('<a:t>变量值</a:t>'), 'manual_var value should be written');
});

test('arrow shapes get a default thickness and direction follows bounds', async () => {
  const template = {
    ...MINIMAL_TEMPLATE,
    slideSize: { width: 10, height: 5.625 },
    shapes: [
      { id: 'a1', type: 'line', role: 'fixed', shapeType: 'downArrow', bounds: { left: 5, top: 4, width: 0, height: 0.866 }, line: { color: 'CF6962', weight: 1.5, visible: true } },
      { id: 'a2', type: 'line', role: 'fixed', shapeType: 'rightArrow', bounds: { left: 1, top: 1, width: 3, height: 0 }, line: { color: '000000', weight: 1, visible: true } },
      { id: 'a3', type: 'line', role: 'fixed', shapeType: 'rightArrow', bounds: { left: 1, top: 3, width: 0, height: 2 }, line: { color: '000000', weight: 1, visible: true } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: null, texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('prst="downArrow"'), 'vertical arrow should be downArrow');
  assert.ok(slideXml.includes('prst="rightArrow"'), 'horizontal arrow should be rightArrow');
  const downExt = slideXml.match(/prst="downArrow"[\s\S]*?<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"/);
  assert.ok(downExt, 'downArrow should have a bounding box');
  assert.ok(Number(downExt[1]) > 0, 'downArrow width should be > 0 (thickness applied)');
  // 旧模板：shapeType=rightArrow 但 bounds 为竖长 → 自动修正为 downArrow
  const a3Ext = slideXml.match(/prst="rightArrow"[\s\S]*?<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"/);
  assert.ok(Number(a3Ext[2]) > 0, 'rightArrow height should be > 0');
});
test('buildSlideBase64 applies full imageStyle (spPrXml + blip effects)', async () => {
  const template = {
    schemaVersion: 1, name: 't',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 'p1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 1, width: 4, height: 3 },
        imageStyle: {
          spPrXml: '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>' +
            '<a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>' +
            '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>' +
            '<a:softEdge rad="50800"/></a:effectLst>',
          blipAttrs: 'cstate="print"',
          blipKids: '<a:alphaModFix amt="70000"/>',
          srcRect: { l: 1000, t: 0, r: 2000, b: 500 }
        } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=', texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('prst="roundRect"'), 'geometry kept from spPrXml');
  assert.ok(slideXml.includes('<a:ln w="19050"'), 'border kept');
  assert.ok(slideXml.includes('outerShdw'), 'shadow kept');
  assert.ok(slideXml.includes('<a:softEdge rad="50800"/>'), 'soft edge kept');
  assert.ok(slideXml.includes('cstate="print"'), 'blip attr kept');
  assert.ok(slideXml.includes('<a:alphaModFix amt="70000"/>'), 'blip kid kept');
  assert.ok(slideXml.includes('<a:srcRect l="1000" t="0" r="2000" b="500"/>'), 'crop kept');
});

test('buildSlideBase64 writes ea font into generated text (eaFont != font)', async () => {
  const template = {
    schemaVersion: 1, name: 't',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 's1', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 },
        textStyle: { font: 'Calibri', eaFont: '微软雅黑', size: 18 } },
      { id: 's2', type: 'text', role: 'ai_text', bounds: { left: 1, top: 3, width: 5, height: 1 },
        textStyle: { font: 'Arial', size: 16 } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: '', texts: { s1: '标题', s2: '正文' }, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  // 第一个文本的 ea 被换成微软雅黑，latin 保持 Calibri；第二个保持 Arial
  const eaMatches = slideXml.match(/<a:ea typeface="([^"]*)"/g) || [];
  const latinMatches = slideXml.match(/<a:latin typeface="([^"]*)"/g) || [];
  assert.ok(eaMatches.some((m) => m.includes('微软雅黑')), 'ea should be 微软雅黑 for first text');
  assert.ok(latinMatches.some((m) => m.includes('Calibri')), 'latin stays Calibri');
  assert.ok(eaMatches.some((m) => m.includes('Arial')), 'second text ea stays Arial');
});

test('buildSlideBase64 applies imageStyle (roundRect + soft edge + crop) to picture', async () => {
  const template = {
    schemaVersion: 1, name: 't',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 'p1', type: 'picture', role: 'ai_image', bounds: { left: 1, top: 1, width: 4, height: 3 },
        imageStyle: { shape: 'roundRect', softEdgeEmu: 50800, srcRect: { l: 1000, t: 0, r: 2000, b: 500 } } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=', texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('prst="roundRect"'), 'geometry should become roundRect');
  assert.ok(slideXml.includes('<a:gd name="adj" fmla="val 16667"/>'), 'roundRect adj should be present');
  assert.ok(slideXml.includes('<a:softEdge rad="50800"/>'), 'soft edge should be applied');
  assert.ok(slideXml.includes('<a:srcRect l="1000" t="0" r="2000" b="500"/>'), 'crop srcRect should be applied');
});

test('buildSlideBase64 keeps plain picture untouched when no imageStyle', async () => {
  const base64 = await buildSlideBase64({ template: MINIMAL_TEMPLATE, imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=', texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(!slideXml.includes('roundRect'), 'no geometry override');
  assert.ok(!slideXml.includes('softEdge'), 'no soft edge');
  assert.ok(!slideXml.includes('srcRect'), 'no crop');
  assert.ok(slideXml.includes('<a:prstGeom prst="rect">'), 'picture stays plain rect');
});

test('buildSlideBase64 applies imageStyle to fixed picture with dataURL content', async () => {
  const template = {
    schemaVersion: 1, name: 't',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      { id: 'f1', type: 'picture', role: 'fixed', bounds: { left: 1, top: 1, width: 4, height: 3 },
        content: 'data:image/png;base64,iVBORw0KGgo=', imageStyle: { shape: 'ellipse' } }
    ]
  };
  const base64 = await buildSlideBase64({ template, imageDataUrl: '', texts: {}, vars: {} });
  const zip = await unzip(base64);
  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(slideXml.includes('prst="ellipse"'), 'fixed picture should get ellipse geometry');
});
