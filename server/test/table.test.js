// 表格（P0-B）测试：XML 回读解析 + 生成重建（尺寸自适应/合并格/样式）
const { test } = require('node:test');
const assert = require('node:assert');
const JSZip = require('jszip');
const { parseTableBlocks, parseTableBlock } = require('../src/readStyles.js');
const { readStylesFromZip } = require('../src/readStyles.js');
const { buildSlideBase64, buildDeckBase64, applyTableMergesXml } = require('../src/slideBuilder.js');
// 表格自动排版引擎（Node 24 type-stripping 直接 require addin 的纯 TS）
const { fitTableLayout, cellsFromGrid } = require('../../addin/src/lib/tableModel.ts');

// 构造含表格的 slide XML（2×3 表：首行合并 2 列、一个填充格、字体信息）
const TABLE_XML = `<p:graphicFrame>
  <p:cNvPr id="4" name="表格 1"/>
  <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1143000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
  <a:tbl>
    <a:tblPr firstRow="1" bandRow="1"/>
    <a:tblGrid>
      <a:gridCol w="1371600"/><a:gridCol w="1371600"/><a:gridCol w="914400"/>
    </a:tblGrid>
    <a:tr h="457200">
      <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1600" lang="zh-CN"><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>季度指标</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:tcPr></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>单位</a:t></a:r></a:p></a:txBody><a:tcPr><a:anchor ctr="1"/></a:tcPr></a:tc>
    </a:tr>
    <a:tr h="342900">
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>营收</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>18.5%</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>万元</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
    </a:tr>
  </a:tbl>
  </a:graphicData></a:graphic>
</p:graphicFrame>`;

test('parseTableBlocks: 解析行列/合并格/文字/字体/填充/bounds（EMU→英寸）', () => {
  const tables = parseTableBlocks(TABLE_XML);
  assert.equal(tables.length, 1);
  const { bounds, table } = tables[0];
  assert.equal(tables[0].name, '表格 1', 'cNvPr 名称');
  assert.ok(Math.abs(bounds.left - 1) < 0.01, 'left 1 英寸');
  assert.ok(Math.abs(bounds.width - 4) < 0.01, 'width 4 英寸');
  assert.equal(table.rows, 2);
  assert.equal(table.cols, 3);
  assert.equal(table.colWidths.length, 3);
  assert.equal(table.rowHeights.length, 2);
  assert.equal(table.tblPr.firstRow, true);
  assert.equal(table.tblPr.bandRow, true);
  // 合并格
  const c00 = table.cells.find((c) => c.row === 0 && c.col === 0);
  assert.equal(c00.colspan, 2, '首行第一格跨 2 列');
  assert.equal(c00.text, '季度指标');
  assert.equal(c00.textStyle.bold, true);
  assert.equal(c00.textStyle.size, 16);
  assert.equal(c00.textStyle.font, 'Arial');
  assert.equal(c00.textStyle.eaFont, '微软雅黑');
  assert.equal(c00.textStyle.color, '1F4E79');
  assert.equal(c00.fill, '1F4E79');
  // 正常格
  const c10 = table.cells.find((c) => c.row === 1 && c.col === 0);
  assert.equal(c10.text, '营收');
});

test('readStylesFromZip: 返回 tables（zip 模式）', async () => {
  const zip = new JSZip();
  zip.file('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld>\n' + TABLE_XML + '\n</p:sld>');
  zip.file('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  const r = await readStylesFromZip({ zip, slideIndex: 1, shapes: [] });
  assert.equal(r.tables.length, 1);
  assert.equal(r.tables[0].table.rows, 2);
});

const TABLE_TEMPLATE = {
  schemaVersion: 1, name: 't',
  slideSize: { width: 13.33, height: 7.5 },
  shapes: [
    {
      id: 'tbl1', type: 'table', role: 'fixed',
      bounds: { left: 1, top: 1, width: 4, height: 1.25 },
      table: {
        rows: 2, cols: 3,
        colWidths: [1.5, 1.5, 1], rowHeights: [0.5, 0.375],
        tblPr: { firstRow: true },
        cells: [
          { row: 0, col: 0, colspan: 2, text: '季度指标', textStyle: { bold: true, size: 16, font: 'Arial', eaFont: '微软雅黑', color: '1F4E79' }, fill: '1F4E79' },
          { row: 0, col: 2, text: '单位' },
          { row: 1, col: 0, text: '营收' },
          { row: 1, col: 1, text: '18.5%' },
          { row: 1, col: 2, text: '万元' }
        ]
      }
    }
  ]
};

test('buildSlideBase64: 表格重建为 a:tbl（文字/合并/填充/尺寸自适应）', async () => {
  const b64 = await buildSlideBase64({ template: TABLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {} });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('<a:tbl>'), '应生成表格');
  assert.ok(xml.includes('季度指标') && xml.includes('18.5%'), '单元格文字');
  // MVP：不传 colspan（pptxgenjs 3.x 合并输出不可靠），合并文字保留在主格
  assert.ok(xml.includes('季度指标') && xml.includes('1F4E79'), '主格文字与填充');
  assert.ok(xml.includes('1F4E79'), '填充色');
  // 尺寸：x/y 与 bounds 一致（914400 = 1 英寸）
  assert.ok(xml.includes('x="914400"'), '表格 x = 1 英寸');
});

test('buildDeckBase64: 多页含表格页正常', async () => {
  const b64 = await buildDeckBase64([
    { template: TABLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {} },
    { template: { ...TABLE_TEMPLATE, shapes: [{ id: 't0', type: 'text', role: 'ai_text', bounds: { left: 1, top: 1, width: 5, height: 1 }, textStyle: { size: 24 } }] }, imageDataUrl: '', texts: { t0: '第二页' }, vars: {} }
  ]);
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const s1 = await zip.file('ppt/slides/slide1.xml').async('string');
  const s2 = await zip.file('ppt/slides/slide2.xml').async('string');
  assert.ok(s1.includes('<a:tbl>'), '第 1 页含表格');
  assert.ok(!s2.includes('<a:tbl>'), '第 2 页无表格');
  assert.ok(s2.includes('第二页'));
});

// —— P0-C：表格位样式保真（列宽/边框/边距/表样式回写）与 tableData（逐格/粘贴 CSV/AI 数据） ——

const STYLE_TEMPLATE = {
  schemaVersion: 1, name: 't',
  slideSize: { width: 13.33, height: 7.5 },
  shapes: [{
    id: 'tbl1', type: 'table', role: 'table',
    bounds: { left: 1, top: 1, width: 4, height: 1.25 },
    table: {
      rows: 1, cols: 2,
      colWidths: [1.5, 2.5], rowHeights: [1.25],
      tblPr: { firstRow: true, bandRow: true, tableStyleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' },
      cells: [
        { row: 0, col: 0, text: '左', border: { top: { width: 2, color: 'FF0000' }, bottom: { width: 1, color: '00FF00' } }, margin: { left: 10, right: 5, top: 4, bottom: 3 } },
        { row: 0, col: 1, text: '右' }
      ]
    }
  }]
};

test('buildSlideBase64: 保存的列宽写回 tblGrid（比例 1.5:2.5 → 1.5"×1.6"）', async () => {
  const b64 = await buildSlideBase64({ template: STYLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {} });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  const grid = xml.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/);
  assert.ok(grid, '有 tblGrid');
  const ws = [...grid[0].matchAll(/<a:gridCol w="([\d.]+)"\/>/g)].map((m) => Number(m[1]));
  assert.equal(ws.length, 2);
  // 总宽 = bounds.width 4 英寸 = 3657600 EMU；比例 1.5:2.5
  assert.ok(Math.abs(ws[0] + ws[1] - 3657600) < 10, '列宽总和 = 4 英寸');
  assert.ok(Math.abs(ws[0] / ws[1] - 1.5 / 2.5) < 0.02, '列宽比例与保存一致');
});

test('buildSlideBase64: 单元格边框/边距写回 tcPr（marL.. + a:lnT/lnB）', async () => {
  const b64 = await buildSlideBase64({ template: STYLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {} });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('marL="127000"'), '左 10pt 边距 → 127000 EMU');
  assert.ok(xml.includes('marR="63500"') && xml.includes('marT="50800"') && xml.includes('marB="38100"'), '四边边距 EMU');
  assert.ok(xml.includes('<a:lnT w="25400"') && xml.includes('val="FF0000"'), '上边框 2pt 红');
  assert.ok(xml.includes('<a:lnB w="12700"') && xml.includes('val="00FF00"'), '下边框 1pt 绿');
});

test('buildSlideBase64: 表样式回写（firstRow/bandRow/tableStyleId）', async () => {
  const b64 = await buildSlideBase64({ template: STYLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {} });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('<a:tblPr firstRow="1" bandRow="1">'), 'firstRow/bandRow 写回');
  assert.ok(xml.includes('<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>'), '表样式 ID 写回');
});

test('buildSlideBase64: tableData 覆盖单元格文字并扩展行列', async () => {
  const data = [['A', 'B', 'C', 'D'], ['1', '2', '3', '4'], ['x', 'y', 'z', 'w']];
  const b64 = await buildSlideBase64({ template: STYLE_TEMPLATE, imageDataUrl: '', texts: {}, vars: {}, tableData: { tbl1: data } });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  for (const t of ['A', 'B', 'C', 'D', '1', '2', '3', '4', 'x', 'y', 'z', 'w']) {
    assert.ok(xml.includes('>' + t + '<'), '单元格 ' + t + ' 存在');
  }
  assert.ok(!xml.includes('>左<'), '原模板文字被数据覆盖');
  const grid = xml.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/);
  assert.equal((grid[0].match(/<a:gridCol/g) || []).length, 4, '列扩展为 4 列');
  assert.equal((xml.match(/<a:tr\b/g) || []).length, 3, '行扩展为 3 行');
});

test('parseTableBlock: 解析表样式 ID 与单元格边距', () => {
  const xml = '<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{12345678-1234-1234-1234-123456789ABC}</a:tableStyleId></a:tblPr>' +
    '<a:tblGrid><a:gridCol w="914400"/></a:tblGrid>' +
    '<a:tr h="457200"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>X</a:t></a:r></a:p></a:txBody>' +
    '<a:tcPr><a:tcMar><a:left w="127000"/><a:right w="63500"/><a:top w="50800"/><a:bottom w="38100"/></a:tcMar></a:tcPr></a:tc></a:tr></a:tbl>';
  const t = parseTableBlock(xml);
  assert.equal(t.tblPr.tableStyleId, '{12345678-1234-1234-1234-123456789ABC}');
  assert.equal(t.tblPr.firstRow, true);
  assert.equal(t.tblPr.bandRow, true);
  const c = t.cells[0];
  assert.equal(c.margin.left, 10);
  assert.equal(c.margin.right, 5);
  assert.equal(c.margin.top, 4);
  assert.equal(c.margin.bottom, 3);
});

// —— P0-D：表格自动排版渲染（tables = { [shapeId]: FitResult }，阶段 4）——

// 排版渲染用例模板：3 列表格位（模板首格保存 16pt 样式，验证被 fit.fontSize 强制覆盖）
const FIT_TEMPLATE = {
  schemaVersion: 1, name: 't',
  slideSize: { width: 13.33, height: 7.5 },
  shapes: [{
    id: 'tbl1', type: 'table', role: 'table',
    bounds: { left: 1, top: 1, width: 4, height: 1.25 },
    table: {
      rows: 2, cols: 3,
      colWidths: [1.5, 1.5, 1], rowHeights: [0.5, 0.375],
      tblPr: { firstRow: true, bandRow: true, tableStyleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' },
      cells: [
        { row: 0, col: 0, text: '季度指标', textStyle: { bold: true, size: 16, font: 'Arial', eaFont: '微软雅黑', color: '1F4E79' }, fill: '1F4E79' },
        { row: 0, col: 2, text: '单位' },
        { row: 1, col: 0, text: '营收' },
        { row: 1, col: 1, text: '18.5%' },
        { row: 1, col: 2, text: '万元' }
      ]
    }
  }]
};

async function buildTableXml(fit) {
  const b64 = await buildSlideBase64({ template: FIT_TEMPLATE, texts: {}, vars: {}, tables: { tbl1: fit } });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  return zip.file('ppt/slides/slide1.xml').async('string');
}

test('buildSlideBase64: tables 布局渲染 → gridCol 宽/tr 高/sz/tblGrid 列数精确照写', async () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['季度', '营收', '利润'],
    ['1', '2', '3']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.equal(fit.rows, 2);
  assert.equal(fit.cols, 3);
  const xml = await buildTableXml(fit);
  const grid = xml.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/);
  assert.ok(grid, '有 tblGrid');
  const ws = [...grid[0].matchAll(/<a:gridCol w="([\d.]+)"\/>/g)].map((m) => Number(m[1]));
  assert.equal(ws.length, fit.colWidths.length, 'tblGrid 列数 = cols');
  ws.forEach((w, i) => assert.ok(Math.abs(w - fit.colWidths[i] * 914400) < 100,
    'gridCol[' + i + '] 宽 = colWidths×EMU：' + w + ' vs ' + fit.colWidths[i] * 914400));
  const trs = [...xml.matchAll(/<a:tr\b[^>]*h="([\d.]+)"/g)].map((m) => Number(m[1]));
  assert.equal(trs.length, fit.rowHeights.length, 'tr 行数 = rows');
  trs.forEach((h, i) => assert.ok(Math.abs(h - fit.rowHeights[i] * 914400) < 5000,
    'tr[' + i + '] 高 = rowHeights×EMU：' + h + ' vs ' + fit.rowHeights[i] * 914400));
  // 字号全表统一 = fit.fontSize×100（模板首格保存的 16pt 被强制覆盖为 14）
  const szs = [...xml.matchAll(/sz="(\d+)"/g)].map((m) => Number(m[1]));
  assert.ok(szs.length >= fit.cells.length, '单元格 rPr 均带字号，实际 ' + szs.length + ' 处');
  assert.ok(szs.every((s) => Math.abs(s - fit.fontSize * 100) < 0.01),
    '全表字号 = fit.fontSize×100（' + fit.fontSize * 100 + '），实际 ' + szs.join(','));
  assert.ok(!xml.includes('sz="1600"'), '模板首格 16pt 被 fit.fontSize 覆盖');
});

test('buildSlideBase64: tables 合并渲染 → gridSpan/rowSpan/vMerge + 主格文字在、被覆盖格无文字', async () => {
  const cells = [
    { r: 0, c: 0, rowspan: 1, colspan: 2, text: '营业收入' },
    { r: 0, c: 2, rowspan: 1, colspan: 1, text: '单位' },
    { r: 1, c: 0, rowspan: 2, colspan: 1, text: '华东' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: '18%' },
    { r: 1, c: 2, rowspan: 1, colspan: 1, text: '万元' },
    { r: 2, c: 1, rowspan: 1, colspan: 1, text: '22%' },
    { r: 2, c: 2, rowspan: 1, colspan: 1, text: '万元' }
  ];
  const fit = fitTableLayout(cells, { slotWidthIn: 4, slotHeightIn: 1.25 });
  const xml = await buildTableXml(fit);
  assert.ok(xml.includes('gridSpan="2"'), 'colspan=2 → gridSpan="2"');
  assert.ok(xml.includes('rowSpan="2"'), 'rowspan=2 → rowSpan="2"');
  assert.ok(xml.includes('vMerge="1"'), '纵向被覆盖格 → vMerge 占位格');
  // normalize 会把「gridSpan 后的 hMerge 占位」清理成 PowerPoint 原生紧凑格式（行内仅剩主格 + 后续真格）
  assert.ok(!xml.includes('hMerge="1"'), 'gridSpan 后的 hMerge 占位已被 normalize 清理');
  assert.ok(xml.includes('营业收入') && xml.includes('华东'), '主格文字存在');
  // 被覆盖格无文字：vMerge 占位块内不含文字 run（<a:txBody 是容器标签，须匹配 <a:t[ >]）
  const vm = xml.match(/<a:tc vMerge="1">([\s\S]*?)<\/a:tc>/);
  assert.ok(vm && !/<a:t[ >]/.test(vm[1]), 'vMerge 占位无文字');
});

test('applyTableMergesXml: 直接注入 gridSpan/rowSpan + hMerge/vMerge 占位格（normalize 前中间产物）', () => {
  const xml = '<a:tbl>' +
    '<a:tr h="457200"><a:tc><a:txBody><a:p><a:t>M</a:t></a:p></a:txBody></a:tc>' +
    '<a:tc><a:txBody><a:p><a:t>X</a:t></a:p></a:txBody></a:tc>' +
    '<a:tc><a:txBody><a:p><a:t>Y</a:t></a:p></a:txBody></a:tc></a:tr>' +
    '<a:tr h="457200"><a:tc><a:txBody><a:p><a:t>Z</a:t></a:p></a:txBody></a:tc>' +
    '<a:tc><a:txBody><a:p><a:t>W</a:t></a:p></a:txBody></a:tc>' +
    '<a:tc><a:txBody><a:p><a:t>V</a:t></a:p></a:txBody></a:tc></a:tr>' +
    '</a:tbl>';
  const plans = [
    [
      { kind: 'master', colspan: 2, rowspan: 1 },
      { kind: 'covered', hMerge: true, vMerge: false },
      { kind: 'master', colspan: 1, rowspan: 1 }
    ],
    [
      { kind: 'master', colspan: 1, rowspan: 2 },
      { kind: 'master', colspan: 1, rowspan: 1 },
      { kind: 'master', colspan: 1, rowspan: 1 }
    ]
  ];
  const out = applyTableMergesXml(xml, [plans]);
  assert.ok(out.includes('gridSpan="2"'), 'master 格注入 gridSpan');
  assert.ok(out.includes('hMerge="1"'), '横向被覆盖格 → hMerge 占位');
  assert.ok(out.includes('rowSpan="2"'), '纵向 master 格注入 rowSpan');
  assert.ok(!out.includes('<a:t>X</a:t>'), '被覆盖格文字被占位替换');
  assert.ok(out.includes('<a:t>Y</a:t>') && out.includes('<a:t>Z</a:t>'), '非覆盖格原样保留');
});

test('buildSlideBase64: tables overflow（rowH 总和 > 槽高）→ 仍正常生成且 tr 高/ext cy = rowHeights×EMU', async () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['名称', '数值'],
    ['A', '9'.repeat(120)], // 120 位数字：缩到 10pt 仍放不下 → 溢出
    ['B', 'x']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  assert.equal(fit.overflow, true, '超长不可断 → overflow');
  assert.ok(fit.rowHeights.reduce((a, b) => a + b, 0) > 1.25, 'ΣrowH > 槽高（向下延伸）');
  const xml = await buildTableXml(fit);
  const trs = [...xml.matchAll(/<a:tr\b[^>]*h="([\d.]+)"/g)].map((m) => Number(m[1]));
  assert.equal(trs.length, fit.rowHeights.length, '溢出时仍正常生成 tr');
  trs.forEach((h, i) => assert.ok(Math.abs(h - fit.rowHeights[i] * 914400) < 5000, 'overflow 下 tr 高仍 = rowHeights×EMU'));
  // graphicFrame ext cy = ΣrowH（表格向下延伸，超出槽位高度）；
  // 注意：XML 中可能另有 cx="0" cy="0" 的 a:ext，须限定在 p:xfrm 内取真实表格尺寸
  const ext = xml.match(/<p:xfrm>[\s\S]*?<a:ext cx="[\d.]+" cy="([\d.]+)"\s*\/>/);
  assert.ok(ext, '有 ext');
  const cy = Number(ext[1]);
  const sumH = fit.rowHeights.reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(cy - sumH * 914400) < 5000, 'ext cy = ΣrowH×EMU（向下延伸）：' + cy + ' vs ' + sumH * 914400);
});

test('buildSlideBase64: tables 与 tableData 同时传 → tables 优先（tableData 被忽略）', async () => {
  const fit = fitTableLayout(cellsFromGrid([
    ['引擎', '列'],
    ['1', '2']
  ]), { slotWidthIn: 4, slotHeightIn: 1.25 });
  const data = [['数据A', '数据B'], ['x', 'y'], ['z', 'w']];
  const b64 = await buildSlideBase64({ template: FIT_TEMPLATE, texts: {}, vars: {}, tables: { tbl1: fit }, tableData: { tbl1: data } });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('引擎') && xml.includes('列'), '引擎文字生效');
  assert.ok(!xml.includes('数据A'), 'tableData 被 tables 覆盖（不渲染其文字）');
  const grid = xml.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/);
  assert.equal((grid[0].match(/<a:gridCol/g) || []).length, 2, '列数 = fit.cols（2），而非 tableData 的 3');
});

test('buildSlideBase64: 同页混合（fit 表 + 旧路径表）→ 合并计划按序对齐，两表各自正确', async () => {
  // 两表模板：tbl1 走自动排版（含合并），tbl2 走旧路径（模板保存内容）
  const mixedTemplate = {
    schemaVersion: 1, name: 't',
    slideSize: { width: 13.33, height: 7.5 },
    shapes: [
      {
        id: 'tbl1', type: 'table', role: 'table',
        bounds: { left: 0.5, top: 0.5, width: 4, height: 1.25 },
        table: { rows: 2, cols: 2, colWidths: [2, 2], rowHeights: [0.6, 0.6], tblPr: { firstRow: true }, cells: [] }
      },
      {
        id: 'tbl2', type: 'table', role: 'table',
        bounds: { left: 0.5, top: 2.5, width: 3, height: 0.8 },
        table: {
          rows: 1, cols: 2, colWidths: [1, 2], rowHeights: [0.8], tblPr: { firstRow: true },
          cells: [
            { row: 0, col: 0, text: '旧路径' },
            { row: 0, col: 1, text: '保持不变' }
          ]
        }
      }
    ]
  };
  const fit = fitTableLayout([
    { r: 0, c: 0, rowspan: 1, colspan: 2, text: '合并表头' },
    { r: 1, c: 0, rowspan: 1, colspan: 1, text: 'A' },
    { r: 1, c: 1, rowspan: 1, colspan: 1, text: 'B' }
  ], { slotWidthIn: 4, slotHeightIn: 1.25 });
  const b64 = await buildSlideBase64({ template: mixedTemplate, texts: {}, vars: {}, tables: { tbl1: fit } });
  const zip = await JSZip.loadAsync(Buffer.from(b64, 'base64'));
  const xml = await zip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('合并表头'), 'fit 表主格文字');
  assert.ok(xml.includes('gridSpan="2"'), 'fit 表合并注入（tableMergePlans 与 tbl1 对齐）');
  assert.ok(xml.includes('旧路径') && xml.includes('保持不变'), '旧路径表按模板内容渲染');
});
