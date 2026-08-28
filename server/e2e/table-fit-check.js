'use strict';
// ===== 表格自动排版引擎 E2E（"用一次"验证）=====
// 验证链路：粘贴 HTML/CSV → 统一模型 → fit 引擎（列宽/行高/字号/合并/超出）→ 后端渲染 → PPTX XML 断言。
// 运行：node server/e2e/table-fit-check.js   （退出码 0=通过 1=失败）
const path = require('node:path');
const JSZip = require('../node_modules/jszip');
const { buildSlideBase64 } = require('../src/slideBuilder.js');
const {
  fitTableLayout, cellsFromGrid, expandGrid, mergeSameTextCells,
} = require('../../addin/src/lib/tableModel.ts');
const { parseTableHtml, parseTableCsv } = require('../../addin/src/lib/tableClipboard.ts');

const EMU = 914400;
const results = [];
function check(name, cond, detail) {
  results.push({ name, ok: !!cond, detail: detail || '' });
  if (!cond) console.log('  ✗ ' + name + (detail ? ' — ' + detail : ''));
}

const SLOT_TEMPLATE = {
  schemaVersion: 1, name: 'fit-tbl',
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

async function loadXml(base64) {
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  return zip.file('ppt/slides/slide1.xml').async('string');
}

async function main() {
  // ---------- 用例 1：正常表格（含 colspan 表头 + rowspan 标签列），14pt 应放下 ----------
  {
    const html = '<table><thead><tr><th colspan="2" style="color:red">营业收入</th><th>单位</th></tr></thead>' +
      '<tbody><tr><td rowspan="2" style="background:#eee">华东</td><td>18.5%</td><td>万元</td></tr>' +
      '<tr><td>22.3%</td><td>万元</td></tr></tbody></table>';
    const cells = parseTableHtml(html);
    check('parseTableHtml 解析出 3 行 3 列含合并', cells.length === 7 && // 表头 2 + 数据行 3 + 2
      cells.some((c) => c.r === 0 && c.c === 0 && c.colspan === 2 && c.text === '营业收入') &&
      cells.some((c) => c.r === 1 && c.c === 0 && c.rowspan === 2 && c.text === '华东'), JSON.stringify(cells));
    check('parseTableHtml 丢弃样式（text 无标签）', cells.every((c) => !/<[a-z]/.test(c.text)));
    const fit = fitTableLayout(cells, { slotWidthIn: 4, slotHeightIn: 1.25, baseFontSizePt: 14, fontFloorPt: 10 });
    check('fit：未缩字 → fontSize=14', fit.fontSize === 14, 'fontSize=' + fit.fontSize + ' overflow=' + fit.overflow);
    check('fit：colW 总和 = 槽宽', Math.abs(fit.colWidths.reduce((a, b) => a + b, 0) - 4) < 1e-6, JSON.stringify(fit.colWidths));
    check('fit：rowH 全部为正且合理', fit.rowHeights.every((h) => h > 0.1), JSON.stringify(fit.rowHeights));
    const b64 = await buildSlideBase64({ template: SLOT_TEMPLATE, texts: {}, vars: {}, tables: { tbl1: fit } });
    const xml = await loadXml(b64);
    check('渲染：含 a:tbl', xml.includes('<a:tbl>'));
    const grid = (xml.match(/<a:gridCol[^>]*\/>/g) || []);
    check('渲染：gridCol 列数 = cols', grid.length === fit.colWidths.length, 'cols=' + grid.length);
    const gw = grid.map((g) => Number(/w="([\d.]+)"/.exec(g)[1]));
    check('渲染：gridCol 宽 = colWidths×EMU', gw.every((w, i) => Math.abs(w - fit.colWidths[i] * EMU) < 50), JSON.stringify(gw));
    const trs = [...xml.matchAll(/<a:tr\b[^>]*h="([\d.]+)"/g)].map((m) => Number(m[1]));
    check('渲染：tr 高 = rowHeights×EMU', trs.length === fit.rowHeights.length && trs.every((h, i) => Math.abs(h - fit.rowHeights[i] * EMU) < 5000), JSON.stringify(trs));
    check('渲染：统一字号 sz=' + fit.fontSize * 100, (xml.match(/sz="([\d.]+)"/g) || []).every((s) => Math.abs(Number(/sz="([\d.]+)"/.exec(s)[1]) - fit.fontSize * 100) < 0.01));
    check('渲染：colspan=2 → gridSpan="2"', xml.includes('gridSpan="2"'));
    check('渲染：rowspan=2 → rowSpan="2"', xml.includes('rowSpan="2"'));
    check('渲染：vMerge 占位格', xml.includes('vMerge="1"'));
    // gridSpan 后的 hMerge 占位被 normalize 清理成 PowerPoint 原生紧凑格式（行内仅剩主格+真格）
    check('渲染：hMerge 占位已被 normalize 清理', !xml.includes('hMerge="1"'));
    check('渲染：主格文字存在', xml.includes('营业收入') && xml.includes('华东'));
    check('渲染：rowspan 覆盖列后的文字存在（22.3% 不被丢弃）', xml.includes('22.3%') && xml.includes('万元'));
    check('渲染：表样式回写', xml.includes('firstRow="1"') && xml.includes('tableStyleId'));
  }

  // ---------- 用例 2：超长不可断内容 → 缩字到下限 → 超出向下延伸 ----------
  {
    const cells = cellsFromGrid([
      ['名称', '数值'],
      ['A', '1234567890123456789012345678901234567890'],
      ['B', '9876543210987654321098765432109876543210']
    ]);
    const fit = fitTableLayout(cells, { slotWidthIn: 3, slotHeightIn: 0.8, baseFontSizePt: 14, fontFloorPt: 10 });
    check('overflow：超长数字 → fontSize 停在 10', fit.fontSize === 10, 'fontSize=' + fit.fontSize);
    check('overflow：overflow=true', fit.overflow === true, 'overflow=' + fit.overflow);
    check('overflow：colW 总和仍 = 槽宽', Math.abs(fit.colWidths.reduce((a, b) => a + b, 0) - 3) < 1e-6, JSON.stringify(fit.colWidths));
    check('overflow：rowH 总和 > 槽高', fit.rowHeights.reduce((a, b) => a + b, 0) > 0.8, JSON.stringify(fit.rowHeights));
    const b64 = await buildSlideBase64({ template: SLOT_TEMPLATE, texts: {}, vars: {}, tables: { tbl1: fit } });
    const xml = await loadXml(b64);
    const trs = [...xml.matchAll(/<a:tr\b[^>]*h="([\d.]+)"/g)].map((m) => Number(m[1]));
    check('overflow：仍正常生成且 tr 高正确', trs.length === 3 && trs.every((h, i) => Math.abs(h - fit.rowHeights[i] * EMU) < 5000));
  }

  // ---------- 用例 3：一鍵合併 + CSV 解析 ----------
  {
    const csv = cellsFromGrid(parseTableCsv('名称\t数值\t备注\nA\t1\t-\nA\t2\t-\nB\t3\tX'));
    check('parseTableCsv TSV 解析', csv.length === 12); // 4 行 × 3 列
    const merged = mergeSameTextCells(csv, 'auto');
    check('一鍵合併：A 纵向合并（rowspan≥2）', merged.some((c) => c.text === 'A' && c.rowspan === 2), JSON.stringify(merged));
  }

  // ---------- 汇总 ----------
  const failed = results.filter((r) => !r.ok);
  console.log('\n==== 表格自动排版引擎 E2E ====');
  console.log('通过 ' + (results.length - failed.length) + '/' + results.length);
  if (failed.length) {
    failed.forEach((f) => console.log('  FAIL: ' + f.name + (f.detail ? ' — ' + f.detail : '')));
    process.exit(1);
  }
  console.log('全部通过 ✓');
}

main().catch((e) => { console.error('E2E 异常：', e); process.exit(1); });
