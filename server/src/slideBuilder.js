const pptxgen = require('pptxgenjs');
const { applyMathOmmlXml } = require('./mathOmml.js');
const JSZip = require('jszip');

// 展开合并网格：主格在左上，被覆盖位置指向所属主格、真空位为 null。
// 与 addin/src/lib/tableModel.ts 的 expandGrid 语义一致（内联副本：本文件被打包进 pkg 单文件 exe，
// 不能 require 跨目录的 .ts——pkg 静态分析不会把 addin 的 TS 打进 exe，运行时会 MODULE_NOT_FOUND）。
function expandGrid(cells, rows, cols) {
  const grid = Array.from({ length: Math.max(0, rows) }, () => new Array(Math.max(0, cols)).fill(null));
  for (const cell of cells || []) {
    if (!cell || !Number.isFinite(cell.r) || !Number.isFinite(cell.c) || cell.r < 0 || cell.c < 0) continue;
    const rs = Math.max(1, Math.floor(cell.rowspan || 1));
    const cs = Math.max(1, Math.floor(cell.colspan || 1));
    for (let i = 0; i < rs && cell.r + i < rows; i++) {
      for (let j = 0; j < cs && cell.c + j < cols; j++) {
        if (!grid[cell.r + i][cell.c + j]) grid[cell.r + i][cell.c + j] = cell;
      }
    }
  }
  return grid;
}

// ============ 多页后处理（单页兼容）============
// 输入 styles/overrides 若为「扁平数组」= 单页（兼容旧调用）；
// 若为「数组的数组」= 每页一个数组，按 slide1.xml..slideN.xml 顺序应用。

function slideEntries(zip) {
  return Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));
}

function normalizePerSlide(v) {
  const list = (v || []).filter(Boolean);
  if (!list.length) return [];
  return Array.isArray(list[0]) ? list : [list];
}

// ============ XML 纯函数（string → string）：供合并后处理与单步封装共用 ============
// 每步只做内存中的字符串改写，不做 zip 解压/压缩（大 PPT 性能关键：避免多次全量往返）

// 1) 表格合并规范化（整页 XML）：清理 hMerge 占位格 + 修正 tblGrid 有效列数
function normalizeTableXml(xml) {
  if (!xml.includes('<a:tbl>')) return xml;
  let out = xml;
  const trRe = /<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g;
  out = out.replace(trRe, (m, inner) => {
    const open = m.match(/<a:tr\b[^>]*>/) && m.match(/<a:tr\b[^>]*>/)[0];
    const tcRe = /<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g;
    const tcs = [];
    let mm;
    while ((mm = tcRe.exec(inner))) tcs.push(mm[0]);
    const kept = [];
    let i = 0;
    while (i < tcs.length) {
      const tag = tcs[i].match(/<a:tc\b[^>]*>/);
      const span = Number((tag && tag[0].match(/gridSpan="(\d+)"/) || [])[1] || 1);
      kept.push(tcs[i]);
      i += Math.max(1, span); // 跳过 span-1 个 hMerge 占位格
    }
    return open + kept.join('') + '</a:tr>';
  });
  let maxCols = 0;
  const trRe2 = /<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g;
  let m2;
  while ((m2 = trRe2.exec(out))) {
    const tcRe2 = /<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g;
    let cm, rowCols = 0;
    while ((cm = tcRe2.exec(m2[1]))) {
      const tag = cm[0].match(/<a:tc\b[^>]*>/);
      rowCols += Number((tag && tag[0].match(/gridSpan="(\d+)"/) || [])[1] || 1);
    }
    maxCols = Math.max(maxCols, rowCols);
  }
  if (maxCols > 0) {
    out = out.replace(/<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/, (gm, inner) => {
      const cols = (inner.match(/<a:gridCol[^>]*\/>/g) || []).length;
      if (cols <= maxCols) return gm;
      const parts = inner.match(/<a:gridCol[^>]*\/>/g) || [];
      return '<a:tblGrid>' + parts.slice(0, maxCols).join('') + '</a:tblGrid>';
    });
  }
  return out;
}

// 1.5) 表格合并注入（gridSpan/rowSpan + hMerge/vMerge 占位格）：plans 为该页按表格添加顺序的
//      mergePlan 数组（每表 = 按行：{kind:'master',colspan,rowspan} | {kind:'covered',hMerge,vMerge}）。
//      以 <a:tbl> 块为单位：块内 <a:tr> 顺序 = plans 每行，行内 <a:tc> 顺序 = 该行计划列顺序。
//      master 格：开标签加 gridSpan="n"（n>1）/ rowSpan="m"（m>1）；
//      covered 格：整个 <a:tc>…</a:tc> 替换为 hMerge/vMerge 占位格（无合并标志的真空位 → 原样保留）。
//      注意顺序：本函数必须在 normalizeTableXml 之前运行 —— normalize 依赖 gridSpan 定位并
//      清理「gridSpan 后的冗余 hMerge 占位格」；vMerge 占位格无 gridSpan，normalize 原样保留。
function applyTableMergesXml(xml, plans) {
  if (!plans || !plans.length) return xml;
  if (!xml.includes('<a:tbl>')) return xml;
  let ti = 0;
  const tblRe = /<a:tbl>([\s\S]*?)<\/a:tbl>/g;
  return xml.replace(tblRe, (tblBlock) => {
    const plan = plans[ti++];
    if (!plan || !plan.length) return tblBlock;
    let rowIdx = 0;
    const trRe = /<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g;
    return tblBlock.replace(trRe, (rowBlock) => {
      const rowPlan = plan[rowIdx++];
      if (!rowPlan || !rowPlan.length) return rowBlock;
      let colIdx = 0;
      const tcRe = /<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g;
      return rowBlock.replace(tcRe, (tcBlock) => {
        const entry = rowPlan[colIdx++];
        if (!entry) return tcBlock;
        if (entry.kind === 'master') {
          const open = (tcBlock.match(/<a:tc\b[^>]*>/) || [''])[0];
          if (!open) return tcBlock;
          let newOpen = open;
          if (entry.colspan > 1) newOpen = newOpen.replace(/<a:tc\b/, '<a:tc gridSpan="' + entry.colspan + '"');
          if (entry.rowspan > 1) newOpen = newOpen.replace(/<a:tc\b/, '<a:tc rowSpan="' + entry.rowspan + '"');
          return newOpen === open ? tcBlock : newOpen + tcBlock.slice(open.length);
        }
        if (!entry.hMerge && !entry.vMerge) return tcBlock; // 真空位：原样保留空格
        const attrs = [];
        if (entry.hMerge) attrs.push('hMerge="1"');
        if (entry.vMerge) attrs.push('vMerge="1"');
        return '<a:tc ' + attrs.join(' ') + '><a:txBody><a:bodyPr/><a:p><a:endParaRPr lang="zh-CN"/></a:p></a:txBody></a:tc>';
      });
    });
  });
}

// 2) 表样式回写（firstRow/bandRow/tableStyleId）：props 为该页按表格添加顺序的 tblPr 对象数组
function applyTablePropsXml(xml, props) {
  if (!props || !props.length) return xml;
  if (!xml.includes('<a:tblPr')) return xml;
  let ti = 0;
  return xml.replace(/<a:tblPr\b[^>]*\/>/g, (m) => {
    const p = props[ti++] || {};
    if (!p.tableStyleId && !p.firstRow && !p.bandRow) return m;
    const attrs = [];
    if (p.firstRow) attrs.push('firstRow="1"');
    if (p.bandRow) attrs.push('bandRow="1"');
    const stylePart = p.tableStyleId ? '<a:tableStyleId>' + p.tableStyleId + '</a:tableStyleId>' : '';
    return '<a:tblPr' + (attrs.length ? ' ' + attrs.join(' ') : '') + '>' + stylePart + '</a:tblPr>';
  });
}

// 3) 图片样式回写（几何形状/柔化边缘/裁剪/完整 spPr）：styles 为该页按 p:pic 顺序的图片样式数组
function applyImageStylesXml(xml, styles) {
  if (!styles || !styles.length) return xml;
  const picRe = /<p:pic>[\s\S]*?<\/p:pic>/g;
  let idx = 0;
  return xml.replace(picRe, (picBlock) => {
    const st = styles[idx++];
    if (!st) return picBlock;
    let b = picBlock;
    // 0) 完整方案：用模板保存的 spPr 样式（除 xfrm 位置外全部原样：边框/阴影/反射/辉光/柔化边缘/棱台/三维旋转/形状）
    if (st.spPrXml) {
      const spPrRe = /(<p:spPr>\s*<a:xfrm[\s\S]*?<\/a:xfrm>)[\s\S]*?(<\/p:spPr>)/;
      if (spPrRe.test(b)) {
        b = b.replace(spPrRe, '$1' + st.spPrXml + '$2');
      }
    } else {
      // 旧字段兼容：几何形状（roundRect 带默认圆角 adj≈16.7%）/ 柔化边缘
      if (st.shape && st.shape !== 'rect') {
        const prstRe = /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom>/;
        if (prstRe.test(b)) {
          const geom = st.shape === 'roundRect'
            ? '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>'
            : '<a:prstGeom prst="' + st.shape + '"><a:avLst/></a:prstGeom>';
          b = b.replace(prstRe, geom);
        }
      }
      if (st.softEdgeEmu) {
        const soft = '<a:softEdge rad="' + st.softEdgeEmu + '"/>';
        if (/<a:effectLst>/.test(b)) {
          b = b.replace(/<a:effectLst>/, '<a:effectLst>' + soft);
        } else {
          b = b.replace(/<\/p:spPr>/, '<a:effectLst>' + soft + '<\/a:effectLst><\/p:spPr>');
        }
      }
    }
    // a:blip 上的效果（透明度/图片更正/艺术效果）：保留原 r:embed，回写保存的属性与子元素
    if (st.blipAttrs || st.blipKids) {
      b = b.replace(/<a:blip r:embed="(rId\d+)"[^>]*>[\s\S]*?<\/a:blip>/, (m, rid) =>
        '<a:blip r:embed="' + rid + '"' + (st.blipAttrs ? ' ' + st.blipAttrs : '') + '>' + (st.blipKids || '') + '<\/a:blip>');
    }
    // 裁剪：blipFill 的 stretch 前插入 a:srcRect（千分比，与 XML 一致）
    if (st.srcRect && (st.srcRect.l || st.srcRect.t || st.srcRect.r || st.srcRect.b)) {
      const sr = st.srcRect;
      const srcRectXml = '<a:srcRect l="' + sr.l + '" t="' + sr.t + '" r="' + sr.r + '" b="' + sr.b + '"/>';
      b = b.replace(/<a:stretch><a:fillRect\/><\/a:stretch>/, srcRectXml + '<a:stretch><a:fillRect/></a:stretch>');
    }
    return b;
  });
}

// 4) ea（东亚）字体改写：pptxgenjs 会把 latin/ea/cs 写成同一字体，回读的 eaFont 才是中文实际显示字体
function applyEaFontsXml(xml, overrides) {
  if (!overrides || !overrides.length) return xml;
  let idx = 0;
  return xml.replace(/<a:ea typeface="([^"]*)"([^>]*)charset="-122"\/>/g, (m, face, rest) => {
    const ea = overrides[idx++] || null;
    if (!ea || ea === face) return m;
    return '<a:ea typeface="' + ea + '"' + rest + 'charset="-122"/>';
  });
}

// ============ 合并后处理（性能核心：一次 load + 内存连续改写 + 一次 generate）============
// perSlide: 每页 { tableProps: [], tableMergePlans: [], imageStyles: [], eaOverrides: [] }，顺序对应 slide1.xml..slideN.xml
// 大 PPT 不再经历 4 次「全量解压+重压」，而是只解压 1 次、逐页改写、只压缩 1 次。
async function postprocessSlides(base64, perSlide) {
  const perPage = (perSlide || []).map((p) => p || {});
  if (!perPage.length) return base64;
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const entries = slideEntries(zip);
  if (!entries.length) return base64;
  // 并行解压所有页 XML（JSZip async 为异步解压，IO 密集可并行）
  const xmls = await Promise.all(entries.map((e) => zip.file(e).async('string')));
  let changed = false;
  for (let si = 0; si < entries.length; si++) {
    const page = perPage[Math.min(si, perPage.length - 1)] || { tableProps: [], tableMergePlans: [], imageStyles: [], eaOverrides: [] };
    const xml = xmls[si];
    // 顺序：先注入合并（gridSpan/rowSpan + hMerge/vMerge），normalize 再清理「gridSpan 后的冗余占位」
    let out = applyTableMergesXml(xml, page.tableMergePlans || []);
    out = normalizeTableXml(out);
    out = applyTablePropsXml(out, page.tableProps || []);
    out = applyImageStylesXml(out, page.imageStyles || []);
    out = applyEaFontsXml(out, page.eaOverrides || []);
    out = applyMathOmmlXml(out, page.mathOmml || []);
    if (out !== xml) { zip.file(entries[si], out); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'base64' }) : base64;
}

// ============ 单步封装（保持旧导出签名兼容；内部复用 XML 纯函数）============
async function normalizeTableMerges(base64) {
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const entries = slideEntries(zip);
  if (!entries.length) return base64;
  let changed = false;
  for (const entry of entries) {
    const xml = await zip.file(entry).async('string');
    const after = normalizeTableXml(xml);
    if (after !== xml) { zip.file(entry, after); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'base64' }) : base64;
}

async function applyTableProps(base64, perPageProps) {
  const perPage = normalizePerSlide(perPageProps);
  if (!perPage.length) return base64;
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const entries = slideEntries(zip);
  if (!entries.length) return base64;
  let changed = false;
  for (let si = 0; si < entries.length; si++) {
    const props = perPage[si] || [];
    if (!props.length) continue;
    const xml = await zip.file(entries[si]).async('string');
    const after = applyTablePropsXml(xml, props);
    if (after !== xml) { zip.file(entries[si], after); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'base64' }) : base64;
}

async function applyImageStyles(base64, styles) {
  const perSlide = normalizePerSlide(styles);
  if (!perSlide.length) return base64;
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const entries = slideEntries(zip);
  if (!entries.length) return base64;
  let changed = false;
  for (let si = 0; si < entries.length; si++) {
    const pageStyles = perSlide[si] || [];
    if (!pageStyles.length) continue;
    const xml = await zip.file(entries[si]).async('string');
    const after = applyImageStylesXml(xml, pageStyles);
    if (after !== xml) { zip.file(entries[si], after); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'base64' }) : base64;
}

async function applyEaFonts(base64, overrides) {
  const perSlide = normalizePerSlide(overrides);
  if (!perSlide.length) return base64;
  const zip = await JSZip.loadAsync(Buffer.from(base64, 'base64'));
  const entries = slideEntries(zip);
  if (!entries.length) return base64;
  let changed = false;
  for (let si = 0; si < entries.length; si++) {
    const pageOverrides = perSlide[si] || [];
    if (!pageOverrides.length) continue;
    const xml = await zip.file(entries[si]).async('string');
    const after = applyEaFontsXml(xml, pageOverrides);
    if (after !== xml) { zip.file(entries[si], after); changed = true; }
  }
  return changed ? zip.generateAsync({ type: 'base64' }) : base64;
}

function normalizeColor(color) {
  if (!color || typeof color !== 'string') return undefined;
  const trimmed = color.trim().replace(/^#/, '');
  return trimmed ? trimmed : undefined;
}

function normalizeAlign(align) {
  const a = String(align || 'left').toLowerCase();
  if (a === 'distributed' || a === 'thaiddistributed' || a === 'justified' || a === 'justifylow') return 'justify';
  if (a === 'centered') return 'center';
  return a; // left / center / right / justify
}

// 统一换行符：Office.js 多段落文本以 \r 分隔，pptxgenjs 只识别 \n，否则会写入非法字符。
// 同时清理 XML 不允许的控制字符。
function normalizeBreaks(text) {
  return (text || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');
}

// 箭头类形状：方向跟随包围盒；连接符 0 尺寸补最小厚度。
const ARROW_SHAPES = new Set(['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow']);

function resolveArrowShape(shapeName, w, h) {
  if (!ARROW_SHAPES.has(shapeName)) return shapeName;
  const vertical = h > w + 0.001;
  const horizontal = w > h + 0.001;
  if (vertical) {
    if (shapeName === 'rightArrow' || shapeName === 'leftArrow') return 'downArrow';
    if (shapeName === 'leftRightArrow') return 'upDownArrow';
    return shapeName;
  }
  if (horizontal) {
    if (shapeName === 'upArrow' || shapeName === 'downArrow') return 'rightArrow';
    if (shapeName === 'upDownArrow') return 'leftRightArrow';
  }
  return shapeName;
}

// 创建统一的 pptx 实例（layout 用模板 slideSize）
function createPptx(template) {
  const { width, height } = template.slideSize || {};
  if (!width || !height) throw new Error('template.slideSize.width/height are required');
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'TPL', width, height });
  pptx.layout = 'TPL';
  return pptx;
}

// 在给定 pptx 实例上添加一页并填充（单页与套版共用）
// 返回 { imageStyles, eaOverrides, tableProps, tableMergePlans }（与 applyImageStyles/applyEaFonts/applyTableProps/
//   applyTableMergesXml 的 per-page 顺序对应；tableMergePlans 与 tableProps 同序）
// tableData：{ [shapeId]: string[][] } —— 表格位生成时的单元格数据（逐格编辑/粘贴 CSV/AI 生成），可扩展行列
// tables：{ [shapeId]: FitResult } —— 自动排版引擎输出，存在时完全按 colW/rowH/fontSize/合并格渲染（tables 优先于 tableData）
// images: { [shapeId]: dataURL }（多图片位，每个 AI 图片位各自选图）；兼容旧调用 imageDataUrl 单值（作用于所有 AI 图片位）
function addSlideToPptx(pptx, { template, images, imageDataUrl, texts = {}, vars = {}, tableData = {}, tables = {} }) {
  const imgOf = (shapeId) => (images && images[shapeId]) || imageDataUrl || '';
  const { width, height } = template.slideSize || {};
  const slide = pptx.addSlide();
  const bg = template.background;
  if (bg) {
    if (bg.type === 'picture' && bg.imageDataUrl) {
      slide.background = { data: bg.imageDataUrl };
    } else if (bg.type === 'solid' && bg.color) {
      slide.background = { color: normalizeColor(bg.color) };
    }
  }
  const tableProps = []; // 该页每个表格（按添加顺序）的 tblPr（表样式回写用）
  const tableMergePlans = []; // 该页每个表格（按添加顺序）的合并计划（与 tableProps 同序）
  const formulas = []; // 该页公式位（semanticRole=formula）：{ shapeId, latex, raw }，后处理注入 OMML
  for (const s of template.shapes || []) {
    const { left, top, width: w, height: h } = s.bounds || {};
    const ts = s.textStyle || {};
    if (s.role === 'ai_image') {
      const imgData = imgOf(s.id);
      if (imgData && w && w > 0 && h && h > 0) slide.addImage({ data: imgData, x: left, y: top, w, h });
    } else if (s.role === 'ai_text' || s.role === 'manual_var' || (s.role === 'fixed' && s.type !== 'picture' && s.type !== 'table' && typeof s.content === 'string' && s.content !== '')) {
      const tw = w && w > 0 ? w : 0.1;
      const th = h && h > 0 ? h : 0.1;
      const text = normalizeBreaks(s.role === 'ai_text'
        ? (texts[s.id] || '')
        : s.role === 'manual_var'
          ? (vars[s.id] || '')
          : (s.content || ''));
      // 公式位：正文用占位标记写入，后处理时替换为 OMML 专业型数学（latexToOmml 失败则回退普通文本）
      let writeText = text;
      if (s.role === 'ai_text' && s.semanticRole === 'formula' && text.trim()) {
        writeText = '@@OMATH:' + s.id + '@@';
        formulas.push({ shapeId: s.id, latex: text, raw: texts[s.id] || '' });
      }
      // pptxgenjs 3.x 数组顺序为 [left, right, bottom, top]，且值按「磅」处理（模板存英寸，需乘 72）
      const margin = ts.margin
        ? [
            (ts.margin.left ?? 0) * 72,
            (ts.margin.right ?? 0) * 72,
            (ts.margin.bottom ?? 0) * 72,
            (ts.margin.top ?? 0) * 72
          ]
        : undefined;
      const textFill = s.fill && s.fill.type === 'Solid' ? normalizeColor(s.fill.color) : undefined;
      // PPTX 中删除线枚举只有 sngStrike / dblStrike
      const strike = ts.doubleStrikethrough ? 'dblStrike' : ts.strikethrough ? 'sngStrike' : undefined;
      slide.addText(writeText, {
        x: left, y: top, w: tw, h: th,
        fontFace: ts.font || 'Microsoft YaHei',
        fontSize: ts.size || 18,
        bold: ts.bold,
        italic: ts.italic,
        underline: ts.underline === true,
        strike,
        superscript: ts.superscript === true,
        subscript: ts.subscript === true,
        color: normalizeColor(ts.color) || '333333',
        align: normalizeAlign(ts.align),
        valign: ts.valign || 'top',
        margin,
        fit: ts.autoFit || 'none',
        wrap: ts.wordWrap === undefined ? true : ts.wordWrap,
        rotate: s.rotation || undefined,
        fill: textFill ? { color: textFill } : undefined,
        isTextBox: true
      });
    } else if (s.role === 'fixed' && s.type === 'picture' && s.content && s.content.startsWith('data:')) {
      if (w && w > 0 && h && h > 0) slide.addImage({ data: s.content, x: left, y: top, w, h });
    } else if (s.type === 'table' && s.table && w > 0 && h > 0) {
      // 表格位：按保存的列宽/行高比例自动适配到模板保存的 bounds 尺寸（长/高）；
      // tableData（逐格编辑/粘贴 CSV/AI 生成）可覆盖单元格文字并扩展行列
      const t = s.table;
      const data = tableData[s.id];
      const fit = tables[s.id];
      // 自动排版分支：传入 FitResult（阶段 4 渲染入口）→ 完全按引擎输出画（列宽/行高/字号/合并），tables 优先于 tableData
      const hasFit = !!(fit && fit.rows > 0 && fit.cols > 0
        && Array.isArray(fit.colWidths) && Array.isArray(fit.rowHeights) && Array.isArray(fit.cells));
      const savedCols = Array.isArray(t.colWidths) ? t.colWidths : [];
      const savedRows = Array.isArray(t.rowHeights) ? t.rowHeights : [];
      const sumCol = savedCols.reduce((a, b) => a + (Number(b) || 0), 0);
      const sumRow = savedRows.reduce((a, b) => a + (Number(b) || 0), 0);
      const rows = data ? data.length : t.rows;
      const cols = data
        ? Math.max(1, ...data.map((r) => (Array.isArray(r) ? r.length : 0)))
        : t.cols;
      // 列宽：保存的列宽归一化到表格总宽 w；数据比模板宽时，旧列占 75% 空间、新增列均分剩余
      let colW;
      if (sumCol > 0) {
        const extraCount = Math.max(cols - savedCols.length, 0);
        const oldBudget = extraCount > 0 ? w * 0.75 : w;
        const savedSum = savedCols.slice(0, cols).reduce((a, b) => a + (Number(b) || 0), 0);
        colW = Array.from({ length: cols }, (_, i) => {
          if (i < savedCols.length) {
            const sc = Number(savedCols[i]) || 0;
            return savedSum > 0 ? (sc / savedSum) * oldBudget : oldBudget / Math.min(cols, savedCols.length);
          }
          return extraCount > 0 ? (w - oldBudget) / extraCount : 0;
        });
      }
      // 行高：保存的行高归一化到 h；新增行每行 0.3 英寸
      let rowH;
      if (sumRow > 0) {
        const extraCount = Math.max(rows - savedRows.length, 0);
        const oldBudget = Math.max(h - extraCount * 0.3, 0);
        const savedSum = savedRows.slice(0, rows).reduce((a, b) => a + (Number(b) || 0), 0);
        rowH = Array.from({ length: rows }, (_, i) => {
          if (i < savedRows.length) {
            const sr = Number(savedRows[i]) || 0;
            return savedSum > 0 ? (sr / savedSum) * oldBudget : oldBudget / Math.min(rows, savedRows.length);
          }
          return 0.3;
        });
      } else if (data) {
        rowH = Array.from({ length: rows }, () => h / rows); // 无保存行高时均分
      }
      // 单元格索引 Map：避免逐格 find（大表格 O(n^2) → O(1)）
      const cellMap = new Map((t.cells || []).map((x) => [x.row + ':' + x.col, x]));
      const cellAt = (r, c) => cellMap.get(r + ':' + c);
      const optFor = (r, c) => {
        const cell = cellAt(r, c);
        const options = {};
        if (cell) {
          const cts = cell.textStyle || {};
          if (cts.bold) options.bold = true;
          if (cts.italic) options.italic = true;
          if (cts.size) options.fontSize = cts.size;
          if (cts.font || cts.eaFont) options.fontFace = cts.eaFont || cts.font;
          if (cts.color) options.color = cts.color;
          if (cts.align) options.align = cts.align;
          if (cell.valign) options.valign = cell.valign === 'ctr' ? 'middle' : cell.valign;
          if (cell.fill) options.fill = { color: cell.fill };
          // 单元格边距（a:tcMar，磅）：显式写回，避免变回 PPT 默认边距
          if (cell.margin) {
            options.margin = [
              cell.margin.top ?? 0,
              cell.margin.right ?? 0,
              cell.margin.bottom ?? 0,
              cell.margin.left ?? 0
            ];
          }
          // 单元格边框（保存的 a:tcBorders）→ [上,右,下,左]，缺省侧 none（表样式边框由 tableStyleId 负责）
          if (cell.border) {
            const sides = ['top', 'right', 'bottom', 'left'];
            options.border = sides.map((sd) => {
              const b = cell.border[sd];
              return b && (b.color || b.width)
                ? { type: 'solid', color: b.color || '000000', pt: b.width || 0.75 }
                : { type: 'none' };
            });
          }
        }
        return options;
      };
      const textAt = (r, c) => {
        if (data) {
          const rowArr = data[r];
          return rowArr && typeof rowArr[c] === 'string' ? rowArr[c] : '';
        }
        const cell = cellAt(r, c);
        return cell ? (cell.text || '') : '';
      };
      if (hasFit) {
        // ===== 自动排版渲染：colW/rowH/fontSize 精确照写；合并格 = 主格 + 被覆盖格占位 =====
        const rows = fit.rows;
        const cols = fit.cols;
        const colW = fit.colWidths.slice(0, cols);
        const rowH = fit.rowHeights.slice(0, rows);
        const grid = expandGrid(fit.cells, rows, cols); // 展开合并网格：被覆盖位指向所属主格
        const fitFont = Number(fit.fontSize) > 0 ? Number(fit.fontSize) : 14;
        // 样式：模板保存样式（optFor 按模板 s.table.cells 匹配），字号强制 = fit.fontSize（全表统一）；
        // 超出模板行列的格：行 0 用表头样式（模板第 1 行首格）、其余用正文样式（模板第 2 行首格或整体末格）
        const tCells = t.cells || [];
        const headerSrc = tCells.find((x) => x.row === 0) || tCells[0] || null;
        const bodySrc = tCells.find((x) => x.row === 1) || tCells[tCells.length - 1] || headerSrc;
        const headerBase = headerSrc ? optFor(headerSrc.row, headerSrc.col) : {};
        const bodyBase = bodySrc ? optFor(bodySrc.row, bodySrc.col) : headerBase;
        const optForFit = (rr, cc) => {
          const known = cellAt(rr, cc);
          const base = known ? optFor(rr, cc) : (rr === 0 ? headerBase : bodyBase);
          return Object.assign({}, base, { fontSize: fitFont });
        };
        const tableRows = [];
        const mergePlan = []; // 按行：{kind:'master',colspan,rowspan} | {kind:'covered',hMerge,vMerge}（顺序 = 展开网格列）
        for (let rr = 0; rr < rows; rr++) {
          const cellArr = [];
          const rowPlan = [];
          for (let cc = 0; cc < cols; cc++) {
            const g = grid[rr][cc];
            const master = !!g && g.r === rr && g.c === cc;
            if (master) {
              cellArr.push({ text: normalizeBreaks(g.text), options: optForFit(rr, cc) });
              rowPlan.push({ kind: 'master', colspan: g.colspan || 1, rowspan: g.rowspan || 1 });
            } else {
              // 被覆盖格（hMerge/vMerge）或真空位：空文本占位，合并注入交给 applyTableMergesXml
              cellArr.push({ text: '', options: optForFit(rr, cc) });
              rowPlan.push({ kind: 'covered', hMerge: !!g && g.c < cc, vMerge: !!g && g.r < rr });
            }
          }
          tableRows.push(cellArr); // pptxgenjs 3.x：行 = 单元格对象数组（每行 cols 个，与合并计划列序一致）
          mergePlan.push(rowPlan);
        }
        slide.addTable(tableRows, {
          x: left, y: top,
          w: colW.reduce((a, b) => a + b, 0), // 列宽总和 = slotWidthIn
          h: rowH.reduce((a, b) => a + b, 0), // 行高总和：可 > bounds.height → 向下延伸
          rowH, colW,
          autoPage: false,
          border: { type: 'none' }
        });
        tableProps.push(t.tblPr || {});
        tableMergePlans.push(mergePlan);
        continue; // 自动排版分支完成，跳过旧逻辑（旧逻辑照常处理无 fit 的表格）
      }
      const tableRows = [];
      for (let rr = 0; rr < rows; rr++) {
        const cellArr = [];
        for (let cc = 0; cc < cols; cc++) {
          cellArr.push({ text: textAt(rr, cc), options: optFor(rr, cc) });
        }
        tableRows.push(cellArr); // pptxgenjs 3.x：行 = 单元格对象数组
      }
      slide.addTable(tableRows, {
        x: left, y: top, w, h, rowH, colW,
        autoPage: false,
        border: { type: 'none' } // 边框逐格写回（保存的 a:tcBorders），避免默认灰框
      });
      tableProps.push(t.tblPr || {});
      tableMergePlans.push(null); // 非自动排版表格：无合并计划（保持与 <a:tbl> 文档顺序对齐，applyTableMergesXml 按序跳过）
    } else if (s.type === 'rectangle' || s.type === 'line' || (s.type === 'other' && s.role === 'fixed' && !(typeof s.content === 'string' && s.content))) {
      // 几何形状：rectangle/line 以及「无文本的 fixed 其它形状」（手绘/标注/装饰等）一律原样画出来；
      // 直线+圆圈点这类母版图案里的每个元素都要保留，不能因为类型不是 text/rect/line 就被丢弃。
      const known = pptx.ShapeType && pptx.ShapeType[s.shapeType];
      const shapeName0 = known ? s.shapeType : (s.type === 'line' ? 'line' : 'rect');
      const shapeName = resolveArrowShape(shapeName0, w || 0, h || 0);
      // 直线/箭头保存时宽或高常为 0，生成时补最小厚度，避免 0×0 形状
      const MIN_DIM = 0.02;
      let aw = w, ah = h;
      if (!aw || aw <= 0) aw = ARROW_SHAPES.has(shapeName) ? 0.15 : MIN_DIM;
      if (!ah || ah <= 0) ah = ARROW_SHAPES.has(shapeName) ? 0.15 : MIN_DIM;
      const fillColor = s.fill && s.fill.type === 'Solid' ? normalizeColor(s.fill.color) : undefined;
      const lineColor = s.line && s.line.visible ? normalizeColor(s.line.color) : undefined;
      const line = lineColor ? { color: lineColor, width: s.line.weight ?? 1 } : undefined;
      slide.addShape(pptx.ShapeType[shapeName], {
        x: left, y: top, w: aw, h: ah,
        rotate: s.rotation || undefined,
        fill: fillColor ? { color: fillColor } : undefined,
        line
      });
    }
  }
  // 该页「实际渲染的图片位」样式（与 applyImageStyles 的 per-page 顺序一致）
  const imageStyles = (template.shapes || [])
    .filter((s) =>
      (s.role === 'ai_image' && imgOf(s.id)) ||
      (s.role === 'fixed' && s.type === 'picture' && s.content && s.content.startsWith('data:')))
    .map((s) => s.imageStyle || null);
  // 该页 ea 字体覆盖（与 addText 调用顺序一致）
  const eaOverrides = (template.shapes || [])
    .filter((s) =>
      s.role === 'ai_text' || s.role === 'manual_var' ||
      (s.role === 'fixed' && s.type !== 'picture' && s.type !== 'table' && typeof s.content === 'string' && s.content !== ''))
    .map((s) => (s.textStyle && s.textStyle.eaFont && s.textStyle.eaFont !== s.textStyle.font ? s.textStyle.eaFont : null));
  return { imageStyles, eaOverrides, tableProps, tableMergePlans, formulas };
}

// 单页生成（保持原签名与行为；tables 为可选新增参数）
async function buildSlideBase64({ template, images, imageDataUrl, texts = {}, vars = {}, tableData = {}, tables = {} }) {
  const pptx = createPptx(template);
  const r = addSlideToPptx(pptx, { template, images, imageDataUrl, texts, vars, tableData, tables });
  let base64 = await pptx.write('base64');
  // 性能：4 步后处理合并为 1 次 zip 解压 + 1 次压缩（大 PPT 提速关键）
  base64 = await postprocessSlides(base64, [{ tableProps: r.tableProps, tableMergePlans: r.tableMergePlans || [], imageStyles: r.imageStyles, eaOverrides: r.eaOverrides, mathOmml: r.formulas || [] }]);
  return base64;
}

// 整份/多页生成（套版 Deck）：单个 pptxgenjs 实例逐页 addSlide，一次输出多页 base64。
// pages: [{ template, imageDataUrl, texts, vars }]
// 跨页 slideSize 必须一致（否则报错，由调用方提示用户）。
// pages: [{ template, images?, imageDataUrl?, texts, vars, tables? }]（images 多图优先，imageDataUrl 兼容；tables 逐页透传）
async function buildDeckBase64(pages) {
  if (!Array.isArray(pages) || !pages.length) throw new Error('deck has no pages');
  const first = pages[0].template;
  const { width: w0, height: h0 } = first.slideSize || {};
  for (const p of pages) {
    const s = (p && p.template && p.template.slideSize) || {};
    if (Math.abs((s.width || 0) - (w0 || 0)) > 0.001 || Math.abs((s.height || 0) - (h0 || 0)) > 0.001) {
      throw new Error('套版内页面尺寸不一致（要求所有页 slideSize 相同）');
    }
  }
  const pptx = createPptx(first);
  const perSlideStyles = [];
  const perSlideEa = [];
  const perSlideTableProps = [];
  const perSlideTableMerges = [];
  const perSlideFormulas = [];
  for (const p of pages) {
    const r = addSlideToPptx(pptx, p);
    perSlideStyles.push(r.imageStyles);
    perSlideEa.push(r.eaOverrides);
    perSlideTableProps.push(r.tableProps || []);
    perSlideTableMerges.push(r.tableMergePlans || []);
    perSlideFormulas.push(r.formulas || []);
  }
  let base64 = await pptx.write('base64');
  const perSlide = [];
  for (let i = 0; i < perSlideStyles.length; i++) {
    perSlide.push({ tableProps: perSlideTableProps[i] || [], tableMergePlans: perSlideTableMerges[i] || [], imageStyles: perSlideStyles[i] || [], eaOverrides: perSlideEa[i] || [], mathOmml: perSlideFormulas[i] || [] });
  }
  base64 = await postprocessSlides(base64, perSlide);
  return base64;
}

module.exports = { buildSlideBase64, buildDeckBase64, postprocessSlides, normalizeTableXml, applyTableMergesXml, applyTablePropsXml, applyImageStylesXml, applyEaFontsXml, applyImageStyles, applyEaFonts, addSlideToPptx, createPptx, normalizeTableMerges, applyTableProps };
