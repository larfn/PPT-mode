const JSZip = require('jszip');
const { slideOrder, parseRels, resolveTarget, normalizePath, parseSlideShapes } = require('./extractBackground.js');

const EMU_PER_INCH = 914400;

function decodeXml(s) {
  return (s || '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function normalizeAlign(v) {
  if (!v) return undefined;
  const a = String(v).toLowerCase();
  if (a === 'ctr') return 'center';
  if (a === 'l') return 'left';
  if (a === 'r') return 'right';
  if (a === 'just') return 'justify';
  if (a === 'dist' || a === 'thaidist') return 'justify';
  return a;
}

function normalizeAnchor(v) {
  if (!v) return undefined;
  const a = String(v).toLowerCase();
  if (a === 'ctr') return 'middle';
  if (a === 'b') return 'bottom';
  if (a === 't') return 'top';
  return undefined;
}

function attrOf(tag, name) {
  const m = tag && tag.match(new RegExp(name + '="([^"]*)"'));
  return m ? m[1] : undefined;
}

// 从 a:defRPr（或任意片段）提取 latin/ea 字体；返回 { latin, ea }
function defRprFonts(xml) {
  const out = { latin: undefined, ea: undefined };
  const latin = xml && xml.match(/<a:latin\b[^>]*typeface="([^"]*)"/);
  if (latin) out.latin = latin[1];
  const ea = xml && xml.match(/<a:ea\b[^>]*typeface="([^"]*)"/);
  if (ea) out.ea = ea[1];
  return out;
}

// 优先取 ea（中文用户看到的字体），无 ea 回退 latin
function pickFont(f) {
  return f.ea || f.latin;
}

// 解析单个形状块（p:sp / p:cxnSp / p:pic）中的文本样式信息
function parseShapeBlock(block) {
  const cNvPr = block.match(/<p:cNvPr\b[^>]*>/);
  const out = {
    name: (cNvPr && attrOf(cNvPr[0], 'name')) || '',
    left: undefined, top: undefined, width: undefined, height: undefined,
    phType: undefined, phIdx: undefined,
    anchor: undefined,
    lstAlign: undefined,
    paragraphs: [],
    // 图片样式（p:pic）：几何形状 / 柔化边缘 / 裁剪，保存模板时随模板 JSON 一并存储
    imgPrst: undefined,
    imgSoftEdgeEmu: undefined,
    imgSrcRect: undefined
  };
  if (/^<p:pic\b/.test(block)) {
    const prstGeom = block.match(/<a:prstGeom\b[^>]*prst="([^"]+)"/);
    if (prstGeom) out.imgPrst = prstGeom[1];
    const softEdge = block.match(/<a:softEdge\b[^>]*rad="([0-9]+)"/);
    if (softEdge) out.imgSoftEdgeEmu = Number(softEdge[1]);
    const srcRect = block.match(/<a:srcRect\b[^>]*\/?>/);
    if (srcRect) {
      const num = (v) => (v === undefined ? 0 : Number(v));
      out.imgSrcRect = {
        l: num(attrOf(srcRect[0], 'l')), t: num(attrOf(srcRect[0], 't')),
        r: num(attrOf(srcRect[0], 'r')), b: num(attrOf(srcRect[0], 'b'))
      };
    }
    // 完整图片样式：spPr 中除 xfrm（位置由模板 bounds 决定）外的全部原样 XML ——
    // 覆盖「图片格式」里的边框(ln)/阴影/反射/辉光/柔化边缘(effectLst)/棱台(sp3d)/三维旋转(scene3d)/形状(prstGeom)
    const spPr = block.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
    if (spPr) {
      const inner = spPr[1].replace(/<a:xfrm[\s\S]*?<\/a:xfrm>/, '').trim();
      if (inner) out.imgSpPrXml = inner;
    }
    // a:blip 上的效果（透明度/图片更正/艺术效果等）：保留 r:embed 之外的属性与子元素。
    // 注意 blip 可能是自闭合（<a:blip r:embed="rId3"/>，无子元素）或带子元素（<a:blip ...>...</a:blip>）。
    const blip = block.match(/<a:blip\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:blip>)/);
    if (blip) {
      const embed = (blip[1] || '').match(/r:embed="([^"]+)"/);
      if (embed) out.imgEmbed = embed[1]; // 图片关系 rId（用于提取图片本体字节，随模板保存图标/装饰图）
      const attrs = (blip[1] || '').replace(/\s*r:embed="[^"]*"/, '').trim();
      const kids = (blip[2] || '').trim();
      if (attrs) out.imgBlipAttrs = attrs;
      if (kids) out.imgBlipKids = kids;
    }
  }
  const off = block.match(/<a:off\s+x="(-?[\d.]+)"\s+y="(-?[\d.]+)"/);
  const ext = block.match(/<a:ext\s+cx="(-?[\d.]+)"\s+cy="(-?[\d.]+)"/);
  if (off) { out.left = Number(off[1]) / EMU_PER_INCH; out.top = Number(off[2]) / EMU_PER_INCH; }
  if (ext) { out.width = Number(ext[1]) / EMU_PER_INCH; out.height = Number(ext[2]) / EMU_PER_INCH; }
  const ph = block.match(/<p:ph\b[^>]*\/?>/);
  if (ph) { out.phType = attrOf(ph[0], 'type'); out.phIdx = attrOf(ph[0], 'idx'); }
  const txBody = block.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
  if (txBody) {
    const bodyPr = txBody[1].match(/<a:bodyPr\b[^>]*\/?>/);
    if (bodyPr) out.anchor = normalizeAnchor(attrOf(bodyPr[0], 'anchor'));
    // 版式/母版占位符常把默认对齐放在 lstStyle 里（页面无显式 algn 时继承）
    const lstStyle = txBody[1].match(/<a:lstStyle\b[^>]*>([\s\S]*?)<\/a:lstStyle>/);
    if (lstStyle) {
      const lvl1 = lstStyle[1].match(/<a:lvl1pPr\b[^>]*>([\s\S]*?)<\/a:lvl1pPr>|<a:lvl1pPr\b[^>]*\/?>/);
      if (lvl1) {
        out.lstAlign = normalizeAlign(attrOf(lvl1[0], 'algn'));
        // 占位符默认字体常写在 lstStyle lvl1 的 defRPr 里（页面 run 无显式字体时继承）
        const def = defRprFonts(lvl1[0]);
        if (def.latin || def.ea) { out.lstFont = def.latin; out.lstEaFont = def.ea; }
      }
    }
    const pRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
    let pm;
    while ((pm = pRe.exec(txBody[1]))) {
      const pBlock = pm[1];
      const pPr = pBlock.match(/<a:pPr\b[^>]*>([\s\S]*?)<\/a:pPr>|<a:pPr\b[^>]*\/?>/);
      const align = pPr ? normalizeAlign(attrOf(pPr[0], 'algn')) : undefined;
      const def = pPr ? defRprFonts(pPr[0]) : {};
      const pDefFont = def.latin;
      const pDefEaFont = def.ea;
      const runs = [];
      const rRe = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g;
      let rm;
      while ((rm = rRe.exec(pBlock))) {
        const rBlock = rm[1];
        const rPrOpen = rBlock.match(/<a:rPr\b[^>]*>/);
        const rPrBlock = rBlock.match(/<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/);
        const attrs = rPrOpen ? rPrOpen[0] : '';
        const inner = rPrBlock ? rPrBlock[1] : '';
        const run = {
          sz: attrOf(attrs, 'sz') ? Number(attrOf(attrs, 'sz')) / 100 : undefined,
          b: attrOf(attrs, 'b') ? attrOf(attrs, 'b') === '1' : undefined,
          i: attrOf(attrs, 'i') ? attrOf(attrs, 'i') === '1' : undefined,
          u: attrOf(attrs, 'u') || undefined,
          strike: attrOf(attrs, 'strike') || undefined,
          baseline: attrOf(attrs, 'baseline') ? Number(attrOf(attrs, 'baseline')) : undefined
        };
        const latin = inner.match(/<a:latin\b[^>]*typeface="([^"]*)"/) || attrs.match(/typeface="([^"]*)"/);
        const ea = inner.match(/<a:ea\b[^>]*typeface="([^"]*)"/);
        run.font = latin ? latin[1] : undefined;
        run.eaFont = ea ? ea[1] : undefined;
        const srgb = inner.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/) || inner.match(/<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
        run.color = srgb ? srgb[1] : undefined;
        const t = rBlock.match(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/);
        run.text = t ? decodeXml(t[1]) : '';
        runs.push(run);
      }
      const text = runs.map((r) => r.text).join('');
      out.paragraphs.push({ align, runs, text, defFont: pDefFont, defEaFont: pDefEaFont });
    }
  }
  return out;
}

// 从幻灯片 XML 提取所有顶层形状块（含文本样式）
function parseSlideShapesDetailed(xml) {
  const shapes = [];
  const re = /<p:(sp|cxnSp|pic)\b[^>]*>([\s\S]*?)<\/p:\1>/g;
  let m;
  while ((m = re.exec(xml))) shapes.push(parseShapeBlock(m[0]));
  return shapes;
}

// 聚合扁平文本样式：与模板 textStyle 结构对应。
// 字体链：run 显式字体 → 段落 defRPr → lstStyle defRPr → 版式/母版占位符 → 母版 txStyles。
function aggregateTextStyle(paras, lstAlign, block) {
  const style = {};
  const paragraphs = paras || [];
  // 对齐：取第一个显式声明的段落对齐；无显式对齐时回退到 lstStyle 默认对齐
  for (const para of paragraphs) if (para.align) { style.align = para.align; break; }
  if (!style.align && lstAlign) style.align = lstAlign;
  // 字体链：font=latin（英文/拉丁字符），eaFont=ea（中文实际显示字体）；
  // 只有 ea 无 latin 时 font 取 ea 兜底（旧模板只有单一字体字段）。
  let font, eaFont;
  for (const para of paragraphs) {
    let found = false;
    for (const r of (para.runs || [])) {
      if (r.font || r.eaFont) { font = r.font || r.eaFont; eaFont = r.eaFont; found = true; break; }
    }
    if (!found && (para.defFont || para.defEaFont)) { font = para.defFont || para.defEaFont; eaFont = para.defEaFont; found = true; }
    if (found) break;
  }
  if (!font && block && (block.lstFont || block.lstEaFont)) { font = block.lstFont || block.lstEaFont; eaFont = block.lstEaFont; }
  if (font) { style.font = font; if (eaFont) style.eaFont = eaFont; }
  const runs = paragraphs.flatMap((para) => para.runs || []);
  if (!runs.length) return Object.keys(style).length ? style : null;
  // 下划线/删除线/上下标：任一 run 出现即视为 true
  if (runs.some((r) => r.u && String(r.u).toLowerCase() !== 'none')) style.underline = true;
  if (runs.some((r) => r.strike && String(r.strike).toLowerCase() === 'sngstrike')) style.strikethrough = true;
  if (runs.some((r) => r.strike && String(r.strike).toLowerCase() === 'dblstrike')) style.doubleStrikethrough = true;
  if (runs.some((r) => r.baseline !== undefined && r.baseline > 0)) style.superscript = true;
  if (runs.some((r) => r.baseline !== undefined && r.baseline < 0)) style.subscript = true;
  const first = runs[0];
  if (runs.every((r) => r.sz === first.sz) && first.sz) style.size = first.sz;
  const definedB = runs.filter((r) => r.b !== undefined).map((r) => r.b);
  if (definedB.length && definedB.every((v) => v === true)) style.bold = true;
  else if (definedB.length && definedB.every((v) => v === false)) style.bold = false;
  const definedI = runs.filter((r) => r.i !== undefined).map((r) => r.i);
  if (definedI.length && definedI.every((v) => v === true)) style.italic = true;
  else if (definedI.length && definedI.every((v) => v === false)) style.italic = false;
  if (runs.every((r) => r.color === first.color) && first.color) style.color = '#' + first.color;
  return style;
}

function shapeName(block) { return block && block.name ? block.name : ''; }

// 将 Office.js 读取的形状与 XML 形状块配对：名字 + 位置打分
function pairwiseMatch(officeShape, blocks) {
  let best = null;
  for (const b of blocks) {
    if (b.left === undefined || b.top === undefined) continue;
    const posPenalty = Math.abs(b.left - officeShape.bounds.left) + Math.abs(b.top - officeShape.bounds.top)
      + Math.abs(b.width - officeShape.bounds.width) * 0.5 + Math.abs(b.height - officeShape.bounds.height) * 0.5;
    const nameBonus = (officeShape.name && b.name && officeShape.name === b.name) ? 10 : 0;
    const score = nameBonus + Math.max(0, 5 - posPenalty);
    if (!best || score > best.score) best = { block: b, score };
  }
  if (best && (best.score >= 9.5 || (best.score >= 4.0 && best.block.name === officeShape.name))) return best.block;
  if (best && best.score >= 4.0 && Math.abs(best.block.left - officeShape.bounds.left) < 0.04 && Math.abs(best.block.top - officeShape.bounds.top) < 0.04) return best.block;
  return null;
}

// 合并策略：Office.js 的值优先（它反映当前实时状态），
// 仅在其缺失或明显是“默认值”时用 XML 的精确值补齐。
function mergeTextStyle(officeStyle, xmlStyle) {
  if (!xmlStyle) return officeStyle || undefined;
  const o = officeStyle || {};
  const x = xmlStyle;
  const merged = { ...o };
  const boolFields = ['underline', 'strikethrough', 'doubleStrikethrough', 'subscript', 'superscript'];
  for (const f of boolFields) {
    if (o[f] === undefined && x[f] !== undefined) merged[f] = x[f];
  }
  if ((o.align === undefined || o.align === 'left') && x.align && x.align !== 'left') merged.align = x.align;
  if ((o.valign === undefined || o.valign === 'top') && x.valign && x.valign !== 'top') merged.valign = x.valign;
  // 字体：XML 链（含占位符/母版继承）优先于 Office.js —— Office.js font.name 只能拿 latin，
  // 且对占位符常返回主题解析值（如 Calibri），与实际显示的中文字体（ea）不符。
  // 有 ea 时 font 取 ea（用户看到的字体），并单独输出 eaFont 供生成端精确写回 <a:ea>。
  if (x.font) merged.font = x.font;
  else if (o.font) merged.font = o.font;
  if (x.eaFont) merged.eaFont = x.eaFont;
  if (!o.size && x.size) merged.size = x.size;
  if (o.bold === undefined && x.bold !== undefined) merged.bold = x.bold;
  if (o.italic === undefined && x.italic !== undefined) merged.italic = x.italic;
  if (!o.color && x.color) merged.color = x.color;
  return merged;
}

// zip 已加载版本：由调用方加载一次 JSZip（read-all 二进制模式与 extractBackground 共享同一份 zip）
async function readStylesFromZip({ zip, slideIndex, shapes }) {
  const order = await slideOrder(zip);
  if (!order.length) return { styles: shapes.map(() => null), tables: [] };

  const slidePaths = [];
  if (typeof slideIndex === 'number' && slideIndex >= 1 && order[slideIndex - 1]) slidePaths.push(order[slideIndex - 1]);
  if (typeof slideIndex === 'number' && order[slideIndex]) slidePaths.push(order[slideIndex]);
  // 性能：slideIndex 有效时目标页已确定，跳过「遍历全部页 XML 打分」的全文档扫描
  //（大文档 100+ 页时该扫描是读取卡顿主因；仅在 slideIndex 缺失/无效时兜底）
  const indexValid = typeof slideIndex === 'number' && slideIndex >= 1 && !!order[slideIndex - 1];
  if (shapes && shapes.length && !indexValid) {
    let best = null;
    for (const slidePath of order) {
      const file = zip.file(slidePath);
      if (!file) continue;
      const xml = await file.async('string');

      const known = parseSlideShapes(xml);
      const sc = matchScoreByOffice(known, shapes);
      if (!best || sc > best.score) best = { path: slidePath, score: sc };
    }
    if (best && best.score > 0.3) slidePaths.push(best.path);
  }

  // 目标页 XML（优先 index，其次形状匹配）
  let slideXml = null, slidePath = null;
  for (const p of slidePaths) {
    const f = zip.file(p);
    if (!f) continue;
    slideXml = await f.async('string');
    slidePath = p;
    break;
  }
  if (!slideXml) return { styles: shapes.map(() => null), tables: [] };

  const slideBlocks = parseSlideShapesDetailed(slideXml);
  const tables = parseTableBlocks(slideXml); // 表格（GraphicFrame/a:tbl），前端 shapes 读不到

  // 版式 / 母版占位符：读取 rels 链
  async function loadRelated(slideXmlPath) {
    const dir = slideXmlPath.slice(0, slideXmlPath.lastIndexOf('/'));
    const baseName = slideXmlPath.slice(slideXmlPath.lastIndexOf('/') + 1);
    const relsFile = zip.file(`${dir}/_rels/${baseName}.rels`);
    if (!relsFile) return null;
    const relsXml = await relsFile.async('string');
    const relMap = parseRels(relsXml);
    const target = Object.values(relMap).find((t) => /slideLayout/i.test(t));
    if (!target) return null;
    const nextPath = normalizePath(resolveTarget(dir, target));
    const file = zip.file(nextPath);
    if (!file) return null;
    return { path: nextPath, xml: await file.async('string') };
  }
  let layout = null, master = null;
  try {
    layout = await loadRelated(slidePath);
    if (layout) {
      const lDir = layout.path.slice(0, layout.path.lastIndexOf('/'));
      const lBase = layout.path.slice(layout.path.lastIndexOf('/') + 1);
      const mRel = zip.file(`${lDir}/_rels/${lBase}.rels`);
      if (mRel) {
        const relsXml = await mRel.async('string');
        const relMap = parseRels(relsXml);
        const target = Object.values(relMap).find((t) => /slideMaster/i.test(t));
        if (target) {
          const mp = normalizePath(resolveTarget(lDir, target));
          const mf = zip.file(mp);
          if (mf) master = { path: mp, xml: await mf.async('string') };
        }
      }
    }
  } catch { layout = null; master = null; }

  const layoutBlocks = layout ? parseSlideShapesDetailed(layout.xml) : [];
  const masterBlocks = master ? parseSlideShapesDetailed(master.xml) : [];

  const styles = [];
  const imageStyles = [];
  for (const officeShape of shapes) {
    if (!officeShape || !officeShape.bounds) { styles.push(null); imageStyles.push(null); continue; }
    const block = pairwiseMatch(officeShape, slideBlocks);
    if (!block) { styles.push(null); imageStyles.push(null); continue; }
    const xmlStyle = aggregateTextStyle(block.paragraphs, block.lstAlign, block);
    if (block.anchor && xmlStyle) xmlStyle.valign = block.anchor;
    // 占位符继承：页面未声明对齐/字体时回退版式/母版同名占位符（优先 phIdx，其次 phType，最后母版 txStyles）
    if (xmlStyle && block.phType) {
      if (!xmlStyle.align) {
        const inherited = findPhAlign(block, layoutBlocks) || findPhAlign(block, masterBlocks);
        if (!inherited && master) {
          const masterAlign = block.phType === 'title' ? txStyleAlign(master.xml, 'title') : txStyleAlign(master.xml, 'body');
          if (masterAlign) xmlStyle.align = masterAlign;
        } else if (inherited) {
          xmlStyle.align = inherited;
        }
      }
      // 字体继承（中文显示字体常在版式/母版占位符样式里，页面 run 无显式 rPr）
      if (!xmlStyle.font) {
        const inheritedF = findPhFont(block, layoutBlocks) || findPhFont(block, masterBlocks);
        if (!inheritedF && master) {
          const mf = masterTxFonts(master.xml, block.phType === 'title' ? 'title' : 'body');
          if (mf && (mf.latin || mf.ea)) { xmlStyle.font = mf.latin || mf.ea; if (mf.ea) xmlStyle.eaFont = mf.ea; }
        } else if (inheritedF) {
          // 继承的 font 已是 latin（英文），eaFont 单独输出（中文）
          xmlStyle.font = inheritedF.latin || inheritedF.ea;
          if (inheritedF.ea) xmlStyle.eaFont = inheritedF.ea;
        }
      }
    }
    // 占位符垂直对齐继承
    if (xmlStyle && !xmlStyle.valign && block.phType) {
      const inheritedV = findPhAnchor(block, layoutBlocks) || findPhAnchor(block, masterBlocks);
      if (inheritedV) xmlStyle.valign = inheritedV;
    }
    styles.push(mergeTextStyle(officeShape.textStyle, xmlStyle));
    // 图片样式：仅图片位（前端标记 type=picture）需要
    if (officeShape.type !== 'picture') { imageStyles.push(null); continue; }
    const img = {};
    // 完整方案：spPr（除 xfrm）原样保存，生成时整体回写（覆盖边框/阴影/反射/辉光/柔化边缘/棱台/三维旋转/形状）
    if (block.imgSpPrXml) img.spPrXml = block.imgSpPrXml;
    // a:blip 上的效果（透明度/图片更正/艺术效果）
    if (block.imgBlipAttrs) img.blipAttrs = block.imgBlipAttrs;
    if (block.imgBlipKids) img.blipKids = block.imgBlipKids;
    // 兼容旧字段（新模板直接走 spPrXml，旧模板无此字段）
    if (block.imgPrst && block.imgPrst !== 'rect') img.shape = block.imgPrst;
    if (block.imgSoftEdgeEmu) img.softEdgeEmu = block.imgSoftEdgeEmu;
    if (block.imgSrcRect && (block.imgSrcRect.l || block.imgSrcRect.t || block.imgSrcRect.r || block.imgSrcRect.b)) {
      img.srcRect = block.imgSrcRect;
    }
    // 图片本体（随模板保存图标/装饰图）：通过图片关系 rId 从 rels → media 提取 base64 dataURL。
    // 仅当该图片是「固定元素」（前端保存时 role=fixed 才写入 shape.content，见 saveTemplate），
    // 此处只负责把本体字节回读给前端；生成端 slideBuilder 对 fixed+picture+content=data: 原样渲染。
    if (block.imgEmbed && slidePath) {
      const dataUrl = await extractPictureDataUrl(zip, slidePath, block.imgEmbed);
      if (dataUrl) img.imageDataUrl = dataUrl;
    }
    imageStyles.push(Object.keys(img).length ? img : null);
  }
  return { styles, imageStyles, tables };
}

// 占位符对齐继承：优先 phIdx 精确匹配，其次同 phType；从占位符段落样式与 lstStyle 中取值
function findPhAlign(block, blocks) {
  const byIdx = [];
  const byType = [];
  for (const b of blocks) {
    if (!b.phType) continue;
    const st = aggregateTextStyle(b.paragraphs, b.lstAlign);
    const align = st && st.align;
    if (!align) continue;
    if (block.phIdx !== undefined && b.phIdx === block.phIdx) byIdx.push(align);
    if (b.phType === block.phType) byType.push(align);
  }
  return byIdx[0] || byType[0] || undefined;
}

// 占位符垂直对齐继承：优先 phIdx 精确匹配，其次同 phType
function findPhAnchor(block, blocks) {
  const byIdx = [];
  const byType = [];
  for (const b of blocks) {
    if (!b.anchor) continue;
    if (block.phIdx !== undefined && b.phIdx === block.phIdx) byIdx.push(b.anchor);
    if (b.phType === block.phType) byType.push(b.anchor);
  }
  return byIdx[0] || byType[0] || undefined;
}

// 母版 txStyles 默认对齐（body/title 第一级），继承链最后一级
function txStyleAlign(xml, type) {
  if (!xml) return undefined;
  const m = xml.match(new RegExp('<' + 'p:' + type + 'Style>([\\s\\S]*?)<' + '/p:' + type + 'Style>'));
  if (!m) return undefined;
  const lvl = m[1].match(/<a:lvl1pPr\b[^>]*>/);
  if (!lvl) return undefined;
  return normalizeAlign(attrOf(lvl[0], 'algn'));
}

// 母版 txStyles 默认字体（body/title 第一级 defRPr 的 latin/ea），字体继承链最后一级
function masterTxFonts(xml, type) {
  if (!xml) return null;
  const m = xml.match(new RegExp('<' + 'p:' + type + 'Style>([\s\S]*?)<' + '/p:' + type + 'Style>'));
  if (!m) return null;
  const lvl = m[1].match(/<a:lvl1pPr\b[^>]*>([\s\S]*?)<\/a:lvl1pPr>|<a:lvl1pPr\b[^>]*\/?>/);
  if (!lvl) return null;
  const f = defRprFonts(lvl[0]);
  return (f.latin || f.ea) ? f : null;
}

// 占位符字体继承：优先 phIdx 精确匹配，其次同 phType；取占位符聚合字体（含 lstStyle defRPr）
function findPhFont(block, blocks) {
  const byIdx = [];
  const byType = [];
  for (const b of blocks) {
    if (!b.phType) continue;
    const st = aggregateTextStyle(b.paragraphs, b.lstAlign, b);
    const font = st && st.font;
    if (!font) continue;
    if (block.phIdx !== undefined && b.phIdx === block.phIdx) byIdx.push({ latin: font, ea: st.eaFont });
    if (b.phType === block.phType) byType.push({ latin: font, ea: st.eaFont });
  }
  return byIdx[0] || byType[0] || null;
}

// 与 extractBackground.matchScore 相同的打分（避免直接依赖其内部实现）
function matchScoreByOffice(disk, office) {
  if (!office || !office.length || !disk || !disk.length) return 0;
  let total = 0;
  for (const d of disk) {
    if (d.left === undefined || d.top === undefined) continue;
    let bestForD = 0;
    for (const o of office) {
      if (o.bounds.left === undefined || o.bounds.top === undefined) continue;
      const posPenalty = Math.abs(d.left - o.bounds.left) + Math.abs(d.top - o.bounds.top)
        + Math.abs(d.width - o.bounds.width) * 0.5 + Math.abs(d.height - o.bounds.height) * 0.5;
      const nameMatch = d.name && o.name && d.name === o.name;
      const score = nameMatch ? 12 - Math.min(posPenalty, 9) : Math.max(0, 6 - posPenalty);
      if (score > bestForD) bestForD = score;
    }
    total += bestForD;
  }
  return total / Math.max(disk.length, office.length);
}

// 兼容旧调用：接收 base64，内部加载 zip 后复用 FromZip 实现
async function readStyles({ zipBase64, slideIndex, shapes }) {
  const zip = await JSZip.loadAsync(Buffer.from(zipBase64, 'base64'));
  return readStylesFromZip({ zip, slideIndex, shapes });
}

// ============ 表格（GraphicFrame / a:tbl）解析 ============
// 表格在 Office.js slide.shapes 中不可见（GraphicFrame），只能从 XML 回读。
// 保存为 TemplateShape.table：结构（行列/合并格）+ 单元格文字/字体/填充/对齐/边框 + 尺寸。
const EMU = 914400;

// 从 <a:tc> 提取单个单元格
function parseTableCell(tcXml, row, col) {
  const tag = tcXml.match(/<a:tc\b[^>]*>/);
  const cell = { row, col, rowspan: 1, colspan: 1, text: '' };
  const gridSpan = attrOf(tag && tag[0], 'gridSpan');
  const rowSpan = attrOf(tag && tag[0], 'rowSpan');
  if (gridSpan) cell.colspan = Number(gridSpan) || 1;
  if (rowSpan) cell.rowspan = Number(rowSpan) || 1;
  if (attrOf(tag && tag[0], 'hMerge') === '1' || attrOf(tag && tag[0], 'vMerge') === '1') cell.merged = true;
  const tcPr = tcXml.match(/<a:tcPr\b[^>]*>([\s\S]*?)<\/a:tcPr>/);
  if (tcPr) {
    const anchor = attrOf(tcPr[0], 'anchor');
    if (anchor) cell.valign = anchor;
    // 单元格填充：取 tcPr 内最后一个 solidFill（边框 ln 的 solidFill 在前，可能误匹配）
    const fills = tcPr[1].match(/<a:solidFill><a:srgbClr\s+val="([0-9A-Fa-f]{6})"/g);
    if (fills && fills.length) {
      const last = fills[fills.length - 1].match(/val="([0-9A-Fa-f]{6})"/);
      if (last) cell.fill = last[1];
    }
    const borders = tcPr[1].match(/<a:tcBorders>([\s\S]*?)<\/a:tcBorders>/);
    if (borders) {
      const border = {};
      for (const side of ['left', 'right', 'top', 'bottom']) {
        const ln = borders[1].match(new RegExp('<a:' + side + '\\b[^>]*>([\\s\\S]*?)<\\/a:' + side + '>'));
        if (ln) {
          const w = attrOf(ln[0], 'w');
          const c = ln[1].match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
          if (w || c) border[side] = { width: w ? Math.max(1, Math.round(Number(w) / 12700)) : 1, color: c ? c[1] : undefined };
        }
      }
      if (Object.keys(border).length) cell.border = border;
    }
    // 单元格边距（a:tcMar：EMU → 磅；PPT 默认 0.1"/0.05" 边距，显式写回避免变化）
    const mar = tcPr[1].match(/<a:tcMar>([\s\S]*?)<\/a:tcMar>/);
    if (mar) {
      const margin = {};
      for (const side of ['left', 'right', 'top', 'bottom']) {
        const el = mar[1].match(new RegExp('<a:' + side + '\\b[^>]*w="([\\d.]+)"'));
        if (el) margin[side] = Math.round((Number(el[1]) / 12700) * 10) / 10;
      }
      if (Object.keys(margin).length) cell.margin = margin;
    }
  }
  // 文字：段落拼接（\n 分隔），首 run 取字体信息
  const paras = [];
  const pRe = /<a:p>([\s\S]*?)<\/a:p>/g;
  let pm;
  while ((pm = pRe.exec(tcXml))) {
    const pBlock = pm[1];
    let align;
    const pPr = pBlock.match(/<a:pPr\b[^>]*>/);
    if (pPr) {
      const a = attrOf(pPr[0], 'algn');
      if (a === 'ctr') align = 'center';
      else if (a === 'r') align = 'right';
      else if (a === 'just') align = 'justify';
      else if (a === 'l' || a === 'dist') align = a === 'dist' ? 'justify' : undefined;
    }
    const texts = [];
    const runRe = /<a:r>([\s\S]*?)<\/a:r>/g;
    let rm;
    while ((rm = runRe.exec(pBlock))) {
      const t = rm[1].match(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/);
      if (t) texts.push(decodeXml(t[1]));
      if (!cell.textStyle) {
        const rPr = rm[1].match(/<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/);
        if (rPr) {
          const sz = attrOf(rPr[0], 'sz');
          const b = attrOf(rPr[0], 'b');
          const i = attrOf(rPr[0], 'i');
          const latin = rPr[1].match(/<a:latin\b[^>]*typeface="([^"]*)"/);
          const ea = rPr[1].match(/<a:ea\b[^>]*typeface="([^"]*)"/);
          const color = rPr[1].match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
          cell.textStyle = {
            ...(latin ? { font: latin[1] } : {}),
            ...(ea ? { eaFont: ea[1] } : {}),
            ...(sz ? { size: Number(sz) / 100 } : {}),
            ...(b === '1' ? { bold: true } : {}),
            ...(i === '1' ? { italic: true } : {}),
            ...(color ? { color: color[1] } : {}),
            ...(align ? { align } : {})
          };
        }
      }
    }
    if (align && cell.textStyle) cell.textStyle.align = align;
    paras.push(texts.join(''));
  }
  cell.text = paras.join('\n');
  return cell;
}

// 解析整个 <a:tbl> → 结构化表格
function parseTableBlock(tblXml) {
  const tblPr = tblXml.match(/<a:tblPr\b[^>]*>/);
  const styleId = tblXml.match(/<a:tableStyleId>([^<]+)<\/a:tableStyleId>/);
  const colWidths = [];
  const gridRe = /<a:gridCol\b[^>]*w="([\d.]+)"\s*\/>/g;
  let gm;
  while ((gm = gridRe.exec(tblXml))) colWidths.push(Number(gm[1]) / EMU);
  const rowHeights = [];
  const cells = [];
  const trRe = /<a:tr\b([^>]*)>([\s\S]*?)<\/a:tr>/g;
  let tm, rowIdx = 0;
  while ((tm = trRe.exec(tblXml))) {
    const hAttr = tm[1].match(/h="([\d.]+)"/);
    rowHeights.push(hAttr ? Number(hAttr[1]) / EMU : 0);
    const tcRe = /<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g;
    let cm, colIdx = 0;
    while ((cm = tcRe.exec(tm[2]))) {
      cells.push(parseTableCell(cm[0], rowIdx, colIdx)); // 完整 tc 块（含 gridSpan/rowSpan 属性）
      colIdx++;
    }
    rowIdx++;
  }
  // 列数：用每行展开后列数（col + colspan）的最大值；gridCol 数量在 pptxgenjs 下不可靠
  const expandedCols = cells.length ? Math.max(...cells.map((c) => c.col + (c.colspan || 1))) : 0;
  return {
    rows: rowIdx,
    cols: expandedCols || colWidths.length || 0,
    colWidths,
    rowHeights,
    tblPr: tblPr
      ? {
          ...(attrOf(tblPr[0], 'firstRow') === '1' ? { firstRow: true } : {}),
          ...(attrOf(tblPr[0], 'bandRow') === '1' ? { bandRow: true } : {}),
          ...(styleId ? { tableStyleId: styleId[1] } : {})
        }
      : {},
    cells
  };
}

// 从幻灯片 XML 提取所有表格（GraphicFrame 内 a:tbl），含 bounds（EMU→英寸）
function parseTableBlocks(xml) {
  const tables = [];
  const re = /<p:graphicFrame\b[^>]*>([\s\S]*?)<\/p:graphicFrame>/g;
  let m;
  while ((m = re.exec(xml))) {
    const block = m[0];
    const tbl = block.match(/<a:tbl>([\s\S]*?)<\/a:tbl>/);
    if (!tbl) continue;
    const off = block.match(/<a:off x="(-?[\d.]+)" y="(-?[\d.]+)"\s*\/>/);
    const ext = block.match(/<a:ext cx="(-?[\d.]+)" cy="(-?[\d.]+)"\s*\/>/);
    // 形状名（p:cNvPr name）：与 Office.js 读到的形状名一致，前端据此去重「Table 1」重复项
    const cNvPr = block.match(/<p:cNvPr\b[^>]*name="([^"]*)"/);
    const table = parseTableBlock('<a:tbl>' + tbl[1] + '</a:tbl>');
    if (table.rows <= 0 || table.cols <= 0) continue; // 无效/空表格跳过
    tables.push({
      name: cNvPr ? cNvPr[1] : undefined,
      bounds: {
        left: off ? Number(off[1]) / EMU : 0,
        top: off ? Number(off[2]) / EMU : 0,
        width: ext ? Number(ext[1]) / EMU : 0,
        height: ext ? Number(ext[2]) / EMU : 0
      },
      table
    });
  }
  return tables;
}

// 图片本体提取：给定 slide XML 路径与图片关系 rId，从该 slide 的 rels 找到 media 文件，
// 读字节 → base64 dataURL。与 extractBackground.readBgFromXml 同款链路（背景图走的是 p:bg，
// 这里是 p:pic 的 a:blip）。SVG 不支持（pptxgenjs addImage 不吃 SVG），跳过。
function mimeForExt(ext) {
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'png') return 'image/png';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'bmp') return 'image/bmp';
  if (ext === 'tif' || ext === 'tiff') return 'image/tiff';
  return 'image/png';
}

async function extractPictureDataUrl(zip, slideXmlPath, rId) {
  try {
    const dir = slideXmlPath.slice(0, slideXmlPath.lastIndexOf('/'));
    const baseName = slideXmlPath.slice(slideXmlPath.lastIndexOf('/') + 1);
    const relsFile = zip.file(dir + '/_rels/' + baseName + '.rels');
    if (!relsFile) return null;
    const relsXml = await relsFile.async('string');
    const relMap = parseRels(relsXml);
    const target = relMap[rId];
    if (!target) return null;
    const mediaPath = resolveTarget(dir, target);
    const media = zip.file(mediaPath);
    if (!media) return null;
    const ext = mediaPath.split('.').pop().toLowerCase();
    if (ext === 'svg') return null; // SVG 无法用 pptxgenjs 渲染，不随存
    const base64 = await media.async('base64');
    return 'data:' + mimeForExt(ext) + ';base64,' + base64;
  } catch {
    return null;
  }
}

module.exports = { readStyles, readStylesFromZip, parseSlideShapesDetailed, aggregateTextStyle, mergeTextStyle, pairwiseMatch, parseShapeBlock, parseTableBlocks, parseTableBlock, extractPictureDataUrl };
