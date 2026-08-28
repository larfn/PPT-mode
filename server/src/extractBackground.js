const JSZip = require('jszip');

const EMU_PER_INCH = 914400;

function parseRels(xml) {
  const out = {};
  const re = /<Relationship\s+Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g;
  let m;
  while ((m = re.exec(xml))) out[m[1]] = m[2];
  return out;
}

function resolveTarget(baseDir, target) {
  const parts = baseDir.split('/').filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function mimeForExt(ext) {
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'png') return 'image/png';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'bmp') return 'image/bmp';
  if (ext === 'svg') return 'image/svg+xml';
  if (ext === 'tif' || ext === 'tiff') return 'image/tiff';
  return 'image/png';
}

function slideOrder(zip) {
  const pres = zip.file('ppt/presentation.xml');
  if (!pres) return [];
  const rels = zip.file('ppt/_rels/presentation.xml.rels');
  if (!rels) return [];
  return Promise.all([pres.async('string'), rels.async('string')]).then(([presXml, relsXml]) => {
    const relMap = parseRels(relsXml);
    const order = [];
    const re = /<p:sldId\s+[^>]*r:id="([^"]+)"/g;
    let m;
    while ((m = re.exec(presXml))) order.push(m[1]);
    return order.map((rId) => {
      const target = relMap[rId];
      if (!target) return null;
      const t = normalizePath(target).replace(/^\/+/, '');
      return t.startsWith('ppt/') ? t : 'ppt/' + t;
    }).filter(Boolean);
  });
}

function normalizePath(p) {
  return p.replace(/\\/g, '/').replace(/^\/+/, '');
}

function parseSlideShapes(xml) {
  const shapes = [];
  const re = /<p:(sp|cxnSp|pic|grpSp)[^>]*>([\s\S]*?)<\/p:\1>/g;
  let m;
  while ((m = re.exec(xml))) {
    const block = m[2];
    const name = (block.match(/name="([^"]*)"/) || [])[1] || '';
    const off = (block.match(/<a:off\s+x="(-?[\d.]+)"\s+y="(-?[\d.]+)"/) || []);
    const ext = (block.match(/<a:ext\s+cx="(-?[\d.]+)"\s+cy="(-?[\d.]+)"/) || []);
    if (off[1] !== undefined && ext[1] !== undefined) {
      shapes.push({
        name,
        left: Number(off[1]) / EMU_PER_INCH,
        top: Number(off[2]) / EMU_PER_INCH,
        width: Number(ext[1]) / EMU_PER_INCH,
        height: Number(ext[2]) / EMU_PER_INCH
      });
    }
  }
  return shapes;
}

function matchScore(disk, office) {
  if (!office || !office.length || !disk.length) return 0;
  let score = 0;
  for (const d of disk) {
    for (const o of office) {
      if (d.name && o.name && d.name === o.name) {
        const posPenalty = Math.abs(d.left - o.left) + Math.abs(d.top - o.top)
          + Math.abs(d.width - o.width) * 0.5 + Math.abs(d.height - o.height) * 0.5;
        score += 10 - Math.min(posPenalty, 9);
        break;
      }
    }
  }
  return score / Math.max(disk.length, office.length);
}

async function readBgFromXml(zip, xmlPath) {
  const file = zip.file(xmlPath);
  if (!file) return null;
  const xml = await file.async('string');
  const bgMatch = xml.match(/<p:bg>[\s\S]*?<\/p:bg>/);
  if (!bgMatch) return null;
  const blip = bgMatch[0].match(/<a:blip\s+r:embed="([^"]+)"/);
  if (!blip) return null;
  const dir = xmlPath.slice(0, xmlPath.lastIndexOf('/'));
  const baseName = xmlPath.slice(xmlPath.lastIndexOf('/') + 1);
  const relsPath = `${dir}/_rels/${baseName}.rels`;
  const relsFile = zip.file(relsPath);
  if (!relsFile) return null;
  const relsXml = await relsFile.async('string');
  const relMap = parseRels(relsXml);
  const target = relMap[blip[1]];
  if (!target) return null;
  const mediaPath = resolveTarget(dir, target);
  const media = zip.file(mediaPath);
  if (!media) return null;
  const base64 = await media.async('base64');
  const ext = mediaPath.split('.').pop().toLowerCase();
  return { type: 'picture', imageDataUrl: `data:${mimeForExt(ext)};base64,${base64}` };
}

async function readLayoutMasterFallback(zip, xmlPath, seen = new Set()) {
  if (!xmlPath || seen.has(xmlPath)) return null;
  seen.add(xmlPath);
  const dir = xmlPath.slice(0, xmlPath.lastIndexOf('/'));
  const baseName = xmlPath.slice(xmlPath.lastIndexOf('/') + 1);
  const relsFile = zip.file(`${dir}/_rels/${baseName}.rels`);
  if (!relsFile) return null;
  const relsXml = await relsFile.async('string');
  const relMap = parseRels(relsXml);
  const target = Object.values(relMap).find((t) => /slideLayout|slideMaster/i.test(t));
  if (!target) return null;
  const nextPath = normalizePath(resolveTarget(dir, target));
  const bg = await readBgFromXml(zip, nextPath);
  if (bg) return { ...bg, source: /slideLayout/i.test(target) ? 'layout' : 'master' };
  return readLayoutMasterFallback(zip, nextPath, seen);
}

// zip 已加载版本：由调用方加载一次 JSZip，避免同一文档重复解析（大文档 base64 解码 + loadAsync 开销大）
async function extractBackgroundFromZip({ zip, slideIndex, shapes }) {
  const order = await slideOrder(zip);
  if (!order.length) return null;

  const candidates = [];
  if (typeof slideIndex === 'number' && slideIndex >= 1 && order[slideIndex - 1]) {
    candidates.push({ path: order[slideIndex - 1], reason: 'index' });
    if (order[slideIndex]) candidates.push({ path: order[slideIndex], reason: 'index0' });
  }
  if (shapes && shapes.length) {
    let best = null;
    for (const slidePath of order) {
      const file = zip.file(slidePath);
      if (!file) continue;
      const xml = await file.async('string');
      const score = matchScore(parseSlideShapes(xml), shapes);
      if (!best || score > best.score) best = { path: slidePath, score };
    }
    if (best && best.score > 0.3) candidates.push({ path: best.path, reason: 'match' });
  }

  for (const cand of candidates) {
    const bg = await readBgFromXml(zip, cand.path);
    if (bg) return { ...bg, source: 'slide' };
    const fallback = await readLayoutMasterFallback(zip, cand.path);
    if (fallback) return fallback;
  }
  // 兜底：任意页有图片背景
  for (const slidePath of order) {
    const bg = await readBgFromXml(zip, slidePath);
    if (bg) return { ...bg, source: 'slide' };
    const fallback = await readLayoutMasterFallback(zip, slidePath);
    if (fallback) return fallback;
  }
  return null;
}

// 兼容旧调用：接收 base64，内部加载 zip 后复用 FromZip 实现
async function extractBackground({ zipBase64, slideIndex, shapes }) {
  const zip = await JSZip.loadAsync(Buffer.from(zipBase64, 'base64'));
  return extractBackgroundFromZip({ zip, slideIndex, shapes });
}

module.exports = { extractBackground, extractBackgroundFromZip, parseSlideShapes, matchScore, parseRels, slideOrder, readBgFromXml, readLayoutMasterFallback, resolveTarget, normalizePath };
