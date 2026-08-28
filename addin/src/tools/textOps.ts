// 小工具 · 纯文本处理函数（无 Office 依赖，可直接单测）
// 段落分隔：读取时统一归一化为 \n 处理；写回时由调用方还原为 PPT 实际分隔符。

// ---------- 空格处理 ----------

// 去除文本中所有空格（半角/全角/不间断空格），保留换行与段落结构。
export function removeAllSpaces(text: string): string {
  return text.replace(/[ \u00A0\u3000]/g, '');
}

// 将连续空格合并为单个半角空格，并去掉行首行尾空白（保留段落结构）。
export function collapseSpaces(text: string): string {
  return text
    .split(/\r?\n/)
    .map((line) => line.replace(/[ \u00A0\u3000]+/g, ' ').trim())
    .join('\n');
}

// ---------- 段首缩进 ----------

// 中文首行缩进：每个非空段段首加两个全角空格。
// 已带两个全角空格开头的段落跳过（避免重复缩进），保留原有段落分隔。
export function indentParagraphs(text: string, indent = '\u3000\u3000'): string {
  const lines = text.split(/\r?\n/);
  const indented = lines.map((line) => {
    if (!line.trim()) return line; // 空段落不动
    if (line.startsWith(indent)) return line; // 已缩进不动
    return indent + line;
  });
  return indented.join('\n');
}

// ---------- 空行空段 ----------

// 删除空段落（trim 后为空的段落），同时把连续换行压缩成单个分隔符，去掉首尾空段。
export function removeEmptyParagraphs(text: string): string {
  return text
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .join('\n');
}

// ---------- 段落分隔（removeEmptyParagraphs 的相反操作） ----------

// 在每个非空段之间插入一个空段，使段落间空一行（已有空段分隔的不重复插入）。
// 规则：连续非空段之间保证恰好 1 个空行；首尾不额外补空行。
export function separateParagraphs(text: string): string {
  const lines = text.split(/\r?\n/);
  const out: string[] = [];
  for (const line of lines) {
    const isBlank = line.trim().length === 0;
    if (!isBlank) {
      // 非空段：如果前一个输出是空段（即上一段与当前段之间已有空行），则不需要再插空段；
      // 如果前一个输出是非空段，则插入一个空段作为分隔。
      if (out.length > 0 && out[out.length - 1].trim().length !== 0) {
        out.push('');
      }
      out.push(line);
    }
  }
  return out.join('\n');
}

// ---------- 提取标题 / 正文 ----------

// 从多段文本中提取「标题」：取第一段非空段；若首段过长（>40 字）则截取到首个句末标点。
// 返回仅标题的文本（单段）。
export function extractTitle(text: string): string {
  const lines = text.split(/\r?\n/);
  const first = lines.find((l) => l.trim().length > 0);
  if (!first) return '';
  let t = first.trim();
  // 首段过长：截到第一个句末标点（。！？；）为止
  if (t.length > 40) {
    const m = t.match(/^[^。！？；]*[。！？；]/);
    if (m) t = m[0];
  }
  return t;
}

// 从多段文本中提取「正文」：去掉标题部分，保留其余段落。
// 标题定义与 extractTitle 一致；若去掉标题后无剩余内容，返回空串（调用方按「无需处理」提示）。
export function extractBody(text: string): string {
  const title = extractTitle(text);
  if (!title) return text;
  // 从头开始匹配标题（精确匹配首段或首段截取部分）
  const lines = text.split(/\r?\n/);
  const idx = lines.findIndex((l) => l.trim() === title);
  if (idx >= 0) {
    const rest = lines.slice(idx + 1);
    // 去掉空段，重组为连续段落
    return rest.filter((l) => l.trim().length > 0).join('\n');
  }
  // 标题在首段中间（截取场景）：去掉开头部分
  const first = lines[0] || '';
  if (first.startsWith(title)) {
    const rest = [first.slice(title.length), ...lines.slice(1)];
    return rest.filter((l) => l.trim().length > 0).join('\n');
  }
  return text;
}


// ---------- 标题/正文定位（供「选中」语义使用） ----------

// 返回标题在文本中的 [start, length) 范围（按段落分隔符计算偏移，兼容 \r 与 \n）。
// 规则与 extractTitle 一致：第一段非空段；首段 >40 字截到首个句末标点。
export function locateTitleRange(text: string): { start: number; length: number } | null {
  const segs = splitParagraphs(text);
  const idx = segs.findIndex((s) => text.slice(s.start, s.start + s.len).trim().length > 0);
  if (idx < 0) return null;
  const seg = segs[idx];
  const raw = text.slice(seg.start, seg.start + seg.len);
  const trimmed = raw.trim();
  let title = trimmed;
  if (trimmed.length > 40) {
    const m = trimmed.match(/^[^。！？；]*[。！？；]/);
    if (m) title = m[0];
  }
  // 偏移 = 段起始 + 前导空白
  const leading = raw.length - trimmed.length;
  return { start: seg.start + leading, length: title.length };
}

// 返回正文范围：标题之后到文本末尾（含标题后的分隔符，便于连续选中）。
export function locateBodyRange(text: string): { start: number; length: number } | null {
  const title = locateTitleRange(text);
  if (!title) return null;
  const start = title.start + title.length;
  // 跳过标题后的分隔符（\r\n / \r / \n，可能多个）
  let s = start;
  while (s < text.length && (text[s] === '\r' || text[s] === '\n')) s++;
  return { start: s, length: text.length - s };
}

// 工具函数：判断文本是否被处理过（供 UI 展示“无需处理”提示）
export function textChanged(before: string, after: string): boolean {
  return before !== after;
}

// ---------- 段落级辅助 ----------

// 段首加一个半角空格（已有半角空格开头的段落跳过），供 selection 层使用
export function indentLines(text: string): string {
  return text
    .split('\n')
    .map((line) => {
      const t = line.trim();
      if (!t) return line;
      if (line.startsWith(' ')) return line;
      return ' ' + line;
    })
    .join('\n');
}

// 删除空行空段：去掉纯空白段落，压缩连续换行
export function removeEmptyLines(text: string): string {
  return text
    .split('\n')
    .filter((line) => line.trim().length > 0)
    .join('\n');
}

// ---------- 段落分隔符归一化（PowerPoint 文本可能是 \r 或 \n） ----------

// 统一归一化为 \n 处理，写回时由 restoreSep 还原
export function normalizeSep(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}
export function detectSep(text: string): string {
  return text.includes('\r') ? '\r' : '\n';
}
export function restoreSep(text: string, sep: string): string {
  return sep === '\r' ? text.replace(/\n/g, '\r') : text;
}

// 按段落分隔符切分，返回每段起始/长度（不含分隔符）。
// 兼容 \r\n / \r / \n 三种分隔符（PowerPoint 实际用 \r，但统一处理避免偏移错位）。
export function splitSegs(text: string, sep?: string): { start: number; len: number }[] {
  return splitParagraphs(text);
}

// 通用段落切分：兼容 \r\n / \r / \n，偏移基于原文本（用于 getSubstring 定位）。
export function splitParagraphs(text: string): { start: number; len: number }[] {
  const out: { start: number; len: number }[] = [];
  let start = 0;
  let i = 0;
  while (i <= text.length) {
    if (i === text.length) {
      out.push({ start, len: i - start });
      break;
    }
    const c = text[i];
    if (c === '\r') {
      out.push({ start, len: i - start });
      start = i + (text[i + 1] === '\n' ? 2 : 1);
      i = start;
    } else if (c === '\n') {
      out.push({ start, len: i - start });
      start = i + 1;
      i = start;
    } else {
      i++;
    }
  }
  return out;
}
