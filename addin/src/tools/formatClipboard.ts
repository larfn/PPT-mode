// 小工具 · 复制格式（格式刷）
// 复制：读取选中文本框逐段的字体+段落格式，存内存（会话级，关闭即失效）。
// 粘贴：把内存格式逐段应用到目标文本框（目标段落更多时沿用最后一段格式）。
// 清除：清空内存格式。

import { getSelectedTarget } from './selection.js';
import { formatIndexForTargetParagraph } from './formatMapping.js';
import { splitParagraphs } from './textOps.js';

// ---------- 格式快照类型 ----------

export interface ParaFontFormat {
  name?: string | null;
  size?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  color?: string | null;
  underline?: string | null;
  allCaps?: boolean | null;
  smallCaps?: boolean | null;
  strikethrough?: boolean | null;
  doubleStrikethrough?: boolean | null;
  subscript?: boolean | null;
  superscript?: boolean | null;
}

export interface ParaFormat {
  font: ParaFontFormat;
  align?: string | null; // horizontalAlignment
}

// ---------- 内存剪贴板（会话级，不落盘） ----------

let clipFormats: ParaFormat[] | null = null;
let clipSource: string | null = null; // 复制来源描述，仅提示用

export function hasClipboard(): boolean {
  return clipFormats !== null && clipFormats.length > 0;
}

export function clearClipboard(): void {
  clipFormats = null;
  clipSource = null;
}

// ---------- 读取选中文本框格式 ----------

// 读取选中文本框每段的格式（getSubstring 逐段定位 → font + paragraphFormat）
export async function copySelectedFormat(): Promise<{ ok: boolean; message: string }> {
  const target = await getSelectedTarget();
  if (!target) return { ok: false, message: '请先选中一个文本框' };
  if (target.kind !== 'text') return { ok: false, message: '当前选中的是表格，请先选中文本框' };

  try {
    const formats = await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load('items');
      await context.sync();
      const items = shapes.items;
      if (!items || items.length === 0) throw new Error('未选中文本框');
      const tr = items[0].textFrame.textRange;
      tr.load('text');
      await context.sync();
      // 必须基于实际文本（tr.text 原样，含 \r）计算偏移
      const actualText = (tr.text || '') as string;
      const segs = splitParagraphs(actualText);

      // 每段：substring → font + paragraphFormat 代理，先统一 load
      const handles: { f: PowerPoint.ShapeFont; pf: PowerPoint.ParagraphFormat }[] = [];
      for (const s of segs) {
        if (s.len === 0) continue; // 空段跳过
        const sub = tr.getSubstring(s.start, s.len);
        const f = sub.font;
        f.load(['name', 'size', 'bold', 'italic', 'color', 'underline', 'allCaps', 'smallCaps', 'strikethrough', 'doubleStrikethrough', 'subscript', 'superscript']);
        const pf = sub.paragraphFormat;
        pf.load('horizontalAlignment');
        handles.push({ f, pf });
      }
      await context.sync();

      // sync 后读取代理属性 → 纯对象快照
      return handles.map((h) => {
        const font: ParaFontFormat = {
          name: h.f.name,
          size: h.f.size,
          bold: h.f.bold,
          italic: h.f.italic,
          color: h.f.color,
          underline: h.f.underline,
          allCaps: h.f.allCaps,
          smallCaps: h.f.smallCaps,
          strikethrough: h.f.strikethrough,
          doubleStrikethrough: h.f.doubleStrikethrough,
          subscript: h.f.subscript,
          superscript: h.f.superscript,
        };
        return { font, align: h.pf.horizontalAlignment as string | null };
      });
    });

    // 保留有内容段的格式（空段格式无意义）
    clipFormats = formats.filter((p) => {
      const hasFont = Object.values(p.font).some((v) => v !== null && v !== undefined);
      return hasFont;
    });
    if (!clipFormats || clipFormats.length === 0) {
      clipFormats = null;
      return { ok: false, message: '未读取到格式（文本框可能为空或无格式）' };
    }
    clipSource = target.text.trim().slice(0, 20) || '文本框';
    return { ok: true, message: '已复制 ' + clipFormats.length + ' 段格式（当前会话有效）' };
  } catch (err) {
    return { ok: false, message: '复制格式失败：' + (err instanceof Error ? err.message : String(err)) };
  }
}

// ---------- 粘贴格式到选中文本框 ----------

// 把内存格式应用到选中文本框：逐段设置 font + horizontalAlignment。
// 段落数不一致：目标段落更多时沿用来源最后一段格式，避免标题/正文模式循环错位。
export async function pasteFormatToSelection(): Promise<{ ok: boolean; message: string }> {
  if (!clipFormats || clipFormats.length === 0) {
    return { ok: false, message: '剪贴板为空，请先复制格式' };
  }
  const target = await getSelectedTarget();
  if (!target) return { ok: false, message: '请先选中一个文本框' };
  if (target.kind !== 'text') return { ok: false, message: '当前选中的是表格，请先选中文本框' };

  try {
    const clip = clipFormats;
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load('items');
      await context.sync();
      const items = shapes.items;
      if (!items || items.length === 0) throw new Error('未选中文本框');
      const tr = items[0].textFrame.textRange;
      tr.load('text');
      await context.sync();
      // 偏移基于实际文本（tr.text 原样）
      const actualText = (tr.text || '') as string;
      const segs = splitParagraphs(actualText);
      if (segs.length === 0) return { ok: false, message: '目标文本框为空' };

      // 只对非空段递增映射索引（空段跳过不占格式位，避免格式错位）
      let fmtIdx = 0;
      let applied = 0;
      for (let i = 0; i < segs.length; i++) {
        const s = segs[i];
        if (s.len === 0) continue;
        const fmt = clip[formatIndexForTargetParagraph(fmtIdx, clip.length)];
        fmtIdx++;
        if (!fmt) continue;
        const sub = tr.getSubstring(s.start, s.len);
        const f = sub.font;
        if (fmt.font.name !== null && fmt.font.name !== undefined) f.name = fmt.font.name;
        if (fmt.font.size !== null && fmt.font.size !== undefined) f.size = fmt.font.size;
        if (fmt.font.bold !== null && fmt.font.bold !== undefined) f.bold = fmt.font.bold;
        if (fmt.font.italic !== null && fmt.font.italic !== undefined) f.italic = fmt.font.italic;
        if (fmt.font.color !== null && fmt.font.color !== undefined) f.color = fmt.font.color;
        if (fmt.font.underline !== null && fmt.font.underline !== undefined) f.underline = fmt.font.underline as never;
        if (fmt.font.allCaps !== null && fmt.font.allCaps !== undefined) f.allCaps = fmt.font.allCaps;
        if (fmt.font.smallCaps !== null && fmt.font.smallCaps !== undefined) f.smallCaps = fmt.font.smallCaps;
        if (fmt.font.strikethrough !== null && fmt.font.strikethrough !== undefined) f.strikethrough = fmt.font.strikethrough;
        if (fmt.font.doubleStrikethrough !== null && fmt.font.doubleStrikethrough !== undefined) f.doubleStrikethrough = fmt.font.doubleStrikethrough;
        if (fmt.font.subscript !== null && fmt.font.subscript !== undefined) f.subscript = fmt.font.subscript;
        if (fmt.font.superscript !== null && fmt.font.superscript !== undefined) f.superscript = fmt.font.superscript;
        if (fmt.align !== null && fmt.align !== undefined) {
          sub.paragraphFormat.horizontalAlignment = fmt.align as never;
        }
        applied++;
      }
      await context.sync();
      return { ok: true, message: '已粘贴 ' + applied + ' 段格式' };
    });
  } catch (err) {
    return { ok: false, message: '粘贴格式失败：' + (err instanceof Error ? err.message : String(err)) };
  }
}
