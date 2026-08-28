// 小工具 · 格式刷段落映射（纯函数，便于单测）

// 目标段数多于来源段数时，不循环套用；超出的段落沿用来源最后一段格式。
// 例如来源「标题 + 正文 + 正文」贴到「标题 + 5 段正文」：1 2 2 2 2 2。
export function formatIndexForTargetParagraph(targetNonEmptyIndex: number, sourceFormatCount: number): number {
  if (sourceFormatCount <= 0) return -1;
  return Math.min(Math.max(0, targetNonEmptyIndex), sourceFormatCount - 1);
}
