// 通用 HTML 转义工具（任务窗格各页面共用，避免重复定义）

// 转义 HTML 文本内容中的特殊字符
export function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// 转义 HTML 属性值（语义区分：属性/文本，实现相同）
export function escapeAttr(s: string): string {
  return escapeHtml(s);
}
