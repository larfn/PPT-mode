// 带圈问号 + 悬停气泡提示：界面上的小字说明统一收纳，保持界面简洁。
// 用法：infoTip('完整说明文字') 生成 <span class="info-tip" data-tip="...">?</span>；
// 鼠标悬停时在旁边弹出对话框样式气泡（自动贴边翻转，不超出任务窗格）。
import { escapeAttr } from './html.js';

let tipEl: HTMLDivElement | null = null;
let hideTimer: ReturnType<typeof setTimeout> | undefined;

function ensureTipEl(): HTMLDivElement {
  if (!tipEl) {
    tipEl = document.createElement('div');
    tipEl.className = 'info-tip-bubble';
    tipEl.style.display = 'none';
    document.body.appendChild(tipEl);
  }
  return tipEl;
}

function hideTip(): void {
  if (hideTimer) clearTimeout(hideTimer);
  hideTimer = setTimeout(() => { if (tipEl) tipEl.style.display = 'none'; }, 100);
}

// 事件委托：页面内容会频繁 innerHTML 重绘，无需每处重新绑定；
// 悬停 .info-tip 时显示气泡，气泡定位在问号右侧，空间不足自动翻到左侧/上下贴边。
export function bindTooltips(): void {
  document.addEventListener('mouseover', (e) => {
    const t = (e.target as HTMLElement)?.closest?.('.info-tip') as HTMLElement | null;
    if (!t) { hideTip(); return; }
    if (hideTimer) clearTimeout(hideTimer);
    const el = ensureTipEl();
    el.textContent = t.getAttribute('data-tip') || '';
    el.style.display = 'block';
    el.style.maxWidth = '240px';
    const r = t.getBoundingClientRect();
    const vw = window.innerWidth || document.documentElement.clientWidth;
    const vh = window.innerHeight || document.documentElement.clientHeight;
    // 先隐藏定位，量出气泡实际尺寸再摆放
    el.style.left = '-9999px';
    el.style.top = '-9999px';
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    let left = r.right + 8;
    if (left + w > vw - 4) left = Math.max(4, r.left - w - 8); // 右侧放不下 → 左侧
    let top = r.top + r.height / 2 - h / 2;
    if (top < 4) top = 4;
    if (top + h > vh - 4) top = vh - h - 4;
    el.style.left = left + 'px';
    el.style.top = top + 'px';
  });
  document.addEventListener('mouseout', () => hideTip());
}

// 生成带圈问号标记（问号文字由 CSS 显示为圆圈，标题文本保持在原位置）
export function infoTip(text: string): string {
  return '<span class="info-tip" data-tip="' + escapeAttr(text) + '"></span>';
}
