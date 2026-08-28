import { translateText } from './lib/i18n.js';

// 轻量进度条：给等待中的按钮提供实时反馈动画，不影响实际业务逻辑。
export interface ProgressHandle {
  setText(text: string): void;
  setPct(pct: number | null): void; // null 表示不确定进度（滑动动画），数字表示 0-100
  done(): void;
}

export function showProgress(anchor: HTMLElement, initialText = '处理中…'): ProgressHandle {
  const wrap = document.createElement('div');
  wrap.className = 'progress-wrap';
  wrap.innerHTML = `<div class="progress-track"><div class="progress-bar indeterminate"></div></div><div class="progress-text"></div>`;
  anchor.insertAdjacentElement('afterend', wrap);
  const bar = wrap.querySelector('.progress-bar') as HTMLElement;
  const text = wrap.querySelector('.progress-text') as HTMLElement;
  text.textContent = translateText(initialText);
  return {
    setText(t: string) { text.textContent = translateText(t); },
    setPct(p: number | null) {
      bar.classList.toggle('indeterminate', p === null);
      if (p !== null) bar.style.width = `${Math.max(0, Math.min(100, p))}%`;
    },
    done() { wrap.remove(); }
  };
}


// 简易弹窗（可自定义按钮），返回被点击按钮的 id。
export interface ModalButton { id: string; label: string; kind?: 'primary' | 'secondary' | 'danger'; }
export function showModal(opts: { title: string; message: string; buttons: ModalButton[] }): Promise<string> {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal-box" role="dialog" aria-modal="true">
        <h3 class="modal-title"></h3>
        <div class="modal-msg"></div>
        <div class="modal-actions"></div>
      </div>`;
    (overlay.querySelector('.modal-title') as HTMLElement).textContent = translateText(opts.title);
    (overlay.querySelector('.modal-msg') as HTMLElement).textContent = translateText(opts.message);
    const actions = overlay.querySelector('.modal-actions') as HTMLElement;
    opts.buttons.forEach((b) => {
      const btn = document.createElement('button');
      btn.textContent = translateText(b.label);
      btn.className = b.kind === 'primary' ? 'primary' : (b.kind === 'danger' ? 'danger' : 'secondary');
      btn.addEventListener('click', () => { overlay.remove(); resolve(b.id); });
      actions.appendChild(btn);
    });
    document.body.appendChild(overlay);
  });
}


// 轻量 toast：短暂提示后自动消失（默认 1 秒）
export function showToast(text: string, durationMs = 1000): void {
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = translateText(text);
  document.body.appendChild(t);
  requestAnimationFrame(() => requestAnimationFrame(() => t.classList.add('show')));
  setTimeout(() => {
    t.classList.remove('show');
    setTimeout(() => t.remove(), 300);
  }, durationMs);
}

// 全屏预览弹窗：X 按钮或点击遮罩（弹窗外区域）关闭，遮罩带虚化效果
export function showPreviewModal(contentHtml: string): void {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay preview-overlay';
  overlay.innerHTML = `
    <div class="modal-box preview-box">
      <div class="preview-head"><span class="preview-title">页面预览</span><button class="preview-close" aria-label="关闭预览">✕</button></div>
      <div class="preview-body"></div>
    </div>`;
  (overlay.querySelector('.preview-body') as HTMLElement).innerHTML = contentHtml;
  const close = () => overlay.remove();
  overlay.querySelector('.preview-close')!.addEventListener('click', close);
  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
  document.body.appendChild(overlay);
}
