// 表单必填项错误提示：
//  - 输入框：红框 + 紧跟「！需要输入内容！」红字，用户一输入/修改就自动清除
//  - 文本元素（如提示行）：红色加粗提示文案（markErrorText / clearErrorText）
import type { } from '../api.js';

export function markInputError(input: HTMLElement | null | undefined, message = '！需要输入内容！'): void {
  if (!input) return;
  input.classList.add('input-error');
  input.removeAttribute('aria-invalid');
  input.setAttribute('aria-invalid', 'true');
  let tip = input.nextElementSibling as HTMLElement | null;
  if (!tip || !tip.classList.contains('input-error-tip')) {
    tip = document.createElement('span');
    tip.className = 'input-error-tip';
    input.insertAdjacentElement('afterend', tip);
  }
  tip.textContent = message;
  const clear = () => clearInputError(input);
  input.addEventListener('input', clear, { once: true });
  input.addEventListener('change', clear, { once: true });
  input.focus();
}

export function clearInputError(input: HTMLElement | null | undefined): void {
  if (!input) return;
  input.classList.remove('input-error');
  input.removeAttribute('aria-invalid');
  const tip = input.nextElementSibling;
  if (tip && tip.classList.contains('input-error-tip')) tip.remove();
}

export function clearAllInputErrors(container: HTMLElement): void {
  container.querySelectorAll('.input-error').forEach((el) => clearInputError(el as HTMLElement));
}

// 非输入框元素（提示行/说明文字）的红色错误文案
export function markErrorText(el: Element | null | undefined, message: string): void {
  if (!el) return;
  el.classList.add('text-error');
  el.textContent = message;
}

export function clearErrorText(el: Element | null | undefined): void {
  if (!el) return;
  el.classList.remove('text-error');
}
