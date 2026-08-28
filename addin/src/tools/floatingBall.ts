// 小工具 · 任务窗格内可拖拽悬浮球
// 点击展开功能菜单；菜单项执行 Office.js 动作（读取选中文本框/表格 → 处理 → 写回）。
// 拖拽：pointer events（down/move/up），限制在视口内；点击（无位移）展开/收起菜单。
// 表格按钮默认隐藏：菜单打开时检测当前选中类型，选中表格才显示表格组。

import { showToast } from '../ui.js';
import { copySelectedFormat, pasteFormatToSelection, clearClipboard } from './formatClipboard.js';
import {
  actionRemoveSpaces,
  actionIndentParagraphs,
  actionRemoveEmptyParagraphs,
  actionSeparateParagraphs,
  actionSelectTitle,
  actionSelectBody,
  actionFitTable,
  actionEvenTable,
  getSelectedTarget,
  type ToolActionResult,
} from './selection.js';

interface MenuItem {
  id: string;
  label: string;
  run: () => Promise<ToolActionResult>;
}

export function mountFloatingBall(): void {
  if (document.getElementById('tool-ball')) return; // 防重复挂载

  const ball = document.createElement('div');
  ball.id = 'tool-ball';
  ball.className = 'tool-ball';
  ball.innerHTML = '<span class="tool-ball-icon">🔧</span>';

  const menu = document.createElement('div');
  menu.id = 'tool-ball-menu';
  menu.className = 'tool-ball-menu';
  menu.hidden = true;

  // ---- 文本组（默认可见）----
  const textGroup = document.createElement('div');
  textGroup.className = 'tool-menu-group';
  textGroup.innerHTML = '<div class="tool-menu-group-title">文本</div>';

  const mkBtn = (it: MenuItem, cls: string): HTMLButtonElement => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'tool-menu-item ' + cls;
    btn.innerHTML = '<span class="tool-menu-label"></span>';
    (btn.querySelector('.tool-menu-label') as HTMLElement).textContent = it.label;
    btn.addEventListener('click', async () => {
      setMenuOpen(false);
      try {
        const res = await it.run();
        showToast(res.message, 2200);
      } catch (err) {
        showToast('操作失败：' + (err instanceof Error ? err.message : String(err)), 3200);
      }
    });
    return btn;
  };

  const row1 = document.createElement('div');
  row1.className = 'tool-menu-row';
  row1.appendChild(mkBtn({ id: 'rm-space', label: '去除空格', run: actionRemoveSpaces }, 'half'));
  row1.appendChild(mkBtn({ id: 'indent', label: '段首空格', run: actionIndentParagraphs }, 'half'));
  textGroup.appendChild(row1);

  const row2 = document.createElement('div');
  row2.className = 'tool-menu-row';
  row2.appendChild(mkBtn({ id: 'rm-empty', label: '删除空段', run: actionRemoveEmptyParagraphs }, 'half'));
  row2.appendChild(mkBtn({ id: 'separate', label: '分隔每段', run: actionSeparateParagraphs }, 'half'));
  textGroup.appendChild(row2);

  const row3 = document.createElement('div');
  row3.className = 'tool-menu-row';
  row3.appendChild(mkBtn({ id: 'select-title', label: '选中标题', run: actionSelectTitle }, 'half'));
  row3.appendChild(mkBtn({ id: 'select-body', label: '选中正文', run: actionSelectBody }, 'half'));
  textGroup.appendChild(row3);

  // ---- 格式组（复制/粘贴/清除，会话级）----
  const fmtGroup = document.createElement('div');
  fmtGroup.className = 'tool-menu-group';
  fmtGroup.innerHTML = '<div class="tool-menu-group-title">格式</div>';
  const fmtRow = document.createElement('div');
  fmtRow.className = 'tool-menu-row';
  fmtRow.appendChild(mkBtn({ id: 'fmt-copy', label: '复制', run: copySelectedFormat }, 'third'));
  fmtRow.appendChild(mkBtn({ id: 'fmt-paste', label: '粘贴', run: pasteFormatToSelection }, 'third'));
  fmtRow.appendChild(mkBtn({ id: 'fmt-clear', label: '清除', run: async () => { clearClipboard(); return { ok: true, message: '已清除复制格式' }; } }, 'third'));
  fmtGroup.appendChild(fmtRow);
  menu.appendChild(textGroup);
  menu.appendChild(fmtGroup);
  const tableGroup = document.createElement('div');
  tableGroup.className = 'tool-menu-group';
  tableGroup.id = 'tool-menu-table-group';
  tableGroup.hidden = true;
  tableGroup.innerHTML = '<div class="tool-menu-group-title">表格</div>';
  const trow = document.createElement('div');
  trow.className = 'tool-menu-row';
  trow.appendChild(mkBtn({ id: 'fit-table', label: '表格适配', run: actionFitTable }, 'half'));
  trow.appendChild(mkBtn({ id: 'even-table', label: '行列均分', run: actionEvenTable }, 'half'));
  tableGroup.appendChild(trow);

  menu.appendChild(tableGroup);
  document.body.appendChild(ball);
  document.body.appendChild(menu);

  // ---------- 拖拽 + 点击 ----------
  let dragging = false;
  let moved = false;
  let startX = 0;
  let startY = 0;
  let origLeft = 0;
  let origTop = 0;

  const setOpen = (open: boolean) => {
    menu.hidden = !open;
    ball.classList.toggle('open', open);
    if (open) {
      // 打开时检测选中类型：选中表格才显示表格组
      void refreshTableGroupVisibility();
      const br = ball.getBoundingClientRect();
      const mw = menu.offsetWidth || 240;
      const mh = menu.offsetHeight || 300;
      let left = Math.min(Math.max(4, br.left + br.width / 2 - mw / 2), Math.max(4, window.innerWidth - mw - 4));
      let top = br.top - mh - 8;
      if (top < 4) top = br.bottom + 8;
      if (top + mh > window.innerHeight - 4) top = Math.max(4, window.innerHeight - mh - 4);
      menu.style.left = left + 'px';
      menu.style.top = top + 'px';
    }
  };
  const setMenuOpen = setOpen;

  async function refreshTableGroupVisibility(): Promise<void> {
    try {
      const target = await getSelectedTarget();
      tableGroup.hidden = !(target && target.kind === 'table');
    } catch {
      tableGroup.hidden = true;
    }
  }

  ball.addEventListener('pointerdown', (e) => {
    dragging = true;
    moved = false;
    startX = e.clientX;
    startY = e.clientY;
    const rect = ball.getBoundingClientRect();
    origLeft = rect.left;
    origTop = rect.top;
    ball.setPointerCapture(e.pointerId);
    ball.classList.add('dragging');
  });

  ball.addEventListener('pointermove', (e) => {
    if (!dragging) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) moved = true;
    if (!moved) return;
    const maxLeft = window.innerWidth - ball.offsetWidth - 4;
    const maxTop = window.innerHeight - ball.offsetHeight - 4;
    const left = Math.min(Math.max(4, origLeft + dx), Math.max(4, maxLeft));
    const top = Math.min(Math.max(4, origTop + dy), Math.max(4, maxTop));
    ball.style.left = left + 'px';
    ball.style.top = top + 'px';
  });

  ball.addEventListener('pointerup', (e) => {
    dragging = false;
    ball.classList.remove('dragging');
    if (!moved) {
      setMenuOpen(menu.hidden);
    }
  });

  ball.addEventListener('pointercancel', () => {
    dragging = false;
    ball.classList.remove('dragging');
  });

  // 点击空白处收起菜单
  document.addEventListener('click', (e) => {
    if (!menu.hidden && !menu.contains(e.target as Node) && !ball.contains(e.target as Node)) {
      setMenuOpen(false);
    }
  });
}
