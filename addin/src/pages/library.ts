import { Api, TemplateMeta } from '../api.js';
import { escapeHtml } from '../lib/html.js';
import { insertSlideBase64 } from '../office/writeSlide.js';
import { showToast, showModal, showPreviewModal } from '../ui.js';
import { renderTemplateDiagram } from '../lib/templateDiagram.js';

// 每排显示数量（1 / 2），本地记住用户选择；默认 1（模板竖向显示，一横排一个）
function getViewCols(): number {
  const v = Number(localStorage.getItem('tplLibCols'));
  return v === 2 ? 2 : 1;
}
function setViewCols(n: number): void { localStorage.setItem('tplLibCols', String(n)); }

// 模板按文件夹分组：未分类（根目录）排最前，其余按名称排序
function groupByFolder(list: TemplateMeta[]): { folder: string; items: TemplateMeta[] }[] {
  const map = new Map<string, TemplateMeta[]>();
  for (const t of list) {
    const k = t.folder || '';
    if (!map.has(k)) map.set(k, []);
    map.get(k)!.push(t);
  }
  const groups: { folder: string; items: TemplateMeta[] }[] = [];
  for (const [folder, items] of map) groups.push({ folder, items });
  groups.sort((a, b) => {
    if (a.folder === '') return -1;
    if (b.folder === '') return 1;
    return a.folder.localeCompare(b.folder, 'zh');
  });
  return groups;
}

function cardHtml(t: TemplateMeta, cols: number): string {
  const row = cols === 1;
  const verBadge = t.version ? ` <span class="ver-badge">v${t.version}</span>` : '';
  // 「编辑 / 版本 / 删除」按钮统一放在卡片右下角同一行；删除用垃圾桶图标（悬停有 title 说明）
  const editBtn = `<button class="secondary edit" data-id="${escapeHtml(t.id)}" data-folder="${escapeHtml(t.folder)}" title="编辑该模板（载入已保存的信息）">编辑</button>`;
  const verBtn = `<button class="secondary ver" data-id="${escapeHtml(t.id)}" data-folder="${escapeHtml(t.folder)}" title="版本历史 / 恢复 / 设为当前">版本</button>`;
  const delBtn = `<button class="secondary del" data-id="${escapeHtml(t.id)}" data-folder="${escapeHtml(t.folder)}" title="移入回收站">🗑️</button>`;
  const actions = `<div class="lib-card-actions">${editBtn}${verBtn}${delBtn}</div>`;
  // 缩略图容器：预览图 + 占位 + 放大查看按钮（图框右上角 ➕）
  const thumbHtml = `<div class="thumb-wrap">
      <img class="thumb" src="${t.preview}" alt="${escapeHtml(t.name)}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';this.parentElement.querySelector('.thumb-zoom').style.display='none';" />
      <div class="thumb thumb-placeholder" style="display:none"><span>${escapeHtml(t.name)}</span></div>
      <button class="thumb-zoom" data-preview="${escapeHtml(t.preview)}" title="放大查看预览">＋</button>
    </div>`;
  const inner = row
    ? `<div class="lib-row">${thumbHtml}
        <div class="lib-row-info">
          <b class="lib-name">${escapeHtml(t.name)}${verBadge}</b>
          <p class="hint">${t.updatedAt ? '更新于 ' + escapeHtml(t.updatedAt.slice(0, 10)) : ''}</p>
        </div>
        ${actions}
      </div>`
    : `${thumbHtml}
        <div class="lib-card-foot">
          <b class="lib-name">${escapeHtml(t.name)}${verBadge}</b>
          ${actions}
        </div>`;
  return `<div class="lib-card" data-id="${escapeHtml(t.id)}" data-folder="${escapeHtml(t.folder)}" style="cursor:pointer">${inner}</div>`;
}

// 版本历史弹窗：当前版本 / 版本历史 / 恢复（设为当前）/ 删除版本（自绘 modal，不用 window.confirm）
async function showVersionsModal(id: string, folder: string, onChanged: () => void): Promise<void> {
  let data: Awaited<ReturnType<typeof Api.listVersions>>;
  try {
    data = await Api.listVersions(id, folder);
  } catch (e) {
    await showModal({ title: '版本历史', message: '读取版本历史失败：' + ((e as Error).message || String(e)), buttons: [{ id: 'ok', label: '知道了' }] });
    return;
  }
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `<div class="modal-box ver-modal">
    <h3 class="modal-title">版本历史（当前 v${data.currentVersion ?? '?'}）</h3>
    <div class="ver-list"></div>
    <div class="modal-actions"><button class="secondary ver-close">关闭</button></div>
  </div>`;
  const listEl = overlay.querySelector('.ver-list') as HTMLElement;
  const close = () => overlay.remove();
  overlay.querySelector('.ver-close')!.addEventListener('click', close);
  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });

  const render = (): void => {
    listEl.innerHTML = data.versions.map((v) => {
      const previewUrl = '/api/templates/preview.png?folder=' + encodeURIComponent(folder) + '&id=' + encodeURIComponent(id) + '&version=' + encodeURIComponent(v.versionId);
      const when = (v.updatedAt || v.createdAt || '').slice(0, 16).replace('T', ' ');
      return `<div class="ver-item ${v.isCurrent ? 'ver-current' : ''}" data-vid="${escapeHtml(v.versionId)}">
        <img class="ver-thumb" src="${previewUrl}" alt="v${v.version}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
        <div class="ver-thumb ver-thumb-ph" style="display:none"><span>v${v.version}</span></div>
        <div class="ver-info">
          <b>v${v.version}${v.isCurrent ? ' <span class="ver-badge">当前</span>' : ''}</b>
          <p class="hint">${when}${v.changeNote ? '：' + escapeHtml(v.changeNote) : ''}</p>
        </div>
        <div class="ver-actions">
          ${v.isCurrent ? '' : '<button class="primary ver-set" title="恢复该版本为当前（内容与预览同步）">设为当前（恢复）</button>'}
          <button class="secondary ver-del" ${v.isCurrent ? 'disabled title="当前版本不能删除"' : 'title="删除该版本（不能删当前/唯一版本）"'}>删除</button>
        </div>
      </div>`;
    }).join('') || '<p class="hint">暂无版本记录。</p>';
  };
  render();

  // 设为当前（恢复）
  listEl.addEventListener('click', async (e) => {
    const setBtn = (e.target as HTMLElement).closest('.ver-set') as HTMLElement | null;
    const delBtn = (e.target as HTMLElement).closest('.ver-del') as HTMLElement | null;
    if (setBtn) {
      const versionId = setBtn.closest('.ver-item')!.getAttribute('data-vid')!;
      (setBtn as HTMLButtonElement).disabled = true;
      (setBtn as HTMLButtonElement).textContent = '恢复中…';
      try {
        await Api.setCurrentVersion(id, folder, versionId);
        showToast('已恢复为 v' + versionId.slice(1) + ' ✓', 1500);
        data = await Api.listVersions(id, folder);
        render();
        onChanged();
      } catch (err) {
        (setBtn as HTMLButtonElement).disabled = false;
        (setBtn as HTMLButtonElement).textContent = '设为当前（恢复）';
        await showModal({ title: '恢复失败', message: (err as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
      }
      return;
    }
    if (delBtn && !(delBtn as HTMLButtonElement).disabled) {
      const versionId = delBtn.closest('.ver-item')!.getAttribute('data-vid')!;
      if (!(delBtn as HTMLButtonElement).classList.contains('confirming')) {
        (delBtn as HTMLButtonElement).classList.add('confirming');
        (delBtn as HTMLButtonElement).textContent = '确认删除？';
        window.setTimeout(() => {
          (delBtn as HTMLButtonElement).classList.remove('confirming');
          (delBtn as HTMLButtonElement).textContent = '删除';
        }, 3000);
        return;
      }
      (delBtn as HTMLButtonElement).textContent = '删除中…';
      try {
        await Api.deleteVersion(id, folder, versionId);
        showToast('已删除 v' + versionId.slice(1), 1200);
        data = await Api.listVersions(id, folder);
        render();
        onChanged();
      } catch (err) {
        (delBtn as HTMLButtonElement).classList.remove('confirming');
        (delBtn as HTMLButtonElement).textContent = '删除';
        await showModal({ title: '删除失败', message: (err as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
      }
    }
  });

  document.body.appendChild(overlay);
}

// 回收站弹窗（P2-F）：列出已删除模板，支持恢复 / 彻底删除 / 清空
async function showRecycleModal(onChanged: () => void): Promise<void> {
  let data: Awaited<ReturnType<typeof Api.listRecycleBin>>;
  try {
    data = await Api.listRecycleBin();
  } catch (e) {
    await showModal({ title: '回收站', message: '读取回收站失败：' + ((e as Error).message || String(e)), buttons: [{ id: 'ok', label: '知道了' }] });
    return;
  }
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `<div class="modal-box ver-modal">
    <h3 class="modal-title">回收站<span class="hint">（删除的模板在这里，可恢复；彻底删除不可恢复）</span></h3>
    <div class="ver-list"></div>
    <div class="modal-actions">
      <button class="secondary rc-empty" style="margin-right:auto">清空回收站</button>
      <button class="secondary rc-close">关闭</button>
    </div>
  </div>`;
  const listEl = overlay.querySelector('.ver-list') as HTMLElement;
  const close = () => overlay.remove();
  overlay.querySelector('.rc-close')!.addEventListener('click', close);
  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });

  const render = (): void => {
    const items = data.items || [];
    listEl.innerHTML = items.map((it) => {
      const when = (it.deletedAt || '').slice(0, 16).replace('T', ' ');
      return `<div class="ver-item" data-entry="${escapeHtml(it.entryId)}">
        <img class="ver-thumb" src="${it.preview}" alt="${escapeHtml(it.name)}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
        <div class="ver-thumb ver-thumb-ph" style="display:none"><span>—</span></div>
        <div class="ver-info">
          <b>${escapeHtml(it.name)}</b>
          <p class="hint">原位置：${it.folder ? escapeHtml(it.folder) : '未分类（根目录）'}</p>
          <p class="hint">删除于 ${when}</p>
        </div>
        <div class="ver-actions">
          <button class="primary rc-restore">恢复</button>
          <button class="secondary rc-purge">彻底删除</button>
        </div>
      </div>`;
    }).join('') || '<p class="hint">回收站是空的。</p>';
  };
  render();

  // 恢复 / 彻底删除（两步确认）
  listEl.addEventListener('click', async (e) => {
    const restoreBtn = (e.target as HTMLElement).closest('.rc-restore') as HTMLElement | null;
    const purgeBtn = (e.target as HTMLElement).closest('.rc-purge') as HTMLElement | null;
    const entryId = (restoreBtn || purgeBtn)?.closest('.ver-item')?.getAttribute('data-entry') || '';
    if (restoreBtn) {
      (restoreBtn as HTMLButtonElement).disabled = true;
      (restoreBtn as HTMLButtonElement).textContent = '恢复中…';
      try {
        await Api.restoreRecycle(entryId);
        showToast('已恢复模板 ✓', 1500);
        data = await Api.listRecycleBin();
        render();
        onChanged();
      } catch (err) {
        (restoreBtn as HTMLButtonElement).disabled = false;
        (restoreBtn as HTMLButtonElement).textContent = '恢复';
        await showModal({ title: '恢复失败', message: (err as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
      }
      return;
    }
    if (purgeBtn) {
      if (!(purgeBtn as HTMLButtonElement).classList.contains('confirming')) {
        (purgeBtn as HTMLButtonElement).classList.add('confirming');
        (purgeBtn as HTMLButtonElement).textContent = '确认彻底删除？';
        window.setTimeout(() => {
          (purgeBtn as HTMLButtonElement).classList.remove('confirming');
          (purgeBtn as HTMLButtonElement).textContent = '彻底删除';
        }, 3000);
        return;
      }
      (purgeBtn as HTMLButtonElement).textContent = '删除中…';
      try {
        await Api.purgeRecycle(entryId);
        showToast('已彻底删除', 1200);
        data = await Api.listRecycleBin();
        render();
      } catch (err) {
        (purgeBtn as HTMLButtonElement).classList.remove('confirming');
        (purgeBtn as HTMLButtonElement).textContent = '彻底删除';
        await showModal({ title: '删除失败', message: (err as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
      }
    }
  });

  // 清空回收站（两步确认）
  const emptyBtn = overlay.querySelector('.rc-empty') as HTMLButtonElement;
  emptyBtn.addEventListener('click', async () => {
    if (!emptyBtn.classList.contains('confirming')) {
      emptyBtn.classList.add('confirming');
      emptyBtn.textContent = '确认清空全部？';
      window.setTimeout(() => { emptyBtn.classList.remove('confirming'); emptyBtn.textContent = '清空回收站'; }, 3000);
      return;
    }
    emptyBtn.textContent = '清空中…';
    try {
      const r = await Api.emptyRecycleBin();
      showToast('已清空回收站（' + r.removed + ' 项）', 1500);
      data = await Api.listRecycleBin();
      render();
    } catch (err) {
      emptyBtn.classList.remove('confirming');
      emptyBtn.textContent = '清空回收站';
      await showModal({ title: '清空失败', message: (err as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
    }
  });

  document.body.appendChild(overlay);
}

// AI 待写队列提示条：MCP/ChatGPT 生成的页面在这里一键写入当前 PPT
async function renderAiPendingBar(container: HTMLElement): Promise<void> {
  let items: { id: string; templateName: string; createdAt: string; written: boolean }[] = [];
  try { items = await Api.aiPendingList(); } catch { return; }
  const todo = items.filter((i) => !i.written);
  if (!todo.length) return;
  const bar = document.createElement('div');
  bar.className = 'card ai-pending-bar';
  bar.innerHTML = '<b>🤖 AI 已生成 ' + todo.length + ' 页，待写入当前 PPT</b>' +
    '<span style="margin-left:8px"><button class="secondary ai-clear-all" style="padding:4px 8px">清空全部</button></span>' +
    todo.map((t) =>
      '<div class="ai-pending-item"><span>' + escapeHtml(t.templateName) + '（' + escapeHtml((t.createdAt || '').slice(5, 16).replace('T', ' ')) + '）</span>' +
      '<span><button class="primary ai-write" data-id="' + escapeHtml(t.id) + '" style="padding:4px 10px">写入 PPT</button> ' +
      '<button class="secondary ai-discard" data-id="' + escapeHtml(t.id) + '" style="padding:4px 8px">忽略</button></span></div>').join('');
  container.prepend(bar);
  bar.querySelectorAll('.ai-write').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = (btn as HTMLElement).getAttribute('data-id')!;
      (btn as HTMLButtonElement).disabled = true;
      (btn as HTMLButtonElement).textContent = '写入中…';
      try {
        const entry = await Api.aiPendingGet(id);
        await insertSlideBase64(entry.base64);
        await Api.aiPendingWrite(id);
        showToast('已写入当前 PPT ✓', 1500);
        bar.remove();
      } catch (e) {
        (btn as HTMLButtonElement).disabled = false;
        (btn as HTMLButtonElement).textContent = '写入 PPT';
        const choice = await showModal({
          title: '写入失败',
          message: (e as Error).message + '\n可重试；或删除该待写项后在 ChatGPT 中重新生成。',
          buttons: [
            { id: 'cancel', label: '知道了' },
            { id: 'discard', label: '丢弃该项', kind: 'danger' }
          ]
        });
        if (choice === 'discard') {
          await Api.aiPendingDelete(id).catch(() => {});
          bar.remove();
        }
      }
    });
  });
  bar.querySelectorAll('.ai-discard').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = (btn as HTMLElement).getAttribute('data-id')!;
      await Api.aiPendingDelete(id).catch(() => {});
      bar.remove();
    });
  });
  bar.querySelector('.ai-clear-all')?.addEventListener('click', async () => {
    const choice = await showModal({
      title: '清空全部待写项',
      message: '将删除全部 ' + todo.length + ' 页待写内容（已写入的也会从队列移除，不影响 PPT）。确认？',
      buttons: [
        { id: 'cancel', label: '取消' },
        { id: 'clear', label: '清空全部', kind: 'danger' }
      ]
    });
    if (choice !== 'clear') return;
    try {
      const r = await Api.aiPendingClearAll();
      showToast('已清空 ' + r.removed + ' 项待写队列', 1500);
      bar.remove();
    } catch (e) {
      await showModal({ title: '清空失败', message: (e as Error).message, buttons: [{ id: 'ok', label: '知道了' }] });
    }
  });
}

export async function renderLibrary(container: HTMLElement): Promise<void> {
  container.innerHTML = '<div class="card"><p>正在加载模板库…</p></div>';
  let list: TemplateMeta[];
  try {
    list = await Api.listTemplates();
  } catch {
    container.innerHTML = `
      <div class="card">
        <p class="error">无法连接模板服务器，请确认本地服务已启动。</p>
        <p class="hint">模板列表加载失败，请稍后重试。</p>
      </div>`;
    return;
  }
  // 缺失预览图的模板：自动生成结构示意图并保存（updateCurrent，不产生新版本），成功后重拉列表
  const missing = list.filter((t) => t.hasPreview === false); // 明确缺失才生成（旧后端无此字段时不误判）
  if (missing.length) {
    showToast('正在生成 ' + missing.length + ' 个模板预览图…', 3000);
    let done = 0;
    for (const t of missing) {
      try {
        const doc = await Api.getTemplate(t.id, t.folder);
        const dataUrl = renderTemplateDiagram(doc.template);
        if (dataUrl) {
          await Api.saveTemplate({ name: t.name, folder: t.folder, template: doc.template, preview: dataUrl, updateCurrent: true });
          done++;
        }
      } catch { /* 单个模板生成失败不阻断其他 */ }
    }
    if (done) list = await Api.listTemplates();
  }
  const cols = getViewCols();
  const groups = groupByFolder(list);
  // 空列表也保留工具栏（回收站 / 视图切换），避免全部删除后按钮消失
  container.innerHTML = `
    <h1 class="page-title">模板库</h1>
    <div class="card lib-toolbar">
      <button class="secondary rc-open" title="查看已删除的模板（可恢复或彻底删除）">🗑️</button>
      <span class="view-toggle">
        <button class="secondary vt" data-cols="1" title="一排 1 个">1 列</button>
        <button class="secondary vt" data-cols="2" title="一排 2 个">2 列</button>
      </span>
    </div>
    ${list.length ? groups.map((g) => `
      <p class="folder-head">${g.folder ? escapeHtml(g.folder) : '未分类'}（${g.items.length}）</p>
      <div class="lib-grid cols-${cols}">${g.items.map((t) => cardHtml(t, cols)).join('')}</div>
    `).join('') : '<div class="card"><p>还没有模板。先在 PPT 里设计一页，然后点顶部「保存模板」。</p></div>'}`;
  await renderAiPendingBar(container);

  // 每排数量切换
  container.querySelectorAll('.view-toggle .vt').forEach((btn) => {
    const b = btn as HTMLButtonElement;
    b.classList.toggle('active', Number(b.getAttribute('data-cols')) === cols);
    b.addEventListener('click', () => {
      setViewCols(Number(b.getAttribute('data-cols')));
      renderLibrary(container);
    });
  });

  container.querySelectorAll('.lib-card[data-id]').forEach((card) => {
    card.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).classList.contains('del')) return;
      if ((e.target as HTMLElement).classList.contains('ver')) return;
      if ((e.target as HTMLElement).classList.contains('edit')) return;
      const id = card.getAttribute('data-id')!;
      const folder = card.getAttribute('data-folder') || '';
      sessionStorage.setItem('templateId', id);
      sessionStorage.setItem('templateFolder', folder);
      location.hash = '#wizard';
    });
  });
  // 放大查看预览（图框右上角 ➕ → 现有全屏预览弹窗）
  container.querySelectorAll('.thumb-zoom').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const url = btn.getAttribute('data-preview') || '';
      if (!url) return;
      showPreviewModal(`<div style="text-align:center"><img src="${url}" alt="模板预览" style="max-width:100%;max-height:70vh;border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.25);display:block;margin:0 auto" /></div>`);
    });
  });
  // 回收站弹窗
  container.querySelector('.rc-open')!.addEventListener('click', () => {
    showRecycleModal(() => renderLibrary(container));
  });
  // 编辑模板：跳转到保存模板界面并载入该模板信息
  container.querySelectorAll('.lib-card-actions .edit').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id')!;
      const folder = btn.getAttribute('data-folder') || '';
      sessionStorage.setItem('editTemplateId', id);
      sessionStorage.setItem('editTemplateFolder', folder);
      location.hash = '#save';
    });
  });
  // 版本历史弹窗
  container.querySelectorAll('.ver').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id')!;
      const folder = btn.getAttribute('data-folder') || '';
      await showVersionsModal(id, folder, () => renderLibrary(container));
    });
  });
  container.querySelectorAll('.del').forEach((btn) => {
    let timer: number | undefined;
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id')!;
      const folder = btn.getAttribute('data-folder') || '';
      // Office 任务窗格中 window.confirm 不可靠，改用两步内联确认。
      if (!btn.classList.contains('confirming')) {
        btn.classList.add('confirming');
        btn.textContent = '确认删除？';
        timer = window.setTimeout(() => {
          btn.classList.remove('confirming');
          btn.textContent = '🗑️';
        }, 3000);
        return;
      }
      if (timer) window.clearTimeout(timer);
      btn.classList.remove('confirming');
      btn.textContent = '移入中…';
      try {
        await Api.deleteTemplate(id, folder);
        showToast('已移到回收站，可随时恢复 ✓', 1800);
      } catch (err) {
        alert(`删除失败：${(err as Error).message}`);
        btn.textContent = '🗑️';
        return;
      }
      renderLibrary(container);
    });
  });
}
