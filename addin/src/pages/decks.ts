// 套版（Deck）管理页：列表 + 可视化新建/编辑（有序模板引用序列）
import { Api, TemplateMeta, VersionInfo } from '../api.js';
import { escapeHtml, escapeAttr } from '../lib/html.js';
import { showToast } from '../ui.js';

interface DeckRow {
  templateId: string;
  templateFolder: string;
  templateVersion: string; // '' = 最新
}

interface DeckMeta {
  id: string;
  name: string;
  folder: string;
  pageCount: number;
  preview: string;
  updatedAt: string | null;
}

type TemplatePick = DeckRow | null;

function templateKey(id: string, folder = ''): string {
  return folder + '\n' + id;
}

function versionPreviewUrl(templateId: string, folder: string, versionId: string): string {
  return '/api/templates/preview.png?folder=' + encodeURIComponent(folder)
    + '&id=' + encodeURIComponent(templateId)
    + '&version=' + encodeURIComponent(versionId);
}

function pageOrdinal(index: number): string {
  const ordinals = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩', '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳'];
  return ordinals[index] || String(index + 1);
}

export async function renderDecks(container: HTMLElement): Promise<void> {
  container.innerHTML = '<div class="card"><p>正在加载套版…</p></div>';
  let decks: DeckMeta[] = [];
  let templates: TemplateMeta[] = [];
  try {
    [decks, templates] = await Promise.all([Api.listDecks(), Api.listTemplates()]);
  } catch {
    container.innerHTML = '<div class="card"><p class="error">无法连接模板服务器。</p></div>';
    return;
  }

  const templateByKey = new Map(templates.map((t) => [templateKey(t.id, t.folder), t]));
  let editingId: string | null = null;
  let editingFolder = '';
  let name = '';
  let folder = '';
  let rows: DeckRow[] = [];
  const versionCache = new Map<string, VersionInfo[]>();

  const templateOf = (row: DeckRow): TemplateMeta | undefined => templateByKey.get(templateKey(row.templateId, row.templateFolder));
  const previewOf = (row: DeckRow): string => {
    const t = templateOf(row);
    if (row.templateVersion) return versionPreviewUrl(row.templateId, row.templateFolder, row.templateVersion);
    return t?.preview || '';
  };
  const firstPreviewOfDeck = (d: DeckMeta): string => d.preview || '';

  const moveRow = (from: number, to: number): void => {
    if (!Number.isInteger(from) || !Number.isInteger(to) || from === to || from < 0 || to < 0 || from >= rows.length || to >= rows.length) return;
    const [item] = rows.splice(from, 1);
    rows.splice(to, 0, item);
  };

  const loadVersions = async (templateId: string, templateFolder: string): Promise<VersionInfo[]> => {
    const key = templateKey(templateId, templateFolder);
    if (versionCache.has(key)) return versionCache.get(key)!;
    let versions: VersionInfo[] = [];
    try {
      versions = (await Api.listVersions(templateId, templateFolder)).versions || [];
    } catch { /* 无版本信息时只允许使用最新版本 */ }
    versionCache.set(key, versions);
    return versions;
  };

  const deckStackHtml = (preview: string, nameText: string, count: number, cls = ''): string => `
    <div class="deck-stack-preview ${cls}">
      <div class="deck-stack-layer layer-3"></div>
      <div class="deck-stack-layer layer-2"></div>
      <div class="deck-stack-front">
        ${preview
          ? `<img class="thumb" src="${escapeAttr(preview)}" alt="${escapeAttr(nameText)}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />`
          : ''}
        <div class="thumb thumb-placeholder" style="${preview ? 'display:none' : 'display:flex'}"><span>${escapeHtml(nameText)}</span></div>
      </div>
      ${count > 1 ? `<span class="deck-stack-count">${count}</span>` : ''}
    </div>`;

  const deckCardHtml = (d: DeckMeta): string => `
    <div class="card deck-item">
      <div class="deck-card-main">
        <div class="deck-preview-wrap deck-preview-landscape">
          ${deckStackHtml(firstPreviewOfDeck(d), d.name, d.pageCount)}
          <button class="thumb-zoom deck-preview-expand" data-id="${escapeAttr(d.id)}" data-folder="${escapeAttr(d.folder)}" title="展开浏览">＋</button>
        </div>
        <div class="deck-card-info">
          <b class="lib-name">${escapeHtml(d.name)}${d.folder ? '（' + escapeHtml(d.folder) + '）' : ''}</b>
          <p class="hint">${d.pageCount} 页${d.updatedAt ? ' · 更新于 ' + escapeHtml(d.updatedAt.slice(0, 10)) : ''}</p>
          <div class="deck-card-actions">
            <button class="primary deck-use" data-id="${escapeAttr(d.id)}" data-folder="${escapeAttr(d.folder)}">使用</button>
            <button class="secondary deck-edit" data-id="${escapeAttr(d.id)}" data-folder="${escapeAttr(d.folder)}">编辑</button>
            <button class="secondary deck-del" data-id="${escapeAttr(d.id)}" data-folder="${escapeAttr(d.folder)}" title="删除套版">🗑️</button>
          </div>
        </div>
      </div>
    </div>`;

  const pageCardHtml = (row: DeckRow, i: number): string => {
    const t = templateOf(row);
    const preview = previewOf(row);
    const ver = row.templateVersion ? ` <span class="ver-badge">${escapeHtml(row.templateVersion)}</span>` : '';
    return `<div class="deck-page-card" data-i="${i}" draggable="true">
      <span class="deck-page-number">${pageOrdinal(i)}</span>
      <div class="thumb-wrap">
        <img class="thumb" src="${escapeAttr(preview)}" alt="${escapeAttr(t?.name || '模板')}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
        <div class="thumb thumb-placeholder" style="display:none"><span>${escapeHtml(t?.name || '模板')}</span></div>
      </div>
      <div class="deck-page-foot">
        <b>${escapeHtml(t?.name || row.templateId)}${ver}</b>
        <button class="secondary row-del" data-i="${i}" title="移除该页">删除</button>
      </div>
    </div>`;
  };

  const render = (): void => {
    container.innerHTML = `
      <h1 class="page-title">套版</h1>
      <div class="card lib-toolbar">
        <span class="hint">${templates.length ? '把多个模板按顺序组成一份 PPT。' : '模板库为空，请先保存模板。'}</span>
        <span style="display:flex;gap:6px;align-items:center">
          <button class="secondary deck-rc-open" title="查看已删除的套版（可恢复或彻底删除）">🗑️</button>
          <button class="primary" id="deck-new" ${templates.length ? '' : 'disabled'}>新建</button>
        </span>
      </div>
      <div id="deck-form" style="display:none"></div>
      <div id="deck-list">${decks.length
        ? decks.map(deckCardHtml).join('')
        : '<div class="card"><p class="hint">还没有套版。点「新建」把多个模板串成一份。</p></div>'}
      </div>`;

    container.querySelector('#deck-new')?.addEventListener('click', () => {
      editingId = null;
      editingFolder = '';
      name = '';
      folder = '';
      rows = [];
      renderForm();
    });
    container.querySelector('.deck-rc-open')?.addEventListener('click', () => {
      void showDeckRecycleModal();
    });
    container.querySelectorAll('.deck-use').forEach((btn) => {
      btn.addEventListener('click', () => {
        sessionStorage.setItem('deckId', (btn as HTMLElement).getAttribute('data-id')!);
        sessionStorage.setItem('deckFolder', (btn as HTMLElement).getAttribute('data-folder') || '');
        location.hash = '#deck-wizard';
      });
    });
    container.querySelectorAll('.deck-edit').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = (btn as HTMLElement).getAttribute('data-id')!;
        const f = (btn as HTMLElement).getAttribute('data-folder') || '';
        try {
          const d = await Api.getDeck(id, f);
          editingId = id;
          editingFolder = f;
          name = d.deck.name;
          folder = f;
          rows = (d.deck.pages || []).map((p) => ({
            templateId: p.templateId,
            templateFolder: p.templateFolder || '',
            templateVersion: p.templateVersion || ''
          }));
          renderForm();
        } catch (e) { showToast('读取套版失败：' + (e as Error).message, 3000); }
      });
    });
    container.querySelectorAll('.deck-del').forEach((btn) => {
      let timer: number | undefined;
      btn.addEventListener('click', async () => {
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
        const id = (btn as HTMLElement).getAttribute('data-id')!;
        const f = (btn as HTMLElement).getAttribute('data-folder') || '';
        try {
          await Api.deleteDeck(id, f);
          showToast('已删除', 1200);
          renderDecks(container);
        } catch (e) { showToast('删除失败：' + (e as Error).message, 3000); }
      });
    });
    container.querySelectorAll('.deck-preview-expand').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = (btn as HTMLElement).getAttribute('data-id')!;
        const f = (btn as HTMLElement).getAttribute('data-folder') || '';
        try {
          const d = await Api.getDeck(id, f);
          showDeckPreviewModal((d.deck.pages || []).map((p) => ({
            templateId: p.templateId,
            templateFolder: p.templateFolder || '',
            templateVersion: p.templateVersion || ''
          })));
        } catch (e) { showToast('展开失败：' + (e as Error).message, 3000); }
      });
    });
  };

  const renderForm = (): void => {
    const formEl = container.querySelector('#deck-form') as HTMLElement;
    formEl.style.display = '';
    formEl.innerHTML = `
      <div class="card deck-editor">
        <div class="deck-editor-head">
          <b>${editingId ? '编辑' : '新建'}</b>
          <div class="deck-editor-actions">
            <button class="primary" id="deck-save">保存</button>
            <button class="secondary" id="deck-cancel">取消</button>
          </div>
        </div>
        <label>名称</label>
        <input id="deck-name" value="${escapeAttr(name)}" placeholder="例如：周报（封面+正文×3）" />
        <div class="deck-seq-head">
          <b>页面序列</b>
          <button class="secondary" id="deck-add-row">+ 添加页面</button>
        </div>
        <div id="deck-rows"></div>
      </div>`;

    const rowsEl = formEl.querySelector('#deck-rows') as HTMLElement;
    const renderRows = (): void => {
      rowsEl.innerHTML = rows.length
        ? `<div class="deck-page-grid">${rows.map(pageCardHtml).join('')}</div>`
        : '<p class="hint">还没有页面。</p>';
      rowsEl.querySelectorAll('.row-del').forEach((btn) => {
        btn.addEventListener('click', () => {
          const i = Number((btn as HTMLElement).getAttribute('data-i'));
          rows.splice(i, 1);
          renderRows();
        });
      });
      rowsEl.querySelectorAll('.deck-page-card').forEach((card) => {
        card.addEventListener('dragstart', (e) => {
          (e as DragEvent).dataTransfer?.setData('text/plain', (card as HTMLElement).getAttribute('data-i') || '');
          (card as HTMLElement).classList.add('dragging');
        });
        card.addEventListener('dragend', () => {
          (card as HTMLElement).classList.remove('dragging');
          rowsEl.querySelectorAll('.drag-over').forEach((el) => el.classList.remove('drag-over'));
        });
        card.addEventListener('dragover', (e) => {
          e.preventDefault();
          (card as HTMLElement).classList.add('drag-over');
        });
        card.addEventListener('dragleave', () => (card as HTMLElement).classList.remove('drag-over'));
        card.addEventListener('drop', (e) => {
          e.preventDefault();
          const from = Number((e as DragEvent).dataTransfer?.getData('text/plain'));
          const to = Number((card as HTMLElement).getAttribute('data-i'));
          moveRow(from, to);
          renderRows();
        });
      });
    };
    renderRows();

    formEl.querySelector('#deck-add-row')!.addEventListener('click', async () => {
      const pick = await openTemplatePicker();
      if (!pick) return;
      rows.push(pick);
      renderRows();
    });
    formEl.querySelector('#deck-cancel')!.addEventListener('click', () => { formEl.style.display = 'none'; });
    formEl.querySelector('#deck-save')!.addEventListener('click', async () => {
      const nameInput = formEl.querySelector('#deck-name') as HTMLInputElement;
      name = nameInput.value.trim();
      if (!name) { showToast('请输入名称', 2000); nameInput.focus(); return; }
      folder = '';
      if (!rows.length) { showToast('请至少添加一个页面', 2000); return; }
      try {
        const firstPreview = previewOf(rows[0]);
        await Api.saveDeck({
          name,
          folder,
          preview: firstPreview,
          deck: {
            schemaVersion: 1,
            name,
            pages: rows.map((r) => ({
              templateId: r.templateId,
              templateFolder: r.templateFolder || undefined,
              templateVersion: r.templateVersion || undefined
            }))
          }
        });
        showToast('套版已保存', 1500);
        formEl.style.display = 'none';
        renderDecks(container);
      } catch (e) {
        showToast('保存失败：' + (e as Error).message, 4000);
      }
    });
  };

  const openTemplatePicker = (): Promise<TemplatePick> => {
    return new Promise((resolve) => {
      let selected: TemplateMeta | null = templates[0] || null;
      let selectedVersion = '';
      const overlay = document.createElement('div');
      overlay.className = 'modal-overlay';
      overlay.innerHTML = `<div class="modal-box deck-picker-modal" role="dialog" aria-modal="true">
        <h3 class="modal-title">添加页面</h3>
        <div class="deck-picker-grid">
          ${templates.map((t) => `<button class="deck-picker-card" data-id="${escapeAttr(t.id)}" data-folder="${escapeAttr(t.folder)}">
            <img class="thumb" src="${escapeAttr(t.preview)}" alt="${escapeAttr(t.name)}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
            <div class="thumb thumb-placeholder" style="display:none"><span>${escapeHtml(t.name)}</span></div>
            <b>${escapeHtml(t.name)}</b>
            <span class="hint">${t.folder ? escapeHtml(t.folder) : '未分类'}${t.version ? ' · v' + t.version : ''}</span>
          </button>`).join('')}
        </div>
        <div class="deck-picker-version">
          <label>版本</label>
          <select class="deck-picker-version-select"><option value="">最新版本</option></select>
        </div>
        <div class="modal-actions">
          <button class="secondary deck-picker-cancel">取消</button>
          <button class="primary deck-picker-ok">确认</button>
        </div>
      </div>`;
      const cards = Array.from(overlay.querySelectorAll('.deck-picker-card')) as HTMLButtonElement[];
      const versionSel = overlay.querySelector('.deck-picker-version-select') as HTMLSelectElement;
      const close = (value: TemplatePick): void => { overlay.remove(); resolve(value); };
      const syncCards = (): void => {
        cards.forEach((card) => {
          card.classList.toggle(
            'selected',
            card.getAttribute('data-id') === selected?.id && (card.getAttribute('data-folder') || '') === (selected?.folder || '')
          );
        });
      };
      const syncVersions = async (): Promise<void> => {
        selectedVersion = '';
        const versions = selected ? await loadVersions(selected.id, selected.folder || '') : [];
        versionSel.innerHTML = '<option value="">最新版本</option>' + versions.map((v) =>
          '<option value="' + escapeAttr(v.versionId) + '">' + escapeHtml('v' + v.version + (v.isCurrent ? '（当前）' : '')) + '</option>'
        ).join('');
      };
      cards.forEach((card) => {
        card.addEventListener('click', async () => {
          selected = templates.find((t) => t.id === card.getAttribute('data-id') && (t.folder || '') === (card.getAttribute('data-folder') || '')) || null;
          syncCards();
          await syncVersions();
        });
      });
      versionSel.addEventListener('change', () => { selectedVersion = versionSel.value; });
      overlay.querySelector('.deck-picker-cancel')!.addEventListener('click', () => close(null));
      overlay.querySelector('.deck-picker-ok')!.addEventListener('click', () => {
        if (!selected) { close(null); return; }
        close({ templateId: selected.id, templateFolder: selected.folder || '', templateVersion: selectedVersion });
      });
      overlay.addEventListener('click', (e) => { if (e.target === overlay) close(null); });
      document.body.appendChild(overlay);
      syncCards();
      void syncVersions();
    });
  };

  const showDeckPreviewModal = (pages: DeckRow[]): void => {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay preview-overlay';
    overlay.innerHTML = `<div class="modal-box preview-box deck-pages-modal">
      <div class="preview-head"><span class="preview-title">套版页面</span><button class="preview-close" aria-label="关闭预览">×</button></div>
      <div class="preview-body deck-pages-list">
        ${pages.map((p, i) => {
          const t = templateOf(p);
          const preview = previewOf(p);
          return `<div class="deck-page-wide">
            <img class="thumb" src="${escapeAttr(preview)}" alt="${escapeAttr(t?.name || '页面')}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
            <div class="thumb thumb-placeholder" style="display:none"><span>${escapeHtml(t?.name || '页面')}</span></div>
            <div class="deck-page-wide-info">
              <b>第 ${i + 1} 页 · ${escapeHtml(t?.name || p.templateId)}</b>
              <p class="hint">${p.templateVersion ? escapeHtml(p.templateVersion) : '最新版本'}</p>
            </div>
          </div>`;
        }).join('') || '<p class="hint">没有页面。</p>'}
      </div>
    </div>`;
    const close = () => overlay.remove();
    overlay.querySelector('.preview-close')!.addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    document.body.appendChild(overlay);
  };

  const showDeckRecycleModal = async (): Promise<void> => {
    let data: Awaited<ReturnType<typeof Api.listDeckRecycleBin>>;
    try {
      data = await Api.listDeckRecycleBin();
    } catch (e) {
      showToast('读取套版回收站失败：' + (e as Error).message, 3000);
      return;
    }
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `<div class="modal-box ver-modal">
      <h3 class="modal-title">套版回收站</h3>
      <div class="ver-list deck-rc-list"></div>
      <div class="modal-actions">
        <button class="secondary deck-rc-empty" style="margin-right:auto">清空回收站</button>
        <button class="secondary deck-rc-close">关闭</button>
      </div>
    </div>`;
    const listEl = overlay.querySelector('.deck-rc-list') as HTMLElement;
    const close = () => overlay.remove();
    const renderRecycle = (): void => {
      const items = data.items || [];
      listEl.innerHTML = items.map((it) => `<div class="ver-item" data-entry="${escapeAttr(it.entryId)}">
        <img class="ver-thumb" src="${escapeAttr(it.preview)}" alt="${escapeAttr(it.name)}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" />
        <div class="ver-thumb ver-thumb-ph" style="display:none"><span>—</span></div>
        <div class="ver-info">
          <b>${escapeHtml(it.name)}</b>
          <p class="hint">${it.pageCount} 页${it.deletedAt ? ' · 删除于 ' + escapeHtml(it.deletedAt.slice(0, 10)) : ''}</p>
        </div>
        <div class="ver-actions">
          <button class="primary deck-rc-restore">恢复</button>
          <button class="secondary deck-rc-purge">彻底删除</button>
        </div>
      </div>`).join('') || '<p class="hint">回收站是空的。</p>';
    };
    renderRecycle();
    overlay.querySelector('.deck-rc-close')!.addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    listEl.addEventListener('click', async (e) => {
      const restore = (e.target as HTMLElement).closest('.deck-rc-restore') as HTMLButtonElement | null;
      const purge = (e.target as HTMLElement).closest('.deck-rc-purge') as HTMLButtonElement | null;
      const entryId = (restore || purge)?.closest('.ver-item')?.getAttribute('data-entry') || '';
      if (restore) {
        restore.disabled = true;
        restore.textContent = '恢复中…';
        try {
          await Api.restoreDeckRecycle(entryId);
          data = await Api.listDeckRecycleBin();
          renderRecycle();
          renderDecks(container);
        } catch (err) {
          restore.disabled = false;
          restore.textContent = '恢复';
          showToast('恢复失败：' + (err as Error).message, 3000);
        }
      }
      if (purge) {
        if (!purge.classList.contains('confirming')) {
          purge.classList.add('confirming');
          purge.textContent = '确认彻底删除？';
          window.setTimeout(() => { purge.classList.remove('confirming'); purge.textContent = '彻底删除'; }, 3000);
          return;
        }
        try {
          await Api.purgeDeckRecycle(entryId);
          data = await Api.listDeckRecycleBin();
          renderRecycle();
        } catch (err) {
          purge.classList.remove('confirming');
          purge.textContent = '彻底删除';
          showToast('删除失败：' + (err as Error).message, 3000);
        }
      }
    });
    overlay.querySelector('.deck-rc-empty')!.addEventListener('click', async () => {
      const btn = overlay.querySelector('.deck-rc-empty') as HTMLButtonElement;
      if (!btn.classList.contains('confirming')) {
        btn.classList.add('confirming');
        btn.textContent = '确认清空全部？';
        window.setTimeout(() => { btn.classList.remove('confirming'); btn.textContent = '清空回收站'; }, 3000);
        return;
      }
      try {
        await Api.emptyDeckRecycleBin();
        data = await Api.listDeckRecycleBin();
        renderRecycle();
      } catch (err) {
        showToast('清空失败：' + (err as Error).message, 3000);
      }
    });
    document.body.appendChild(overlay);
  };

  render();
}
