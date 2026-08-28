// 套版生成向导：顶部复用普通生成向导的控件结构，页面内容按套版页分组。
import { Api, TemplateDoc, TemplateShape, ImageResult, getDefaultImageProvider } from '../api.js';
import { renderPreview } from '../previewPanel.js';
import { escapeHtml, escapeAttr } from '../lib/html.js';
import { showToast, showModal, showPreviewModal } from '../ui.js';
import { openImageCropEditor } from '../lib/cropEditor.js';
import { shapeConstraints, constraintSummary, semanticRoleLabel, sortShapesByPosition } from '../lib/semantic.js';
import { infoTip } from '../lib/tooltip.js';
import { loadOutputMode, resolveSlotMode, LIMIT_CHOICES } from '../lib/outputMode.js';
import type { OutputMode } from '../lib/outputMode.js';
import { insertSlideBase64 } from '../office/writeSlide.js';
import { localTextToLatex, todayStr } from '../lib/formula.js';

interface PageState {
  templateId: string;
  template: TemplateDoc;
  templateFolder: string;
  templateVersion?: string;
  templateName: string;
  vars: Record<string, string>;
  texts: Record<string, string>;
  prompts: Record<string, string>;
  formulas: Record<string, string>;
  imageDataUrl: string;
  imagePrompt: string;
  selected: string;
  imgState: { q: string; images: ImageResult[]; error: string };
  outputModes: Record<string, OutputMode>;
  tableData: Record<string, string[][]>;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const ROLE_ORDER: Record<string, number> = { seq: 0, title: 1, subtitle: 2, body: 3, caption: 4, formula: 5, date: 6 };

function cloneOutputMode(m: OutputMode): OutputMode {
  return { plain: m.plain, bullets: m.bullets, condense: m.condense, maxChars: m.maxChars, touched: m.touched };
}

function sortTextSlots(shapes: TemplateShape[]): TemplateShape[] {
  return [...shapes].sort((a, b) => {
    const pa = a.role === 'manual_var' ? 8 : (ROLE_ORDER[a.semanticRole || ''] ?? 7);
    const pb = b.role === 'manual_var' ? 8 : (ROLE_ORDER[b.semanticRole || ''] ?? 7);
    return pa - pb;
  });
}

function defaultTableData(s: TemplateShape): string[][] {
  const t = s.table;
  if (!t) return [[]];
  const rows = t.rows || 1;
  const cols = t.cols || 1;
  const out: string[][] = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));
  for (const c of t.cells || []) {
    if (c.row >= 0 && c.row < rows && c.col >= 0 && c.col < cols) out[c.row][c.col] = c.text || '';
  }
  return out;
}

function modeForPageSlot(page: PageState, slotId: string): OutputMode {
  if (!page.outputModes[slotId]) page.outputModes[slotId] = cloneOutputMode(loadOutputMode());
  return page.outputModes[slotId];
}

function backgroundPanelHtml(): string {
  return `
    <details class="wb-adv gp-adv deck-bg-adv">
      <summary>背景${infoTip('套版生成时默认沿用每个模板保存的背景。')}</summary>
      <div class="page-info wb-bg-panel">
        <div class="pi-row">
          <span class="pi-k">套版背景</span>
          <span class="pi-v wb-bg-desc">跟随各模板</span>
        </div>
        <p class="pi-hint">整份套版默认沿用每个模板保存时的背景。</p>
      </div>
    </details>`;
}

function readLocalImage(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function parseTableText(raw: string): string[][] {
  return raw.split(/\r?\n/)
    .map((line) => line.trimEnd())
    .filter((line) => line.length)
    .map((line) => line.includes('\t') ? line.split('\t') : line.split(','))
    .map((row) => row.map((cell) => cell.trim()));
}

export async function renderDeckWizard(container: HTMLElement): Promise<void> {
  const id = sessionStorage.getItem('deckId');
  if (!id) {
    container.innerHTML = '<div class="card"><p>请先在「套版」页选择要使用的套版。</p><button class="primary" id="go-decks">去套版管理</button></div>';
    container.querySelector('#go-decks')!.addEventListener('click', () => { location.hash = '#decks'; });
    return;
  }

  let deck;
  try {
    deck = await Api.getDeck(id, sessionStorage.getItem('deckFolder') || '');
  } catch (e) {
    container.innerHTML = '<div class="card"><p class="error">加载套版失败：' + escapeHtml((e as Error).message) + '</p></div>';
    return;
  }

  let globalTheme = '';
  const pages: PageState[] = [];
  for (const spec of deck.deck.pages) {
    let template: TemplateDoc;
    if (spec.templateVersion) {
      const gv = await Api.getVersion(spec.templateId, spec.templateFolder || '', spec.templateVersion);
      template = gv.version;
    } else {
      const t = await Api.getTemplate(spec.templateId, spec.templateFolder || '');
      template = t.template;
    }
    template.shapes = sortShapesByPosition(template.shapes);
    const textSlots = sortTextSlots(template.shapes.filter((s) => s.role === 'ai_text' || s.role === 'manual_var'));
    const tableShapes = template.shapes.filter((s) => s.role === 'table' && s.table);
    const texts: Record<string, string> = {};
    const prompts: Record<string, string> = {};
    const tableData: Record<string, string[][]> = {};
    for (const s of textSlots) {
      if (s.semanticRole === 'date') texts[s.id] = todayStr();
      if (s.prompt && s.semanticRole !== 'formula') prompts[s.id] = s.prompt;
    }
    for (const s of tableShapes) tableData[s.id] = defaultTableData(s);
    pages.push({
      templateId: spec.templateId,
      template,
      templateFolder: spec.templateFolder || '',
      templateVersion: spec.templateVersion,
      templateName: template.name || spec.templateId,
      vars: { ...(spec.variables || {}) },
      texts,
      prompts,
      formulas: {},
      imageDataUrl: '',
      imagePrompt: spec.image?.prompt || '',
      selected: '',
      imgState: { q: spec.image?.prompt || '', images: [], error: '' },
      outputModes: {},
      tableData
    });
  }

  const textSlotHtml = (page: PageState, pi: number, s: TemplateShape, i: number): string => {
    const isVar = s.role === 'manual_var';
    const cur = isVar ? (page.vars[s.id] ?? '') : (page.texts[s.id] ?? '');
    const roleLabel = semanticRoleLabel(s.semanticRole);
    const label = isVar
      ? (s.varName ? '变量「' + escapeHtml(s.varName) + '」' : '变量 ' + (i + 1))
      : (roleLabel || '文本位 ' + (i + 1)) + (s.name ? '：' + escapeHtml(s.name) : '');
    const cons = constraintSummary(s);
    const lab = `<label class="wb-slot-label"><b>${label}</b>${cons ? ` <span class="wb-constraint">${escapeHtml(cons)}</span>` : ''}</label>`;
    const status = `<span class="hint slot-status" data-p="${pi}" data-id="${s.id}"></span>`;
    const sem = s.semanticRole;
    if (sem === 'seq') {
      return `<div class="wb-text-slot" data-p="${pi}" data-id="${s.id}">
        ${lab}
        <div class="wb-seq-row" data-p="${pi}" data-id="${s.id}">
          ${['01', '02', '03', '04', '05', '06'].map((n) => `<button type="button" class="om-chip seq-chip${cur === n ? ' active' : ''}" data-p="${pi}" data-id="${s.id}" data-seq="${n}">${n}</button>`).join('')}
          <input class="seq-input" data-p="${pi}" data-id="${s.id}" inputmode="numeric" pattern="[0-9]*" maxlength="6" placeholder="自定义" value="${escapeAttr(/^\d*$/.test(cur) ? cur : '')}" title="仅限阿拉伯数字" />
        </div>
        ${status}
      </div>`;
    }
    if (sem === 'date') {
      return `<div class="wb-text-slot" data-p="${pi}" data-id="${s.id}">
        ${lab}
        <div class="wb-date-val" data-p="${pi}" data-id="${s.id}">${escapeHtml(page.texts[s.id] || todayStr())}</div>
        ${status}
      </div>`;
    }
    if (sem === 'caption') {
      return `<div class="wb-text-slot" data-p="${pi}" data-id="${s.id}">
        ${lab}
        <div class="wb-slot-row"><input class="gen-in wb-gen" data-p="${pi}" data-gen="${s.id}" value="${escapeAttr(cur)}" placeholder="填写图片诠释；图片搜索的关键词会自动填入（留空则不带）" /></div>
        ${status}
      </div>`;
    }
    if (sem === 'formula') {
      return `<div class="wb-text-slot" data-p="${pi}" data-id="${s.id}">
        ${lab}
        <div class="wb-slot-row"><textarea class="gen-ta wb-gen" data-p="${pi}" data-gen="${s.id}" rows="2" placeholder="粘贴或输入公式">${escapeAttr(cur)}</textarea></div>
        <details class="wb-adv"><summary>AI服务</summary>
          <p class="hint" style="margin:4px 0">内置提示词：把输入的公式/描述转换为完整规范的专业型公式（输出 LaTeX），生成 PPT 时转为专业型数学排版。</p>
          <div class="wb-adv-actions"><button class="secondary slot-gen" data-p="${pi}" data-id="${s.id}">AI 生成公式</button></div>
        </details>
        ${status}
      </div>`;
    }
    const mode = modeForPageSlot(page, s.id);
    const modeChips = `<span class="wb-mode-chips">
      <button type="button" class="om-chip${mode.plain ? ' active' : ''}" data-p="${pi}" data-slot="${s.id}" data-om="plain" title="整段：一段连续正文，不用 Markdown 和列表符号">整段</button>
      <button type="button" class="om-chip${mode.bullets ? ' active' : ''}" data-p="${pi}" data-slot="${s.id}" data-om="bullets" title="分点：每个要点单独一行">分点</button>
      <button type="button" class="om-chip${mode.condense ? ' active' : ''}" data-p="${pi}" data-slot="${s.id}" data-om="condense" title="精简：删冗余、句子短">精简</button>
      <select class="om-limit" data-p="${pi}" data-slot="${s.id}" title="限制 AI 输出总字数">${LIMIT_CHOICES.map((n) => `<option value="${n}"${mode.maxChars === n ? ' selected' : ''}>${n === 0 ? '不限字数' : '≤' + n + ' 字'}</option>`).join('')}</select>
    </span>`;
    return `<div class="wb-text-slot" data-p="${pi}" data-id="${s.id}">
      ${lab}
      <div class="wb-slot-row"><textarea class="gen-ta wb-gen" data-p="${pi}" data-gen="${s.id}" rows="1" placeholder="输入内容，或展开「AI服务」生成">${escapeAttr(cur)}</textarea></div>
      <details class="wb-adv"><summary>AI服务</summary>
        <textarea class="slot-prompt" data-p="${pi}" data-id="${s.id}" rows="2" placeholder="为该段单独写要求（可选），留空则用主题生成">${escapeAttr(page.prompts[s.id] || '')}</textarea>
        <div class="wb-adv-row">${modeChips}<span class="wb-adv-actions"><button class="secondary slot-gen" data-p="${pi}" data-id="${s.id}">AI 生成</button></span></div>
      </details>
      ${status}
    </div>`;
  };

  const imgSlotHtml = (page: PageState, pi: number, s: TemplateShape, i: number): string => {
    const chosen = page.imageDataUrl || '';
    return `<div class="img-slot wb-img-slot" data-p="${pi}" data-id="${s.id}">
      <b>图片位 ${i + 1}${s.name ? '：' + escapeHtml(s.name) : ''}</b>
      <div class="img-chosen" data-p="${pi}" data-id="${s.id}" style="margin:4px 0${chosen ? '' : ';display:none'}">
        <img src="${chosen}" style="max-width:120px;max-height:80px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />
        <button class="secondary img-crop" data-p="${pi}" data-id="${s.id}" title="按图片位比例重新裁剪">✂ 裁剪</button>
        <button class="secondary img-clear" data-p="${pi}" data-id="${s.id}">✕ 清除</button>
      </div>
      <div class="wb-img-actions">
        <button class="secondary img-pick" data-p="${pi}" data-id="${s.id}" title="从文件夹选择本地图片">本地</button>
        <input class="img-q" data-p="${pi}" data-id="${s.id}" value="${escapeAttr(page.imgState.q || page.imagePrompt)}" placeholder="描述图片，如：深蓝色科技感机房" />
        <button class="secondary img-search" data-p="${pi}" data-id="${s.id}">搜索</button>
      </div>
      <input type="file" class="img-file" data-p="${pi}" data-id="${s.id}" accept="image/*" style="display:none" />
      <div class="img-dropzone" data-p="${pi}" data-id="${s.id}"></div>
      <div class="img-results" data-p="${pi}" data-id="${s.id}" style="display:grid;grid-template-columns:repeat(3,1fr);gap:4px"></div>
      <div class="img-nav" data-p="${pi}" data-id="${s.id}"></div>
    </div>`;
  };

  const tableSlotHtml = (page: PageState, pi: number, s: TemplateShape, ti: number): string => {
    const data = page.tableData[s.id] || defaultTableData(s);
    const nc = Math.max(1, (data[0] || []).length);
    return `<div class="wb-tbl-slot" data-p="${pi}" data-tbl="${s.id}">
      <div class="wb-tbl-summary">
        <span class="wb-tbl-info"><b>表格 ${ti + 1}${s.name ? '：' + escapeHtml(s.name) : ''}</b>　<span class="wb-tbl-size">已导入 ${data.length} × ${nc} 数据</span></span>
      </div>
      <div class="wb-tbl-editor" data-p="${pi}" data-tbl="${s.id}" style="display:block">
        <div class="tbl-tabs" data-p="${pi}" data-tbl="${s.id}" style="display:flex;gap:4px;margin:6px 0">
          <button class="secondary tbl-tab active" data-p="${pi}" data-tbl="${s.id}" data-mode="manual">手动编辑</button>
          <button class="secondary tbl-tab" data-p="${pi}" data-tbl="${s.id}" data-mode="paste">粘贴</button>
          <button class="secondary tbl-tab" data-p="${pi}" data-tbl="${s.id}" data-mode="ai">AI 生成</button>
        </div>
        <div class="tbl-panel" data-p="${pi}" data-tbl="${s.id}" data-mode="manual" style="display:block">
          <div class="tbl-grid" style="display:grid;grid-template-columns:repeat(${nc}, 1fr);gap:2px;margin:4px 0">
            ${data.map((row, r) => row.map((cell, c) =>
              `<textarea class="tbl-cell" data-p="${pi}" data-tbl="${s.id}" data-r="${r}" data-c="${c}" rows="1" placeholder=" ">${escapeAttr(cell)}</textarea>`).join('')).join('')}
          </div>
        </div>
        <div class="tbl-panel" data-p="${pi}" data-tbl="${s.id}" data-mode="paste" style="display:none">
          <textarea class="tbl-paste-ta" data-p="${pi}" data-tbl="${s.id}" rows="6" placeholder="在此 Ctrl+V 粘贴 Excel/网页复制的表格（按制表符/逗号分隔解析）"></textarea>
          <div style="display:flex;gap:6px;margin-top:6px">
            <button class="secondary tbl-paste-do" data-p="${pi}" data-tbl="${s.id}">导入到表格</button>
            <button class="secondary tbl-paste-clear" data-p="${pi}" data-tbl="${s.id}">清空全部单元格</button>
          </div>
        </div>
        <div class="tbl-panel" data-p="${pi}" data-tbl="${s.id}" data-mode="ai" style="display:none">
          <textarea class="tbl-ai-req" data-p="${pi}" data-tbl="${s.id}" rows="2" placeholder="生成要求（可选），如：补充 5 行近三年各季度营收数据，数值用万元"></textarea>
          <div style="display:flex;gap:6px;margin-top:6px">
            <button class="secondary tbl-ai-do" data-p="${pi}" data-tbl="${s.id}">AI 生成数据</button>
          </div>
          <div class="tbl-ai-out" data-p="${pi}" data-tbl="${s.id}"></div>
        </div>
      </div>
    </div>`;
  };

  const pageImages = (page: PageState): Record<string, string> => {
    const out: Record<string, string> = {};
    if (!page.imageDataUrl) return out;
    for (const s of page.template.shapes) if (s.role === 'ai_image') out[s.id] = page.imageDataUrl;
    return out;
  };

  const previewHtml = (): string => pages.map((p, i) => `
    <div class="deck-preview-page">
      <div class="hint" style="margin:0 0 4px">第 ${i + 1} 页：${escapeHtml(p.templateName)}</div>
      ${renderPreview(p.template, pageImages(p), p.texts, p.vars, p.tableData)}
    </div>`).join('');

  let previewCollapsed = true;

  const refreshPreview = (): void => {
    const inner = container.querySelector('#wb-preview-inner') as HTMLElement | null;
    if (inner) inner.innerHTML = previewHtml();
  };

  const refreshOutputModeForPageSlot = (root: HTMLElement, page: PageState, slot: string): void => {
    const mode = modeForPageSlot(page, slot);
    root.querySelectorAll('.om-chip[data-om]').forEach((chipRaw) => {
      const chip = chipRaw as HTMLElement;
      if (chip.getAttribute('data-slot') !== slot) return;
      const key = chip.getAttribute('data-om') as 'plain' | 'bullets' | 'condense';
      chip.classList.toggle('active', !!mode[key]);
    });
    root.querySelectorAll('.om-limit').forEach((limitRaw) => {
      const limit = limitRaw as HTMLSelectElement;
      if (limit.getAttribute('data-slot') === slot) limit.value = String(mode.maxChars);
    });
  };

  const cropPageImage = async (page: PageState): Promise<void> => {
    if (!page.imageDataUrl) return;
    const slot = page.template.shapes.find((s) => s.role === 'ai_image');
    if (!slot?.bounds?.width || !slot.bounds.height) return;
    const crop = await openImageCropEditor({
      imageDataUrl: page.imageDataUrl,
      frameRatio: slot.bounds.width / slot.bounds.height,
      frameSizeLabel: '模板图片位：' + Math.round(slot.bounds.width * 96) + ' × ' + Math.round(slot.bounds.height * 96) + ' px'
    });
    if (!crop.canceled && crop.dataUrl) {
      page.imageDataUrl = crop.dataUrl;
      render();
    }
  };

  const render = (): void => {
    container.innerHTML = `
      <h1 class="page-title">生成向导</h1>
      <div class="card wb-tpl-card">
        <label>套版</label>
        <div style="display:flex;gap:8px;align-items:center">
          <select id="wb-template" style="flex:1" aria-label="当前套版">
            <option selected>${escapeHtml(deck.deck.name)}</option>
          </select>
          <button class="secondary" id="dw-back" title="返回套版">返回</button>
        </div>
      </div>
      <div class="card wb-write-bar">
        <button class="primary" id="write-ppt" style="width:100%">生成并插入 PPT</button>
        <p class="hint" id="deck-write-result"></p>
      </div>
      <div class="card">
        <details class="wb-adv gp-adv"><summary>全局提示词${infoTip('作为各段 AI 生成的统一背景；不填时按各字段单独要求生成')}</summary>
          <input id="global-theme" value="${escapeAttr(globalTheme)}" placeholder="例如：输电线路异物清除机器人设计" style="margin-top:6px" />
        </details>
      </div>
      <div class="card">${backgroundPanelHtml()}</div>
      <div id="dw-pages">
        ${pages.map((p, pi) => {
          const textSlots = sortTextSlots(p.template.shapes.filter((s) => s.role === 'ai_text' || s.role === 'manual_var'));
          const imageShapes = p.template.shapes.filter((s) => s.role === 'ai_image');
          const tableShapes = p.template.shapes.filter((s) => s.role === 'table' && s.table);
          return `<div class="card deck-page-fill-card accent-${(pi % 4) + 1}" data-p="${pi}">
            <h3 style="margin:0 0 10px">第 ${pi + 1} 页：${escapeHtml(p.templateName)}${p.templateVersion ? '（' + escapeHtml(p.templateVersion) + '）' : ''}</h3>
            ${textSlots.length ? `<div class="wb-sec"><div class="wb-sec-title"><span>文字</span></div>${textSlots.map((s, i) => textSlotHtml(p, pi, s, i)).join('')}</div>` : ''}
            ${imageShapes.length ? `<div class="wb-sec"><div class="wb-sec-title"><span>图片${infoTip('把本地图片拖到这里直接上传，或点「本地」选择')}</span></div>${imageShapes.map((s, i) => imgSlotHtml(p, pi, s, i)).join('')}</div>` : ''}
            ${tableShapes.length ? `<div class="wb-sec"><div class="wb-sec-title">表格</div>${tableShapes.map((s, ti) => tableSlotHtml(p, pi, s, ti)).join('')}</div>` : ''}
          </div>`;
        }).join('')}
      </div>
      <div class="card wb-preview-card deck-preview-card">
        <div class="wb-preview-head">
          <b>实时预览${infoTip('输入文字 / 换图时自动刷新；点击预览可放大查看。')}</b>
          <span style="display:flex;gap:6px;align-items:center">
            <button class="secondary" id="wb-preview-toggle" title="收起 / 展开预览">${previewCollapsed ? '展开 ▴' : '收起 ▾'}</button>
            <button class="secondary" id="wb-preview-zoom" title="放大预览">放大</button>
          </span>
        </div>
        <div id="wb-preview-body" class="deck-preview-body" style="display:${previewCollapsed ? 'none' : 'block'}">
          <div id="wb-preview-inner" class="deck-preview-inner">${previewHtml()}</div>
          <div id="wb-quality"></div>
        </div>
      </div>`;

    bindEvents();
  };

  const bindEvents = (): void => {
    container.querySelector('#dw-back')!.addEventListener('click', () => { location.hash = '#decks'; });
    const themeInput = container.querySelector('#global-theme') as HTMLInputElement | null;
    themeInput?.addEventListener('input', () => { globalTheme = themeInput.value; });
    const toggleBtn = container.querySelector('#wb-preview-toggle') as HTMLButtonElement | null;
    const previewBody = container.querySelector('#wb-preview-body') as HTMLElement | null;
    toggleBtn?.addEventListener('click', () => {
      previewCollapsed = !previewCollapsed;
      if (previewBody) previewBody.style.display = previewCollapsed ? 'none' : 'block';
      toggleBtn.textContent = previewCollapsed ? '展开 ▴' : '收起 ▾';
    });
    const zoomPreview = () => {
      showPreviewModal('<div style="display:flex;flex-direction:column;gap:12px">' + previewHtml() + '</div>');
    };
    container.querySelector('#wb-preview-zoom')?.addEventListener('click', zoomPreview);
    container.querySelector('#wb-preview-inner')?.addEventListener('click', zoomPreview);

    pages.forEach((page, pi) => {
      const root = container.querySelector(`.deck-page-fill-card[data-p="${pi}"]`) as HTMLElement;
      root.querySelectorAll('textarea.slot-prompt').forEach((ta) => {
        ta.addEventListener('input', () => { page.prompts[(ta as HTMLElement).getAttribute('data-id') || ''] = (ta as HTMLTextAreaElement).value; });
      });
      root.querySelectorAll('.gen-ta.wb-gen').forEach((ta) => {
        ta.addEventListener('input', () => {
          const id = (ta as HTMLElement).getAttribute('data-gen') || '';
          const shape = page.template.shapes.find((s) => s.id === id);
          if (shape?.role === 'manual_var') page.vars[id] = (ta as HTMLTextAreaElement).value;
          else page.texts[id] = (ta as HTMLTextAreaElement).value;
          refreshPreview();
        });
      });
      root.querySelectorAll('.gen-in.wb-gen').forEach((inp) => {
        inp.addEventListener('input', () => {
          page.texts[(inp as HTMLElement).getAttribute('data-gen') || ''] = (inp as HTMLInputElement).value;
          refreshPreview();
        });
      });
      root.querySelectorAll('.seq-chip').forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = (btn as HTMLElement).getAttribute('data-id') || '';
          const n = (btn as HTMLElement).getAttribute('data-seq') || '';
          page.texts[id] = n;
          const inp = root.querySelector(`input.seq-input[data-id="${id}"]`) as HTMLInputElement | null;
          if (inp) inp.value = n;
          root.querySelectorAll(`button.seq-chip[data-id="${id}"]`).forEach((b) => b.classList.toggle('active', b === btn));
          refreshPreview();
        });
      });
      root.querySelectorAll('.seq-input').forEach((inpRaw) => {
        const inp = inpRaw as HTMLInputElement;
        inp.addEventListener('input', () => {
          const v = inp.value.replace(/\D/g, '');
          if (inp.value !== v) inp.value = v;
          page.texts[inp.getAttribute('data-id') || ''] = v;
          root.querySelectorAll(`button.seq-chip[data-id="${inp.getAttribute('data-id') || ''}"]`).forEach((b) => b.classList.remove('active'));
          refreshPreview();
        });
      });
      root.querySelectorAll('.om-chip[data-om]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const slot = (btn as HTMLElement).getAttribute('data-slot') || '';
          const mode = modeForPageSlot(page, slot);
          const key = (btn as HTMLElement).getAttribute('data-om') as 'plain' | 'bullets' | 'condense';
          if (key === 'plain') {
            mode.plain = true;
            mode.bullets = false;
          } else if (key === 'bullets') {
            mode.bullets = true;
            mode.plain = false;
          } else {
            mode.condense = !mode.condense;
          }
          mode.touched = true;
          refreshOutputModeForPageSlot(root, page, slot);
        });
      });
      root.querySelectorAll('.om-limit').forEach((sel) => {
        sel.addEventListener('change', () => {
          const slot = (sel as HTMLElement).getAttribute('data-slot') || '';
          const mode = modeForPageSlot(page, slot);
          mode.maxChars = Number((sel as HTMLSelectElement).value) || 0;
          mode.touched = true;
          refreshOutputModeForPageSlot(root, page, slot);
        });
      });
      root.querySelectorAll('.slot-gen').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', async () => {
          const shapeId = btn.getAttribute('data-id') || '';
          const shape = page.template.shapes.find((s) => s.id === shapeId);
          if (!shape) return;
          const status = root.querySelector(`.slot-status[data-id="${shapeId}"]`) as HTMLElement | null;
          const cfg = await Api.getConfig().catch(() => null);
          if (!cfg || !cfg.text?.apiKey) { showToast('未配置文本 AI，请先前往配置。', 2500); return; }
          btn.disabled = true;
          const oldText = btn.textContent || '';
          btn.textContent = shape.semanticRole === 'formula' ? '转换中…' : '生成中…';
          if (status) status.textContent = '';
          try {
            if (shape.semanticRole === 'formula') {
              const raw = page.texts[shapeId] || '';
              const { text } = await Api.generateText(
                '你是数学公式转换助手。把用户给出的公式文本或描述转换为完整、规范、可直接使用的专业型数学公式，输出 LaTeX 代码。只输出 LaTeX 本体，不要解释。',
                raw ? '请把下面的公式转换为完整专业型 LaTeX：\n' + raw : '请生成一个完整的专业型数学公式（LaTeX）',
                shapeConstraints(shape)
              );
              page.formulas[shapeId] = text.trim().replace(/^```(latex)?/i, '').replace(/```$/, '').trim();
              if (status) status.textContent = '已生成完整公式（生成 PPT 时转为专业型）';
            } else {
              const mode = modeForPageSlot(page, shapeId);
              const { instruction, clean } = resolveSlotMode(mode, shape.semanticRole);
              const per = (page.prompts[shapeId] || '').trim();
              const userPrompt = per || `请围绕主题「${globalTheme || page.imagePrompt || deck.deck.name}」生成`;
              const { text } = await Api.generateText(
                '你是 PPT 文案助手，输出简洁专业的中文文案，不要多余解释。' + (instruction ? '\n' + instruction : ''),
                `${userPrompt}\n段落要求：${shape.prompt || '内容简洁，符合演示文稿风格'}`,
                shapeConstraints(shape),
                clean
              );
              page.texts[shapeId] = text.trim();
              const ta = root.querySelector(`textarea.gen-ta[data-gen="${shapeId}"]`) as HTMLTextAreaElement | null;
              if (ta) ta.value = page.texts[shapeId];
              if (status) status.textContent = '已生成';
              refreshPreview();
            }
          } catch (e) {
            if (status) status.textContent = '生成失败：' + (e as Error).message;
          } finally {
            btn.disabled = false;
            btn.textContent = oldText;
          }
        });
      });
      root.querySelectorAll('.img-pick').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', () => {
          const fileInput = root.querySelector(`input.img-file[data-id="${btn.getAttribute('data-id') || ''}"]`) as HTMLInputElement | null;
          fileInput?.click();
        });
      });
      root.querySelectorAll('.img-file').forEach((inpRaw) => {
        const inp = inpRaw as HTMLInputElement;
        inp.addEventListener('change', async () => {
          const file = inp.files?.[0];
          inp.value = '';
          if (!file) return;
          page.imageDataUrl = await readLocalImage(file);
          await cropPageImage(page);
          render();
        });
      });
      root.querySelectorAll('.img-dropzone').forEach((zoneRaw) => {
        const zone = zoneRaw as HTMLElement;
        zone.addEventListener('click', () => {
          const fileInput = root.querySelector(`input.img-file[data-id="${zone.getAttribute('data-id') || ''}"]`) as HTMLInputElement | null;
          fileInput?.click();
        });
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.style.borderColor = '#1f3864'; });
        zone.addEventListener('dragleave', () => { zone.style.borderColor = '#ccc'; });
        zone.addEventListener('drop', async (e) => {
          e.preventDefault();
          zone.style.borderColor = '#ccc';
          const file = e.dataTransfer?.files?.[0];
          if (!file || !/^image\//.test(file.type)) { showToast('请拖入图片文件', 1800); return; }
          page.imageDataUrl = await readLocalImage(file);
          await cropPageImage(page);
          render();
        });
      });
      root.querySelectorAll('.img-crop').forEach((btn) => {
        btn.addEventListener('click', () => { void cropPageImage(page); });
      });
      root.querySelectorAll('.img-clear').forEach((btn) => {
        btn.addEventListener('click', () => {
          page.imageDataUrl = '';
          page.selected = '';
          render();
        });
      });
      root.querySelectorAll('.tbl-cell').forEach((cellRaw) => {
        const cell = cellRaw as HTMLTextAreaElement;
        cell.addEventListener('input', () => {
          const tbl = cell.getAttribute('data-tbl') || '';
          const r = Number(cell.getAttribute('data-r'));
          const c = Number(cell.getAttribute('data-c'));
          if (!page.tableData[tbl] || !page.tableData[tbl][r]) return;
          page.tableData[tbl][r][c] = cell.value;
          refreshPreview();
        });
      });
      root.querySelectorAll('.tbl-tab').forEach((tabRaw) => {
        const tab = tabRaw as HTMLButtonElement;
        tab.addEventListener('click', () => {
          const tbl = tab.getAttribute('data-tbl') || '';
          const mode = tab.getAttribute('data-mode') || 'manual';
          root.querySelectorAll(`.tbl-tab[data-tbl="${tbl}"]`).forEach((btn) => btn.classList.toggle('active', btn === tab));
          root.querySelectorAll(`.tbl-panel[data-tbl="${tbl}"]`).forEach((panel) => {
            (panel as HTMLElement).style.display = panel.getAttribute('data-mode') === mode ? 'block' : 'none';
          });
        });
      });
      root.querySelectorAll('.tbl-paste-do').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', () => {
          const tbl = btn.getAttribute('data-tbl') || '';
          const ta = root.querySelector(`textarea.tbl-paste-ta[data-tbl="${tbl}"]`) as HTMLTextAreaElement | null;
          const parsed = parseTableText(ta?.value || '');
          if (!parsed.length) { showToast('请先粘贴表格内容', 1800); return; }
          page.tableData[tbl] = parsed;
          render();
        });
      });
      root.querySelectorAll('.tbl-paste-clear').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', () => {
          const tbl = btn.getAttribute('data-tbl') || '';
          if (!page.tableData[tbl]) return;
          page.tableData[tbl] = page.tableData[tbl].map((row) => row.map(() => ''));
          render();
        });
      });
      root.querySelectorAll('.tbl-ai-do').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', async () => {
          const tbl = btn.getAttribute('data-tbl') || '';
          const out = root.querySelector(`.tbl-ai-out[data-tbl="${tbl}"]`) as HTMLElement | null;
          const req = (root.querySelector(`textarea.tbl-ai-req[data-tbl="${tbl}"]`) as HTMLTextAreaElement | null)?.value.trim() || '';
          const shape = page.template.shapes.find((s) => s.id === tbl);
          const current = page.tableData[tbl] || (shape ? defaultTableData(shape) : [[]]);
          const cfg = await Api.getConfig().catch(() => null);
          if (!cfg || !cfg.text?.apiKey) { showToast('未配置文本 AI，请先前往配置。', 2500); return; }
          btn.disabled = true;
          const oldText = btn.textContent || '';
          btn.textContent = '生成中…';
          if (out) out.innerHTML = '';
          try {
            const { text } = await Api.generateText(
              '你是 PPT 表格数据助手。按用户要求生成表格数据，只输出 TSV 或 CSV 纯文本，不要解释，不要 Markdown。',
              `当前表格规模约 ${current.length} 行 × ${Math.max(1, (current[0] || []).length)} 列。\n${req || '请补全一组适合演示文稿的表格数据。'}`,
              undefined,
              { plain: true }
            );
            const parsed = parseTableText(text);
            if (!parsed.length) throw new Error('未得到可导入的表格数据');
            page.tableData[tbl] = parsed;
            render();
          } catch (e) {
            if (out) out.innerHTML = '<p class="error">生成失败：' + escapeHtml((e as Error).message) + '</p>';
          } finally {
            btn.disabled = false;
            btn.textContent = oldText;
          }
        });
      });
      root.querySelectorAll('.img-search').forEach((btnRaw) => {
        const btn = btnRaw as HTMLButtonElement;
        btn.addEventListener('click', async () => {
          const qInp = root.querySelector(`input.img-q[data-id="${btn.getAttribute('data-id') || ''}"]`) as HTMLInputElement | null;
          const q = (qInp?.value || '').trim();
          if (!q) { showToast('请输入图片描述', 1500); return; }
          page.imagePrompt = q;
          page.imgState.q = q;
          const box = root.querySelector(`.img-results[data-id="${btn.getAttribute('data-id') || ''}"]`) as HTMLElement | null;
          if (box) box.innerHTML = '<p class="hint">搜索中…</p>';
          try {
            const res = await Api.searchImages(q, 12, 1, await getDefaultImageProvider());
            page.imgState = { q, images: res.images, error: res.error ? res.error.message : '' };
            for (const cs of page.template.shapes) {
              if (cs.role === 'ai_text' && cs.semanticRole === 'caption' && !(page.texts[cs.id] || '').trim()) page.texts[cs.id] = q;
            }
            render();
            const newRoot = container.querySelector(`.deck-page-fill-card[data-p="${pi}"]`) as HTMLElement;
            const newBox = newRoot.querySelector(`.img-results[data-id="${btn.getAttribute('data-id') || ''}"]`) as HTMLElement | null;
            if (!newBox) return;
            newBox.innerHTML = page.imgState.images.map((img) => `<img class="thumb deck-img-thumb" src="${escapeAttr(img.thumbnailUrl)}" data-url="${escapeAttr(img.imageUrl)}" style="cursor:pointer;width:72px;height:48px;object-fit:cover;border-radius:4px" />`).join('')
              + (page.imgState.error ? '<p class="error">' + escapeHtml(page.imgState.error) + '</p>' : '');
            newBox.querySelectorAll('.deck-img-thumb').forEach((el) => {
              el.addEventListener('click', async () => {
                const url = (el as HTMLElement).getAttribute('data-url') || '';
                newBox.innerHTML = '<p class="hint">下载中…</p>';
                try {
                  const { taskId } = await Api.downloadImage(url, 'baidu_page');
                  let st: any = {};
                  for (let i = 0; i < 60 && !st.done; i++) { st = await Api.getDownloadStatus(taskId); if (!st.done) await sleep(200); }
                  if (st.error) throw new Error(st.error);
                  page.imageDataUrl = st.dataUrl; page.selected = url;
                  const slot = page.template.shapes.find((s) => s.role === 'ai_image');
                  if (slot?.bounds?.width && slot?.bounds?.height) {
                    const crop = await openImageCropEditor({
                      imageDataUrl: st.dataUrl,
                      frameRatio: slot.bounds.width / slot.bounds.height,
                      frameSizeLabel: '模板图片位：' + Math.round(slot.bounds.width * 96) + ' × ' + Math.round(slot.bounds.height * 96) + ' px'
                    });
                    if (!crop.canceled && crop.dataUrl) page.imageDataUrl = crop.dataUrl;
                  }
                  render();
                } catch (e) {
                  newBox.innerHTML = '<p class="error">下载失败：' + escapeHtml((e as Error).message) + '</p>';
                }
              });
            });
          } catch (e) {
            if (box) box.innerHTML = '<p class="error">搜索失败：' + escapeHtml((e as Error).message) + '</p>';
          }
        });
      });
    });

    container.querySelector('#write-ppt')!.addEventListener('click', async () => {
      const buildBtn = container.querySelector('#write-ppt') as HTMLButtonElement;
      const resultEl = container.querySelector('#deck-write-result') as HTMLElement;
      const missing = pages.some((p) => p.template.shapes.some((s) => s.role === 'ai_image') && !p.imageDataUrl);
      if (missing) {
        const choice = await showModal({ title: '有页面未选图片', message: '部分页面包含图片位但尚未选择图片，生成后该页图片位会留空。仍要生成吗？', buttons: [{ id: 'cancel', label: '取消' }, { id: 'go', label: '仍然生成', kind: 'primary' }] });
        if (choice !== 'go') return;
      }
      buildBtn.disabled = true;
      buildBtn.textContent = '正在生成整份…';
      resultEl.textContent = '';
      try {
        const deckTexts = pages.map((pg) => {
          const t = { ...pg.texts };
          for (const fs of pg.template.shapes) {
            if (fs.role === 'ai_text' && fs.semanticRole === 'formula') {
              const raw = pg.texts[fs.id] || '';
              if (!raw.trim()) continue;
              t[fs.id] = (pg.formulas[fs.id] || '').trim() || localTextToLatex(raw);
            }
          }
          return t;
        });
        const res = await Api.buildDeck(pages.map((p, dpi) => ({
          templateId: p.templateId,
          templateFolder: p.templateFolder || undefined,
          templateVersion: p.templateVersion || undefined,
          texts: deckTexts[dpi],
          variables: p.vars,
          imageDataUrl: p.imageDataUrl,
          tableData: p.tableData
        })));
        const failed = (res.pageResults || []).filter((r) => !r.ok);
        if (res.ok && res.base64) {
          await insertSlideBase64(res.base64);
          showToast(`已写入整份 PPT（${res.pageCount} 页）`, 2500);
          if (failed.length) resultEl.textContent = failed.length + ' 页生成失败（已跳过）：' + failed.map((f) => '第' + (f.index + 1) + '页 ' + (f.error || '')).join('；');
          setTimeout(() => { location.hash = '#library'; }, 1200);
        } else {
          resultEl.textContent = '生成失败：' + ((res as { error?: string }).error || failed.map((f) => (f.index + 1) + ':' + f.error).join('；') || '未知错误');
          buildBtn.disabled = false;
          buildBtn.textContent = '生成并插入 PPT';
        }
      } catch (e) {
        resultEl.textContent = '生成失败：' + (e as Error).message;
        buildBtn.disabled = false;
        buildBtn.textContent = '生成并插入 PPT';
      }
    });
  };

  render();
}
