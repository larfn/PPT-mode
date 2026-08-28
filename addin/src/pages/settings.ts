import { Api, ImageSourceDef } from '../api.js';
import { infoTip } from '../lib/tooltip.js';
import { escapeHtml, escapeAttr } from '../lib/html.js';
import { getLanguage, normalizeLanguage, setLanguage, translateDom } from '../lib/i18n.js';

type SrcBundle = { builtins: ImageSourceDef[]; custom: ImageSourceDef[]; presets: ImageSourceDef[] };

type ImageSourceTestSample = { title?: string; imageUrl?: string; thumbnailUrl?: string } | null | undefined;

export function previewImageUrlOf(sample: ImageSourceTestSample): string {
  return (sample?.thumbnailUrl || sample?.imageUrl || '').trim();
}

export async function renderSettings(container: HTMLElement): Promise<void> {
  let cfg: any = {};
  try { cfg = await Api.getConfig(); } catch { /* use empty config when backend is down */ }
  let srcs: SrcBundle = { builtins: [], custom: [], presets: [] };
  try { srcs = await Api.getImageSources(); } catch { /* 后端未就绪：图源管理不可用 */ }
  const highlightDurationSec = (Math.max(0, Math.min(500, Number(cfg.highlight?.durationMs ?? 500))) / 1000).toFixed(1);
  const uiLanguage = getLanguage();

  container.innerHTML = `
    <h1 class="page-title">AI 配置</h1>
    <div class="card">
      <h3>查图服务${infoTip('默认供应商在下方直接切换；自定义图源点「＋ 添加自定义源」在弹窗中添加/管理，支持粘贴 JSON、.json 文件导入与导出，详见项目根目录「自定义图源导入说明.txt」。')}</h3>
      <label>默认供应商</label>
      <select id="img-provider"></select>
      <label>每次显示图片数量${infoTip('搜索图片时每页展示的数量（6-14 张）。')}</label>
      <select id="img-pagesize">
        ${Array.from({ length: 9 }, (_, i) => i + 6).map((n) => `<option value="${n}">${n} 张</option>`).join('')}
      </select>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px">
        <button class="secondary" id="src-add">＋ 添加自定义源</button>
        <button class="secondary" id="src-import">导入 JSON</button>
        <button class="secondary" id="src-file">从文件导入</button>
        <button class="secondary" id="src-export">导出 JSON</button>
      </div>
      <div id="src-msg" class="hint" style="margin-top:6px"></div>
    </div>
    <div class="modal-overlay" id="src-modal" style="display:none">
      <div class="modal-box">
        <div class="modal-head">
          <h3 id="src-modal-title" style="margin:0">添加自定义图源</h3>
          <button class="ghost" id="src-modal-close" title="关闭" style="font-size:18px;line-height:1;padding:2px 8px">×</button>
        </div>
        <div id="src-presets" class="hint-sm" style="margin-bottom:10px"></div>
        <div id="src-custom-list" style="display:flex;flex-direction:column;gap:6px;margin-bottom:10px"></div>
        <div id="src-import-box" style="display:none;margin-bottom:10px">
        <textarea id="src-import-text" rows="5" placeholder='粘贴一个 JSON 对象或数组 [ {…}, {…} ]，字段：id / name / endpoint / headers / cookies / resultsPath / fields / key / enabled（详见根目录说明文档）'></textarea>
        <div style="display:flex;gap:8px;margin-top:6px">
          <button class="primary" id="src-import-do">导入</button>
          <button class="secondary" id="src-import-cancel">取消</button>
        </div>
      </div>
      <div id="src-form" style="display:none;border-top:1px solid var(--border);margin-top:8px;padding-top:10px">
        <input id="sf-id" type="hidden" />
        <label style="display:block">名称<b style="color:var(--danger)">*</b><input id="sf-name" placeholder="如：我的 Pixabay" /></label>
        <label style="display:block;margin-top:10px">接口地址模板<b style="color:var(--danger)">*</b>${infoTip('支持占位符：{query} 关键词、{count} 每页数量、{page} 页码、{start} 偏移量 (page-1)*count、{key} API Key')}</label>
        <input id="sf-endpoint" placeholder="https://api.example.com/search?q={query}&page={page}&n={count}" />
        <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px">
          <label style="flex:1;min-width:200px">结果数组路径<input id="sf-results" placeholder="如 hits / data.items / query.pages，留空=整个 JSON 是数组" /></label>
          <label style="flex:1;min-width:200px">API Key<input id="sf-key" type="password" placeholder="可留空（免 Key 源）" /></label>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px">
          <label style="flex:1;min-width:200px">原图字段<b style="color:var(--danger)">*</b><input id="sf-image-field" placeholder="如 url / image / src.large2x" /></label>
          <label style="flex:1;min-width:200px">缩略图字段<input id="sf-thumb-field" placeholder="如 thumb / thumbnail，可留空" /></label>
        </div>
        <details style="margin-top:10px">
          <summary class="hint-sm" style="cursor:pointer">高级设置（ID / Headers / Cookies / 更多字段）</summary>
          <label style="display:block;margin-top:8px">ID<input id="sf-id-advanced" placeholder="留空自动生成；字母/数字/下划线/中划线" /></label>
          <label style="display:block;margin-top:8px">请求头 Headers（JSON，值里可用 {key}）</label>
          <textarea id="sf-headers" rows="2" placeholder='{ "Authorization": "Client-ID {key}" }'></textarea>
          <label style="display:block;margin-top:8px">Cookies（JSON 对象）</label>
          <textarea id="sf-cookies" rows="2" placeholder='{ "SNUID": "xxxx" }'></textarea>
          <label style="display:block;margin-top:8px">更多字段 Fields（JSON）${infoTip('可选键：width/height/title/sourceUrl/author/license/mimeType/source；原图和缩略图请优先用上方输入框')}</label>
          <textarea id="sf-fields" rows="3" placeholder='{ "width": "w", "height": "h", "title": "title", "sourceUrl": "page" }'></textarea>
        </details>
        <div style="display:flex;gap:14px;margin-top:10px;align-items:center;flex-wrap:wrap">
          <label class="sem-check" style="flex-direction:row;align-items:center;gap:6px;font-weight:normal"><input type="checkbox" id="sf-enabled" checked /> 启用</label>
          <label class="sem-check" style="flex-direction:row;align-items:center;gap:6px;font-weight:normal"><input type="checkbox" id="sf-allow" /> 允许内网/本机地址${infoTip('接口指向局域网或本机图库时勾选（默认禁止内网地址，防止误伤）')}</label>
        </div>
        <div style="display:flex;gap:8px;margin-top:12px">
          <button class="primary" id="sf-save">保存图源</button>
          <button class="secondary" id="sf-test">测试连接</button>
          <button class="secondary" id="sf-cancel">取消</button>
        </div>
        <div id="sf-result" class="hint" style="margin-top:8px"></div>
      </div>
      </div>
    </div>
    <div class="card">
      <h3>文本生成服务${infoTip('Key 只保存在本机，不会发送到插件代码之外。')}</h3>
      <label style="display:block;margin-top:12px">接口地址</label>
      <input id="text-base" value="${escapeAttr(cfg.text?.baseUrl || 'https://api.deepseek.com')}" />
      <label style="display:block;margin-top:12px">API Key${infoTip('留空或保持掩码不变即不修改，输入新 Key 会覆盖旧 Key。')}</label>
      <input id="text-key" type="password" value="${escapeAttr(cfg.text?.apiKey || '')}" placeholder="输入新 Key 可覆盖已保存的 Key" />
      ${cfg.text?.apiKey
        ? '<p class="hint">已保存（此处为掩码）。</p>'
        : '<p class="hint">未配置文本 AI，生成文字时需先在此填写 Key。</p>'}
      <label style="display:block;margin-top:12px">模型名</label>
      <input id="text-model" value="${escapeAttr(cfg.text?.model || 'deepseek-chat')}" />
    </div>
    <div class="card">
      <h3>AI 模板分析${infoTip('保存模板页始终使用内置规则分析并推荐元素角色（免费、不联网）。勾选后，会额外调用上方配置的文本生成服务（AI 增强）优化推荐；不勾选则只使用内置规则，不调用 AI 服务。')}</h3>
      <label class="sem-check" style="flex-direction:row;align-items:center;gap:6px;font-weight:normal">
        <input type="checkbox" id="analyze-enabled" ${cfg.analyze?.enabled ? 'checked' : ''} /> 使用 AI 分析模板服务
      </label>
    </div>
    <div class="card">
      <h3>下载图片库${infoTip('搜图时选中的图片会自动保存到本机文件夹，方便复用。')}</h3>
      <p id="dl-dir" class="dl-dir"></p>
      <button class="secondary" id="open-dl-dir">打开文件夹</button>
    </div>
    <div class="card">
      <h3>界面显示</h3>
      <label>语言</label>
      <select id="ui-language">
        <option value="zh">中文</option>
        <option value="en">English</option>
      </select>
      <label>界面字体大小${infoTip('即时生效，保存后记住设置。')}</label>
      <select id="ui-fontsize">
        <option value="12">小（12）</option>
        <option value="13">较小（13）</option>
        <option value="14">标准（14）</option>
        <option value="15">较大（15）</option>
        <option value="16">大（16）</option>
        <option value="18">特大（18）</option>
      </select>
    </div>
    <div class="card">
      <h3>模板编辑辅助${infoTip('保存模板时点击「定位」，会在 PPT 页面上用该颜色框出对应元素（激光笔效果，自动消失）；时长指框停留显示的时间。')}</h3>
      <label>元素定位框颜色</label>
      <div style="display:flex;align-items:center;gap:8px">
        <input id="hl-color" type="color" value="${escapeAttr(cfg.highlight?.color || '#FF0000')}" />
        <input id="hl-color-text" value="${escapeAttr(cfg.highlight?.color || '#FF0000')}" style="width:90px" />
      </div>
      <label>定位框显示时长（秒）</label>
      <div class="sem-char-range">
        <div class="sem-range-head"><span></span><span class="sem-range-val" id="hl-duration-val">${highlightDurationSec}s</span></div>
        <div class="sem-range-control">
          <input id="hl-duration" type="range" min="0" max="0.5" step="0.1" class="sem-range single-range" value="${highlightDurationSec}" />
        </div>
      </div>
    </div>
    <div class="card">
      <h3>关于 / 系统诊断${infoTip('安装后若功能异常先看这里：前后端版本不一致 = 运行的是旧 exe，请重新运行 安装.bat 部署最新版本并完全退出重开 PowerPoint。')}</h3>
      <div id="diag-info" class="diag-info">加载中…</div>
      <button class="secondary" id="diag-refresh">重新检测</button>
    </div>
    <button class="primary" id="save-cfg">保存配置</button>
    <input type="file" id="src-file-input" accept=".json,application/json" style="display:none" />`;

  // —— 图源管理：状态与渲染 ——
  const msgEl = () => container.querySelector('#src-msg') as HTMLElement;
  const setMsg = (text: string, isErr = false) => {
    const el = msgEl();
    el.textContent = text;
    el.style.color = isErr ? 'var(--danger)' : '';
  };
  const sfResult = () => container.querySelector('#sf-result') as HTMLElement;

  let defaultProvider = cfg.image?.provider || 'baidu_page';

  function updateProviderOptions() {
    const sel = container.querySelector('#img-provider') as HTMLSelectElement;
    const opts: { id: string; name: string; tag: string }[] = [
      ...srcs.builtins.map((b) => ({ id: b.id, name: b.name, tag: '内置' })),
      ...srcs.custom.filter((c) => c.enabled !== false).map((c) => ({ id: c.id, name: c.name, tag: '' }))
    ];
    sel.innerHTML = opts.map((o) => `<option value="${escapeAttr(o.id)}"${o.id === defaultProvider ? ' selected' : ''}>${escapeHtml(o.name)}${o.tag ? '（' + o.tag + '）' : ''}</option>`).join('')
      || '<option value="baidu_page">百度图片（内置）</option>';
  }

  function renderCustomList() {
    const el = container.querySelector('#src-custom-list') as HTMLElement;
    if (!srcs.custom.length) { el.innerHTML = ''; return; }
    const rows: string[] = ['<div class="hint-sm" style="color:var(--text-2)"><b>已添加的自定义源：</b></div>'];
    for (const c of srcs.custom) {
      rows.push(`<div class="src-row" data-id="${escapeAttr(c.id)}">
        <span style="font-weight:600">${escapeHtml(c.name)}</span>
        <span class="hint-sm">${escapeHtml(c.id)}</span>
        ${c.preset ? '<span class="hint-sm" style="color:var(--text-3)">预置</span>' : ''}
        <span style="flex:1"></span>
        ${c.enabled !== false
          ? '<span class="hint-sm" style="color:var(--success)">已启用</span>'
          : '<span class="hint-sm" style="color:var(--danger)">已禁用</span>'}
        <button class="secondary" data-act="toggle" data-id="${escapeAttr(c.id)}">${c.enabled !== false ? '禁用' : '启用'}</button>
        <button class="secondary" data-act="edit" data-id="${escapeAttr(c.id)}">编辑</button>
        <button class="secondary" data-act="del" data-id="${escapeAttr(c.id)}">删除</button>
      </div>`);
    }
    el.innerHTML = rows.join('');
  }

  function renderPresets() {
    const el = container.querySelector('#src-presets') as HTMLElement;
    const added = new Set(srcs.custom.map((c) => c.id));
    el.innerHTML = '<b style="color:var(--text-2)">预置模板：</b>'
      + srcs.presets.map((p) => {
        const isAdded = added.has(p.id);
        return `<span style="display:inline-flex;align-items:center;gap:6px;margin:2px 8px 2px 0;flex-wrap:wrap">
          <span>${escapeHtml(p.name)}</span>
          ${isAdded
            ? '<span class="hint-sm" style="color:var(--text-3)">已添加</span>'
            : `<button class="secondary" data-act="preset" data-id="${escapeAttr(p.id)}" style="padding:1px 8px;font-size:12px">添加</button>`}
        </span>`;
      }).join('')
      || '<span class="hint-sm">（无预置模板）</span>';
  }

  function refresh() {
    void Api.getImageSources().then((r) => {
      srcs = r;
      renderCustomList();
      renderPresets();
      updateProviderOptions();
    }).catch((e) => setMsg('加载图源列表失败：' + ((e as Error)?.message || '后端不可用'), true));
  }

  // —— 弹窗与表单 ——
  const formEl = () => container.querySelector('#src-form') as HTMLElement;
  const modalEl = () => container.querySelector('#src-modal') as HTMLElement;
  const setModalTitle = (t: string) => { (container.querySelector('#src-modal-title') as HTMLElement).textContent = t; };
  let editingId: string | null = null;

  function openModal() {
    modalEl().style.display = 'flex';
    renderPresets();
    renderCustomList();
  }

  function closeModal() {
    modalEl().style.display = 'none';
    formEl().style.display = 'none';
    container.querySelector('#src-import-box')!.setAttribute('style', 'display:none;margin-bottom:10px');
    editingId = null;
  }

  function showImportBox() {
    setModalTitle('导入图源 JSON');
    formEl().style.display = 'none';
    container.querySelector('#src-import-box')!.setAttribute('style', 'display:block;margin-bottom:10px');
    openModal();
  }

  function showForm(def: ImageSourceDef | null, fromPreset = false) {
    editingId = def ? def.id : null;
    setModalTitle(def && !fromPreset ? '编辑图源' : '添加自定义图源');
    container.querySelector('#src-import-box')!.setAttribute('style', 'display:none;margin-bottom:10px');
    formEl().style.display = 'block';
    openModal();
    const setId = container.querySelector('#sf-id') as HTMLInputElement;
    const setAdvancedId = container.querySelector('#sf-id-advanced') as HTMLInputElement;
    const fields = def?.fields || {};
    (container.querySelector('#sf-name') as HTMLInputElement).value = def?.name || '';
    setId.value = def?.id || '';
    setAdvancedId.value = def?.id || '';
    setAdvancedId.disabled = !!def && !fromPreset; // 编辑已有源时 ID 不可改
    (container.querySelector('#sf-endpoint') as HTMLInputElement).value = def?.endpoint || '';
    (container.querySelector('#sf-results') as HTMLInputElement).value = def?.resultsPath || '';
    (container.querySelector('#sf-key') as HTMLInputElement).value = def?.key && def.key !== '****' ? def.key : '';
    (container.querySelector('#sf-headers') as HTMLTextAreaElement).value = def?.headers && Object.keys(def.headers).length ? JSON.stringify(def.headers, null, 2) : '';
    (container.querySelector('#sf-cookies') as HTMLTextAreaElement).value = def?.cookies && Object.keys(def.cookies).length ? JSON.stringify(def.cookies, null, 2) : '';
    (container.querySelector('#sf-image-field') as HTMLInputElement).value = fields.imageUrl || '';
    (container.querySelector('#sf-thumb-field') as HTMLInputElement).value = fields.thumbnailUrl || '';
    const extraFields = { ...fields } as Record<string, string>;
    delete extraFields.imageUrl;
    delete extraFields.thumbnailUrl;
    (container.querySelector('#sf-fields') as HTMLTextAreaElement).value = Object.keys(extraFields).length ? JSON.stringify(extraFields, null, 2) : '';
    (container.querySelector('#sf-enabled') as HTMLInputElement).checked = def ? def.enabled !== false : true;
    (container.querySelector('#sf-allow') as HTMLInputElement).checked = def?.allowPrivate === true;
    sfResult().textContent = def?.note ? ('提示：' + def.note) : '';
    if (fromPreset) sfResult().textContent = '已填入预置模板，确认/修改后保存。' + (def?.note ? ' ' + def.note : '');
  }

  function hideForm() {
    closeModal();
  }

  function readJson(sel: string, fallback: any): any {
    const v = (container.querySelector(sel) as HTMLTextAreaElement).value.trim();
    if (!v) return fallback;
    return JSON.parse(v); // 失败向上抛，由调用方捕获提示
  }

  function sourceIdFromName(name: string): string {
    const base = name.trim()
      .replace(/[\s.]+/g, '_')
      .replace(/[^a-zA-Z0-9_-]/g, '')
      .replace(/^[-_]+|[-_]+$/g, '')
      .slice(0, 48);
    return (base || 'custom_image_source') + '_' + Date.now().toString(36).slice(-6);
  }

  function gatherDef(): ImageSourceDef {
    const name = (container.querySelector('#sf-name') as HTMLInputElement).value.trim();
    const hiddenId = (container.querySelector('#sf-id') as HTMLInputElement).value.trim();
    const advancedId = (container.querySelector('#sf-id-advanced') as HTMLInputElement).value.trim();
    const fields = readJson('#sf-fields', {});
    const imageField = (container.querySelector('#sf-image-field') as HTMLInputElement).value.trim();
    const thumbField = (container.querySelector('#sf-thumb-field') as HTMLInputElement).value.trim();
    if (imageField) fields.imageUrl = imageField;
    if (thumbField) fields.thumbnailUrl = thumbField;
    const def: ImageSourceDef = {
      id: hiddenId || advancedId || sourceIdFromName(name),
      name,
      endpoint: (container.querySelector('#sf-endpoint') as HTMLInputElement).value.trim(),
      headers: readJson('#sf-headers', {}),
      cookies: readJson('#sf-cookies', {}),
      fields,
      key: (container.querySelector('#sf-key') as HTMLInputElement).value,
      allowPrivate: (container.querySelector('#sf-allow') as HTMLInputElement).checked,
      enabled: (container.querySelector('#sf-enabled') as HTMLInputElement).checked
    };
    const rp = (container.querySelector('#sf-results') as HTMLInputElement).value.trim();
    if (rp) def.resultsPath = rp;
    return def;
  }

  async function saveForm() {
    let def: ImageSourceDef;
    try {
      def = gatherDef();
    } catch (e) {
      setMsg('JSON 解析失败：' + (e as Error).message, true);
      return;
    }
    if (!def.name || !def.endpoint || !def.fields?.imageUrl) { setMsg('名称、接口地址、原图字段为必填项', true); return; }
    try {
      await Api.saveImageSource(def);
      setMsg('已保存图源：' + def.name);
      hideForm();
      refresh();
    } catch (e) {
      setMsg('保存失败：' + (e as Error).message, true);
    }
  }

  async function testForm() {
    let def: ImageSourceDef;
    try {
      def = gatherDef();
    } catch (e) {
      setMsg('JSON 解析失败：' + (e as Error).message, true);
      return;
    }
    if (!def.endpoint || !def.fields?.imageUrl) { setMsg('请先填写接口地址和原图字段', true); return; }
    const btn = container.querySelector('#sf-test') as HTMLButtonElement;
    btn.disabled = true; btn.textContent = '测试中…';
    sfResult().textContent = '';
    try {
      const r = await Api.testImageSource(def);
      if (r.ok) {
        const sample = r.sample;
        const previewUrl = previewImageUrlOf(sample);
        sfResult().innerHTML = `<span style="color:var(--success)">✓ 测试通过，搜到 ${r.resultsCount} 张</span>`
          + (previewUrl
            ? ` <img src="${escapeAttr(previewUrl)}" style="width:64px;height:44px;object-fit:cover;border-radius:4px;vertical-align:middle" />`
            : '')
          + (sample?.title ? ` <span class="hint-sm">${escapeHtml(sample.title)}</span>` : '');
      } else {
        sfResult().innerHTML = '<span style="color:var(--danger)">✗ 测试失败：' + escapeHtml(r.error?.message || '未知错误') + '</span>';
      }
    } catch (e) {
      sfResult().innerHTML = '<span style="color:var(--danger)">✗ 测试失败：' + escapeHtml((e as Error).message) + '</span>';
    } finally {
      btn.disabled = false; btn.textContent = '测试连接';
    }
  }

  async function doImportText() {
    const ta = container.querySelector('#src-import-text') as HTMLTextAreaElement;
    const raw = ta.value.trim();
    if (!raw) { setMsg('请先粘贴 JSON', true); return; }
    let parsed: any;
    try { parsed = JSON.parse(raw); } catch (e) { setMsg('JSON 解析失败：' + (e as Error).message, true); return; }
    const arr = Array.isArray(parsed) ? parsed : [parsed];
    try {
      const r = await Api.importImageSources(arr);
      const parts = ['成功添加 ' + r.added.length + ' 个'];
      if (r.skipped.length) parts.push('跳过 ' + r.skipped.length + ' 个：' + r.skipped.map((s) => (s.id || '?') + '（' + s.reason + '）').join('；'));
      setMsg(parts.join('；'));
      ta.value = '';
      closeModal();
      refresh();
    } catch (e) {
      setMsg('导入失败：' + (e as Error).message, true);
    }
  }

  function exportJson() {
    const arr = srcs.custom.map((c) => {
      const o: any = { ...c };
      if (o.key === '****') delete o.key; // 不导出 Key
      delete o.builtin; delete o.preset; delete o.note;
      return o;
    });
    const blob = new Blob([JSON.stringify(arr, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'ppt-ai-addin-图源自定义.json';
    a.click();
    URL.revokeObjectURL(a.href);
    setMsg('已导出 ' + arr.length + ' 个自定义源');
  }

  // —— 事件绑定（图源管理） ——
  updateProviderOptions();

  container.querySelector('#src-add')!.addEventListener('click', () => showForm(null));
  container.querySelector('#src-import')!.addEventListener('click', showImportBox);
  container.querySelector('#src-modal-close')!.addEventListener('click', closeModal);
  modalEl().addEventListener('click', (ev) => { if (ev.target === modalEl()) closeModal(); });
  container.querySelector('#src-import-do')!.addEventListener('click', () => void doImportText());
  container.querySelector('#src-import-cancel')!.addEventListener('click', closeModal);
  container.querySelector('#src-export')!.addEventListener('click', exportJson);
  const fileInput = container.querySelector('#src-file-input') as HTMLInputElement;
  container.querySelector('#src-file')!.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', () => {
    const f = fileInput.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(String(reader.result || ''));
        const arr = Array.isArray(parsed) ? parsed : [parsed];
        const r = await Api.importImageSources(arr);
        setMsg('文件导入完成：添加 ' + r.added.length + ' 个，跳过 ' + r.skipped.length + ' 个');
        refresh();
      } catch (e) {
        setMsg('文件导入失败：' + (e as Error).message, true);
      }
    };
    reader.readAsText(f);
    fileInput.value = '';
  });
  container.querySelector('#sf-save')!.addEventListener('click', () => void saveForm());
  container.querySelector('#sf-test')!.addEventListener('click', () => void testForm());
  container.querySelector('#sf-cancel')!.addEventListener('click', hideForm);

  container.querySelector('#src-custom-list')!.addEventListener('click', async (ev) => {
    const btn = (ev.target as HTMLElement).closest('button[data-act]') as HTMLButtonElement | null;
    if (!btn) return;
    const act = btn.getAttribute('data-act') || '';
    const id = btn.getAttribute('data-id') || '';
    if (act === 'toggle') {
      const c = srcs.custom.find((x) => x.id === id);
      if (!c) return;
      try {
        await Api.saveImageSource({ ...c, enabled: c.enabled === false });
        refresh();
      } catch (e) { setMsg('操作失败：' + (e as Error).message, true); }
    } else if (act === 'edit') {
      const c = srcs.custom.find((x) => x.id === id);
      if (c) showForm(c);
    } else if (act === 'del') {
      if (!confirm('确认删除自定义源「' + id + '」？')) return;
      try {
        await Api.deleteImageSource(id);
        setMsg('已删除：' + id);
        refresh();
      } catch (e) { setMsg('删除失败：' + (e as Error).message, true); }
    }
  });

  container.querySelector('#src-presets')!.addEventListener('click', (ev) => {
    const btn = (ev.target as HTMLElement).closest('button[data-act="preset"]') as HTMLButtonElement | null;
    if (!btn) return;
    const p = srcs.presets.find((x) => x.id === (btn.getAttribute('data-id') || ''));
    if (p) showForm(p, true);
  });

  // —— 其余配置项（与旧版一致） ——
  const pageSizeSel = container.querySelector('#img-pagesize') as HTMLSelectElement;
  const curPageSize = Number(cfg.image?.pageSize);
  pageSizeSel.value = (curPageSize >= 6 && curPageSize <= 14) ? String(curPageSize) : '9';

  const dlDirEl = container.querySelector('#dl-dir') as HTMLElement;
  if (cfg.downloadDir) dlDirEl.textContent = cfg.downloadDir;
  container.querySelector('#open-dl-dir')!.addEventListener('click', async () => {
    try {
      const { dir } = await Api.openDownloadDir();
      alert(`已打开文件夹：${dir}`);
    } catch (e) {
      alert(`打开文件夹失败：${(e as Error).message}`);
    }
  });

  const fontSel = container.querySelector('#ui-fontsize') as HTMLSelectElement;
  fontSel.value = String(cfg.ui?.fontSize || 14);
  fontSel.addEventListener('change', () => {
    document.documentElement.style.setProperty('--ui-font', fontSel.value + 'px');
  });
  const languageSel = container.querySelector('#ui-language') as HTMLSelectElement;
  languageSel.value = uiLanguage;
  languageSel.addEventListener('change', () => {
    setLanguage(normalizeLanguage(languageSel.value));
    translateDom(document.body);
  });

  const hlColor = container.querySelector('#hl-color') as HTMLInputElement;
  const hlColorText = container.querySelector('#hl-color-text') as HTMLInputElement;
  hlColor.addEventListener('input', () => { hlColorText.value = hlColor.value; });
  hlColorText.addEventListener('input', () => {
    const v = hlColorText.value.trim();
    if (/^#[0-9a-fA-F]{6}$/.test(v)) hlColor.value = v;
  });
  const hlDuration = container.querySelector('#hl-duration') as HTMLInputElement;
  const hlDurationVal = container.querySelector('#hl-duration-val') as HTMLElement;
  const syncHighlightDuration = () => {
    hlDurationVal.textContent = Number(hlDuration.value).toFixed(1) + 's';
  };
  hlDuration.addEventListener('input', syncHighlightDuration);
  syncHighlightDuration();

  const diagInfo = container.querySelector('#diag-info') as HTMLElement;
  const diagRows: [string, string, 'ok' | 'bad' | ''][] = [];
  const pushRow = (k: string, v: string, cls: 'ok' | 'bad' | '' = '') => diagRows.push([k, v, cls]);
  const renderDiag = () => {
    diagInfo.innerHTML = diagRows.map(([k, v, cls]) =>
      `<div class="diag-row"><span class="diag-k">${k}</span><span class="diag-v ${cls}">${v}</span></div>`).join('');
  };
  async function loadDiagnostics() {
    diagRows.length = 0;
    try {
      const v = await Api.getAppVersion();
      pushRow('版本状态', v.appVersion === __APP_VERSION__
        ? '一致（' + __APP_VERSION__ + '）'
        : '前端 ' + __APP_VERSION__ + ' / 后端 ' + v.appVersion, v.appVersion === __APP_VERSION__ ? 'ok' : 'bad');
      pushRow('后端连接', '正常', 'ok');
      pushRow('MCP', v.mcp.running ? '运行中' : '未运行（需要时启动）', v.mcp.running ? 'ok' : '');
    } catch (e) {
      pushRow('后端连接', '失败：' + (e as Error).message, 'bad');
    }
    renderDiag();
  }
  container.querySelector('#diag-refresh')!.addEventListener('click', loadDiagnostics);
  void loadDiagnostics();

  container.querySelector('#save-cfg')!.addEventListener('click', async () => {
    const body: any = {
      image: {
        provider: (container.querySelector('#img-provider') as HTMLSelectElement).value,
        pageSize: Number((container.querySelector('#img-pagesize') as HTMLSelectElement).value) || 9
      },
      text: {
        baseUrl: (container.querySelector('#text-base') as HTMLInputElement).value.trim(),
        model: (container.querySelector('#text-model') as HTMLInputElement).value.trim()
      },
      highlight: {
        color: (container.querySelector('#hl-color-text') as HTMLInputElement).value.trim() || '#FF0000',
        durationMs: Math.max(0, Math.min(500, Math.round(Number((container.querySelector('#hl-duration') as HTMLInputElement).value || 0) * 1000)))
      },
      ui: {
        fontSize: Number((container.querySelector('#ui-fontsize') as HTMLSelectElement).value) || 14,
        language: (container.querySelector('#ui-language') as HTMLSelectElement).value
      },
      analyze: {
        enabled: (container.querySelector('#analyze-enabled') as HTMLInputElement).checked === true
      }
    };
    // 已保存的 Key 以掩码形式显示在输入框中；留空或保持掩码不变时不覆盖原 Key
    const textKey = (container.querySelector('#text-key') as HTMLInputElement).value.trim();
    if (textKey && textKey !== (cfg.text?.apiKey || '')) { body.text.apiKey = textKey; }
    try {
      await Api.saveConfig(body);
      alert('配置已保存');
    } catch (e) {
      alert(`保存失败：${(e as Error).message}`);
    }
  });
}
