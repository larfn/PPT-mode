import { Api, TemplateBackground, TemplateDoc, TemplateShape } from '../api.js';
import { readCurrentSlide, ShapeInfo, SlideBackgroundInfo } from '../office/readSlide.js';
import { highlightShapeOnSlide, getHighlightConfig } from '../office/highlight.js';
import { showProgress, showToast, showModal } from '../ui.js';
import { escapeHtml, escapeAttr } from '../lib/html.js';
import { SEMANTIC_ROLES, defaultSemanticRole, semanticRoleLabel, sortShapesByPosition } from '../lib/semantic.js';
import { analyzeShapesByRules, HIGH_CONFIDENCE, ShapeRecommendation } from '../lib/analyze.js';
import { infoTip } from '../lib/tooltip.js';
import { markInputError, clearInputError, markErrorText, clearErrorText } from '../lib/formHint.js';
import { getDocumentZipBytes } from '../lib/zip.js';
import { compressImageDataUrl } from '../lib/image.js';
import { defaultRoleForShape } from '../lib/roleDefaults.js';
import { setRouteLeaveGuard } from '../lib/navigation.js';
import { insertTemporarySlideBase64, deleteSlideById } from '../office/writeSlide.js';

const ROLE_LABELS: Record<string, string> = {
  ai_image: 'AI 图片位', ai_text: 'AI 文本位', manual_var: '手动变量位', fixed: '固定元素',
  table: '表格位（生成时填数据）'
};

const SOURCE_LABELS: Record<string, string> = { slide: '页面', layout: '版式', master: '母版' };

const SHAPE_OPTIONS: [string, string][] = [
  ['rect', '矩形'], ['roundRect', '圆角矩形'], ['ellipse', '椭圆'], ['line', '直线'],
  ['rightArrow', '右箭头'], ['leftArrow', '左箭头'], ['upArrow', '上箭头'], ['downArrow', '下箭头'],
  ['leftRightArrow', '左右箭头'], ['diamond', '菱形'], ['triangle', '三角形'], ['chevron', '燕尾形'],
  ['pentagon', '五边形'], ['hexagon', '六边形'], ['trapezoid', '梯形'], ['parallelogram', '平行四边形'],
  ['donut', '环形'], ['heart', '心形'], ['star5', '五角星'], ['cloud', '云朵']
];

const BG_LABELS: Record<string, string> = {
  none: '无（白色）', solid: '纯色', picture: '图片', gradient: '渐变', pattern: '图案', unsupported: '无法读取'
};

const TEMP_PREVIEW_NOTICE = '临时预览页：仅用于定位查看，离开或保存时会自动删除。';

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

// 当前选中页的页码（1 起），用于精确定位背景所在页。
async function getSelectedSlideIndex(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load('index');
    await context.sync();
    return slide.index;
  });
}


// 用 getSelectedDataAsync(SlideRange) 拿选中页的稳定 slideId（与文件 presentation.xml 的 sldId 对应，
// 不受页码偏移/隐藏页影响——此前用 slide.index 解析有 1 页偏差）。
function getSelectedSlideId(): Promise<number | null> {
  return new Promise((resolve) => {
    try {
      const office = (window as unknown as { Office?: any }).Office;
      office?.context?.document?.getSelectedDataAsync(office?.CoercionType?.SlideRange, (result: any) => {
        try {
          const slides = result?.value?.slides;
          if (result?.status === office.AsyncResultStatus.Failed || !slides || !slides.length) resolve(null);
          else resolve(Number(slides[0].id) || null);
        } catch { resolve(null); }
      });
    } catch { resolve(null); }
  });
}

// 从 Office.context.document.url 拿本地文档磁盘路径。兼容两种格式：
//  - file:///C:/Users/...（部分 Office 版本返回 URL 格式）
//  - C:\\Users\\...（部分 Office 版本直接返回 Windows 路径——实测当前环境）
// 不需要 COM，权限问题绕开；未保存的文档 url 可能是空。
function getDocumentPathFromOffice(): string | null {
  try {
    const office = (window as unknown as { Office?: any }).Office;
    const url = office?.context?.document?.url;
    if (typeof url !== 'string' || !url) return null;
    if (url.startsWith('file:///')) {
      return decodeURIComponent(url.slice(7).replace(/\//g, '\\'));
    }
    // 裸 Windows 路径（C:\... 或 C:/...）
    if (/^[A-Za-z]:[\\/]/.test(url)) {
      return decodeURIComponent(url.replace(/\//g, '\\'));
    }
    return null;
  } catch { /* ignore */ }
  return null;
}

export async function renderSaveTemplate(container: HTMLElement): Promise<void> {
  container.innerHTML = `
    <div class="page">
      <h1 class="page-title">保存模板</h1>
      <section class="module">
        <h2 class="module-title">模板信息</h2>
        <div class="field-row">
          <label for="tpl-name">模板名称</label>
          <input id="tpl-name" placeholder="例如：季度汇报-封面" />
        </div>
        <div class="field-row">
          <label for="tpl-folder">分类</label>
          <select id="tpl-folder"></select>
          <input id="tpl-folder-new" placeholder="新文件夹名称" style="display:none" />
          <details class="wb-adv" style="margin-top:6px">
            <summary>版本说明（可选）</summary>
            <input id="tpl-note" placeholder="例如：调整了标题配色" style="margin-top:6px" />
          </details>
        </div>
      </section>
      <div class="page-actions sticky-save">
        <button id="save-tpl" class="primary" style="display:none">保存</button>
        <button id="read-slide" class="secondary">读取</button>
      </div>
      <section class="module" id="slide-info-module" style="display:none">
        <h2 class="module-title">页面</h2>
        <div id="slide-info"></div>
      </section>
      <section class="module">
        <h2 class="module-title">元素</h2>
        <div id="shape-list"></div>
      </section>
    </div>`;

  // 分类文件夹：加载已有分类 + 新建入口（返回 Promise，编辑模式需等它完成后再选中原分类）
  const folderSel = container.querySelector('#tpl-folder') as HTMLSelectElement;
  const folderNew = container.querySelector('#tpl-folder-new') as HTMLInputElement;
  const folderReady = (async () => {
    let folders: { name: string; count: number }[] = [];
    try { folders = await Api.listFolders(); } catch { /* 后端未就绪时仅保留默认项 */ }
    folderSel.innerHTML = '<option value="">未分类</option>'
      + folders.map((f) => '<option value="' + escapeHtml(f.name) + '">' + escapeHtml(f.name) + '（' + f.count + '）</option>').join('')
      + '<option value="__new__">＋ 新建文件夹…</option>';
    folderSel.addEventListener('change', () => {
      folderNew.style.display = folderSel.value === '__new__' ? 'block' : 'none';
    });
  })();

  const infoEl = container.querySelector('#slide-info') as HTMLElement;
  const listEl = container.querySelector('#shape-list') as HTMLElement;
  const saveBtn = container.querySelector('#save-tpl') as HTMLButtonElement;
  let shapes: ShapeInfo[] = [];
  let slideSize = { width: 13.33, height: 7.5 };
  let background: SlideBackgroundInfo | null = null; // 仅用于界面展示（类型/来源）
  let bgUseDefault = true; // 勾选「跟随文档」：保存当前文档/母版解析到的背景；取消后使用自选图片背景
  let bgOverride = ''; // 自选背景图（dataURL，未勾选沿用默认时写入模板）
  const buildTemplateBackground = async (): Promise<TemplateBackground | undefined> => {
    if (!bgUseDefault) {
      return { type: 'picture', imageDataUrl: await compressImageDataUrl(bgOverride, 1280, 0.85) };
    }
    if (!background) return undefined;
    if (background.type === 'picture' && background.imageDataUrl) {
      return { type: 'picture', source: background.source || 'slide', imageDataUrl: await compressImageDataUrl(background.imageDataUrl, 1280, 0.85) };
    }
    if (background.type === 'solid' && background.color) {
      return { type: 'solid', source: background.source || 'slide', color: background.color };
    }
    return undefined;
  };
  let unsupported: string[] = [];
  const roles: Record<string, string> = {};
  const shapeTypes: Record<string, string> = {};
  let enrichTask: Promise<void> | null = null; // 后台任务：回读形状精确样式（不阻塞保存）
  let enrichDone = false; // 样式回读是否已成功并入 shapes
  let enrichFailed = false; // 样式回读是否失败（保存后非阻塞提示）
  let highlightColor = '#FF0000'; // 定位框颜色（从配置读取）
  let highlightDuration = 500; // 定位框停留时长（ms，从配置读取）
  const refreshHighlightConfig = async (): Promise<void> => {
    try {
      const c = await getHighlightConfig();
      highlightColor = c.color;
      highlightDuration = c.durationMs;
    } catch { /* 使用默认定位配置 */ }
  };
  let hasLoadedTemplateState = false; // 已读取/载入模板信息，离开保存页前需确认
  let allowLeaveWithoutConfirm = false; // 保存成功等程序化跳转不弹确认
  let tempPreviewSlideId: string | null = null; // 编辑已保存模板时插入的临时预览页
  const cleanupTempPreviewSlide = async (): Promise<void> => {
    const id = tempPreviewSlideId;
    tempPreviewSlideId = null;
    if (!id) return;
    try { await deleteSlideById(id); } catch { /* 临时页可能已被用户手动删除，忽略 */ }
  };
  setRouteLeaveGuard('#save', async () => {
    if (!hasLoadedTemplateState || allowLeaveWithoutConfirm) return true;
    const choice = await showModal({
      title: '确认离开当前界面',
      message: '离开后，本次已读取/载入的模板信息不会保留。确定离开当前界面？',
      buttons: [
        { id: 'leave', label: '确认离开', kind: 'danger' },
        { id: 'cancel', label: '取消', kind: 'secondary' }
      ]
    });
    if (choice === 'leave') {
      await cleanupTempPreviewSlide();
      setRouteLeaveGuard('#save', null);
      return true;
    }
    return false;
  });
  const varLabels: Record<string, string> = {}; // 各元素的提示词/变量名输入值
  // 「固定」勾选：勾选 = 元素原样保留（role=fixed），不参与生成（默认不勾；几何/无文本元素默认勾选）
  const fixedShapes: Record<number, boolean> = {};
  // 图标类小图（icon）：最大边 ≤ 0.6 英寸视为图标 → 与简单几何一样默认固定；大图不默认固定（作图片位参与生成）
  const isIconPicture = (s: ShapeInfo): boolean => {
    if (s.type !== 'picture') return false;
    const w = s.bounds?.width || 0, h = s.bounds?.height || 0;
    return w > 0 && h > 0 && Math.max(w, h) <= 0.6;
  };
  // 模板语义层：各元素的语义角色与生成约束（仅 AI 文本位参与保存）
  const semantics: Record<number, {
    semanticRole?: string; contentType?: string; required?: boolean;
    maxChars?: number; maxLines?: number; minChars?: number; preferredLength?: number;
    generationInstruction?: string;
    translate?: boolean; translateSource?: string;
  }> = {};
  // —— AI 自动模板分析 ——
  const recommendations: Record<number, ShapeRecommendation> = {}; // 当前生效的推荐（规则或 AI）
  const appliedRecs: Record<number, true> = {};    // 已接受的推荐
  const ignoredRecs: Record<number, true> = {};    // 用户忽略的推荐
  const userTouchedRoles: Record<number, true> = {};    // 用户手动改过 role（AI/规则绝不覆盖）
  const userTouchedSemantics: Record<number, true> = {}; // 用户手动改过 semanticRole
  const analysisState = { running: false, done: false, aiUsed: false, aiError: '', autoAppliedCount: 0 };

  const previewLabelFor = (s: TemplateShape, index: number): string => {
    if (s.semanticRole) return semanticRoleLabel(s.semanticRole);
    if (s.contentType) return s.contentType;
    const name = (s.name || '').trim();
    if (name && name.length <= 8 && !/[{}：:]/.test(name)) return name;
    return `文字位 ${index + 1}`;
  };

  const previewImageLabelFor = (s: TemplateShape, index: number): string => {
    if (s.semanticRole) return semanticRoleLabel(s.semanticRole);
    if (s.contentType) return s.contentType;
    const name = (s.name || '').trim();
    if (name && name.length <= 8 && !/[{}：:]/.test(name)) return name;
    return `图片位 ${index + 1}`;
  };

  const buildTemporaryPreviewPayload = (template: TemplateDoc): {
    template: TemplateDoc;
    texts: Record<string, string>;
    vars: Record<string, string>;
  } => {
    const textsOut: Record<string, string> = {};
    const varsOut: Record<string, string> = {};
    const slideH = template.slideSize?.height || 7.5;
    const slideW = template.slideSize?.width || 13.33;
    const shapesOut = (template.shapes || []).map((s, i) => {
      if (s.role === 'ai_text') {
        textsOut[s.id] = previewLabelFor(s, i);
        return s;
      }
      if (s.role === 'manual_var') {
        varsOut[s.id] = previewLabelFor(s, i);
        return s;
      }
      if (s.role === 'ai_image') {
        return {
          ...s,
          type: 'text',
          role: 'fixed',
          content: previewImageLabelFor(s, i),
          textStyle: {
            font: 'Microsoft YaHei',
            eaFont: '微软雅黑',
            size: 15,
            color: '#666666',
            align: 'center',
            valign: 'middle',
            wordWrap: true
          },
          fill: { type: 'Solid', color: '#F3F5F8' }
        } as TemplateShape;
      }
      return s;
    });
    shapesOut.push({
      id: '__temporary_preview_notice',
      type: 'text',
      role: 'fixed',
      source: 'slide',
      bounds: { left: 0.2, top: Math.max(0, slideH - 0.38), width: Math.max(1, slideW - 0.4), height: 0.28 },
      content: TEMP_PREVIEW_NOTICE,
      textStyle: {
        font: 'Microsoft YaHei',
        eaFont: '微软雅黑',
        size: 15,
        color: '#FF0000',
        bold: true,
        align: 'center',
        valign: 'middle',
        wordWrap: true,
        margin: { top: 0, right: 0, bottom: 0, left: 0 }
      }
    });
    return {
      template: { ...template, shapes: shapesOut },
      texts: textsOut,
      vars: varsOut
    };
  };

  const createTemporaryPreviewSlide = async (template: TemplateDoc): Promise<void> => {
    await cleanupTempPreviewSlide();
    const payload = buildTemporaryPreviewPayload(template);
    const built = await Api.buildSlideBase64({
      template: payload.template,
      texts: payload.texts,
      vars: payload.vars
    });
    tempPreviewSlideId = await insertTemporarySlideBase64(built.base64);
    showToast('已插入临时预览页，可用「定位」查看元素位置', 2500);
  };

// 从 PPT 当前选中页删除指定页面元素（可 Ctrl+Z 撤回）。
// 仅支持 source='slide' 的页面元素；版式/母版元素属于全局，无法单独删除。
async function deleteShapeOnSlide(s: ShapeInfo): Promise<void> {
  if (s.source !== 'slide') {
    throw new Error('版式/母版元素无法单独删除（会影响所有使用该版式/母版的页面）');
  }
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const target = slide.shapes.getItem(s.id);
    target.delete();
    await context.sync();
  });
}

  // 手动变量位角色已移除：旧模板/AI 推荐里的 manual_var 一律按 ai_text（文字位）处理
  const normalizeRole = (r: string): string => (r === 'manual_var' ? 'ai_text' : r);

  // 应用一条推荐（用户手动点「接受」= 明确意图，不检查 userTouched；自动应用会检查）
  const applyRecommendation = (idx: number): void => {
    const rec = recommendations[idx];
    const s = shapes[idx];
    if (!rec || !s) return;
    // 表格位角色只对表格形状生效（防 AI 对普通形状误推荐 table）
    if (rec.recommendedRole === 'table' && s.type !== 'table') return;
    const role = normalizeRole(rec.recommendedRole);
    roles[idx] = role;
    fixedShapes[idx] = role === 'fixed'; // 推荐「固定元素」= 勾选固定
    if (role === 'ai_text') {
      const sem = (semantics[idx] = semantics[idx] || {});
      if (rec.recommendedSemanticRole) sem.semanticRole = rec.recommendedSemanticRole;
      else if (!sem.semanticRole) sem.semanticRole = defaultSemanticRole(s.textStyle?.size);
      if (rec.suggestedConstraints) {
        for (const [k, v] of Object.entries(rec.suggestedConstraints)) {
          if (k === 'maxChars' || k === 'minChars') (sem as Record<string, unknown>)[k] = v;
        }
      }
      if (rec.suggestedPrompt) varLabels[idx] = rec.suggestedPrompt;
    }
    appliedRecs[idx] = true;
    renderList();
  };

  // 忽略推荐：标记后不再显示
  const ignoreRecommendation = (idx: number): void => {
    ignoredRecs[idx] = true;
    renderList();
  };

  // 一键接受全部高置信度建议：跳过用户已手动设置的 role/semanticRole 与已忽略项
  const acceptAllHighConfidence = (): void => {
    let n = 0;
    for (const [k, rec] of Object.entries(recommendations)) {
      const idx = Number(k);
      if (rec.confidence < HIGH_CONFIDENCE) continue;
      if (ignoredRecs[idx]) continue;
      if (userTouchedRoles[idx] || userTouchedSemantics[idx]) continue; // 保护用户设置
      if (appliedRecs[idx]) continue;
      applyRecommendation(idx); // applyRecommendation 会 renderList，循环后统一刷新
      n++;
    }
    if (n) showToast(`已应用 ${n} 条高置信度推荐 ✓`, 1800);
    else showToast('没有可自动应用的高置信度建议（已应用或已手动设置）', 2000);
    renderList();
  };

  // 分析主流程：规则分类（必做）→ AI 增强（可选，失败回退规则）。异步执行，不阻塞保存。
  const runAnalysis = async (): Promise<void> => {
    analysisState.running = true;
    analysisState.done = false;
    analysisState.aiError = '';
    // 先让 UI 显示「正在分析」
    renderList();
    // 规则分类（纯本地，快）
    const ruleRecs = analyzeShapesByRules(shapes, slideSize);
    for (const r of ruleRecs) recommendations[r.idx] = r;
    // AI 增强（可选）：配置页勾选「使用 AI 分析模板服务」且配置了文本 AI 才调用；
    // 不勾选 = 只用内置规则；任何失败都保持规则结果
    try {
      const cfg = await Api.getConfig();
      if (cfg?.analyze?.enabled === true && cfg?.text?.apiKey) {
        const aiRes = await Api.analyzeShapes(shapes.map((s, i) => ({
          shapeId: String(i),
          name: s.name || s.type || '',
          type: s.type || '',
          fontSize: s.textStyle?.size || 0,
          bold: s.textStyle?.bold === true,
          phType: s.phType || '',
          align: s.textStyle?.align || '',
          color: s.textStyle?.color || '',
          lines: (s.text || '').split(/\r?\n/).length,
          text: (s.text || '').replace(/\s+/g, ' ').slice(0, 120),
          left: s.bounds?.left || 0, top: s.bounds?.top || 0,
          width: s.bounds?.width || 0, height: s.bounds?.height || 0,
          source: s.source || ''
        })));
        if (aiRes.ok && Array.isArray(aiRes.recommendations)) {
          for (const ai of aiRes.recommendations) {
            const idx = Number(ai.shapeId);
            if (!Number.isFinite(idx) || idx < 0 || idx >= shapes.length) continue; // shapeId 不存在
            if (appliedRecs[idx] || ignoredRecs[idx] || userTouchedRoles[idx]) continue; // 保护
            recommendations[idx] = {
              ...recommendations[idx],
              idx,
              recommendedRole: ai.recommendedRole,
              recommendedSemanticRole: ai.recommendedSemanticRole,
              confidence: ai.confidence,
              reason: ai.reason || (recommendations[idx]?.reason || ''),
              // 提示词约定：AI 识别仅增强识别效果，绝不改动提示词（提示词只由规则生成）
              suggestedPrompt: recommendations[idx]?.suggestedPrompt,
              suggestedConstraints: (ai.suggestedConstraints as ShapeRecommendation['suggestedConstraints']) || recommendations[idx]?.suggestedConstraints,
              source: 'ai'
            };
          }
          analysisState.aiUsed = true;
        } else {
          analysisState.aiError = (aiRes as { error?: string }).error || 'AI 分析不可用';
        }
      }
    } catch (e) {
      analysisState.aiError = 'AI 分析不可用：' + ((e as Error).message || String(e));
    }
    // —— 自动应用高置信度推荐（P2 优化）：读页后默认预标角色，免去逐个点「接受」——
    // 只自动应用「规则来源且置信度 ≥ HIGH_CONFIDENCE」的推荐（规则免费/本地/可靠）；
    // AI 增强推荐仍保留手动确认（可能不准，用户可点「一键接受」）。用户手动改过/已忽略的绝不覆盖。
    let autoApplied = 0;
    for (const [k, rec] of Object.entries(recommendations)) {
      const idx = Number(k);
      if (rec.source !== 'rule' || rec.confidence < HIGH_CONFIDENCE) continue;
      if (ignoredRecs[idx]) continue;
      if (userTouchedRoles[idx] || userTouchedSemantics[idx]) continue; // 保护用户设置
      if (appliedRecs[idx]) continue;
      const role = normalizeRole(rec.recommendedRole);
      roles[idx] = role;
      fixedShapes[idx] = role === 'fixed';
      if (role === 'ai_text') {
        const sem = (semantics[idx] = semantics[idx] || {});
        if (rec.recommendedSemanticRole) sem.semanticRole = rec.recommendedSemanticRole;
        else if (!sem.semanticRole) sem.semanticRole = defaultSemanticRole(shapes[idx]?.textStyle?.size);
        if (rec.suggestedConstraints) {
          for (const [ck, cv] of Object.entries(rec.suggestedConstraints)) {
            if (ck === 'maxChars' || ck === 'minChars') (sem as Record<string, unknown>)[ck] = cv;
          }
        }
        if (rec.suggestedPrompt) varLabels[idx] = rec.suggestedPrompt;
      }
      appliedRecs[idx] = true;
      autoApplied++;
    }
    analysisState.running = false;
    analysisState.done = true;
    if (autoApplied) {
      // 保存到渲染后提示一次（renderList 之后显示）
      analysisState.autoAppliedCount = autoApplied;
    }
    renderList();
  };

  const renderList = (): void => {
    const counts = { slide: 0, layout: 0, master: 0 };
    shapes.forEach((s) => { if (s.source && counts[s.source] !== undefined) counts[s.source]++; });
    const countsLine = `<p class="hint" style="margin:2px 0 10px">共 ${shapes.length} 个元素：页面 ${counts.slide} · 版式 ${counts.layout} · 母版 ${counts.master}</p>`;
    // —— AI 自动分析状态条（不阻塞保存，可先做其他操作）——
    let analyzeBar = '';
    if (analysisState.running) {
      analyzeBar = `<div class="analyze-bar"><p class="hint" style="margin:0">正在分析模板…（完成后给出角色建议，可先进行其他操作）</p></div>`;
    } else if (analysisState.done) {
      const pending = Object.entries(recommendations).filter(([k, rec]) => {
        const i = Number(k);
        return !appliedRecs[i] && !ignoredRecs[i] && !userTouchedRoles[i] && !userTouchedSemantics[i] && rec.confidence >= HIGH_CONFIDENCE;
      }).length;
      const total = Object.values(recommendations).length;
      const autoNote = analysisState.autoAppliedCount ? '已自动应用 ' + analysisState.autoAppliedCount + ' 条建议' : '';
      const aiNote = analysisState.aiError ? escapeHtml(analysisState.aiError) : '';
      analyzeBar = `<div class="analyze-bar done">
        <div class="analyze-title"><span>AI 分析</span><span class="analyze-count">已识别 ${total} 个元素${analysisState.aiUsed ? '（含 AI 增强）' : ''}</span><span class="flex-spacer"></span>${(autoNote || aiNote) ? `<span class="analyze-auto">${autoNote}${aiNote ? (autoNote ? '；' : '') + aiNote : ''}</span>` : ''}</div>
        ${pending ? `<button class="secondary accept-all">应用全部建议（${pending} 条）</button>` : ''}
      </div>`;
    }
    // 单行推荐（接受/忽略）
    const recBox = (idx: number): string => {
      const rec = recommendations[idx];
      if (!rec) return '';
      const recRole = normalizeRole(rec.recommendedRole); // manual_var → ai_text（手动变量位角色已移除）
      const recLabel = recRole === 'ai_text' ? '文字' : recRole === 'ai_image' ? '图片' : (ROLE_LABELS[recRole] || recRole);
      if (appliedRecs[idx]) return `<span class="rec-line rec-applied"><span class="rec-text">✓ 已应用推荐：${escapeHtml(recLabel)}${rec.recommendedSemanticRole ? '（' + escapeHtml(semanticRoleLabel(rec.recommendedSemanticRole)) + '）' : ''}</span></span>`;
      if (ignoredRecs[idx] || userTouchedRoles[idx] || userTouchedSemantics[idx]) return ''; // 尊重用户
      const roleLabel = recLabel;
      const semLabel = rec.recommendedSemanticRole ? semanticRoleLabel(rec.recommendedSemanticRole) : '';
      return `<div class="rec-line rec-pending">
        <span class="rec-text">建议：${escapeHtml(roleLabel)}${semLabel ? '（' + escapeHtml(semLabel) + '）' : ''}</span>
        <span class="rec-actions">
          <button class="ghost rec-apply" data-idx="${idx}" title="应用">√</button>
          <button class="ghost rec-ignore" data-idx="${idx}" title="忽略">×</button>
        </span>
      </div>`;
    };
    const fixedPics = shapes.filter((s, i) => s.type === 'picture' && (fixedShapes[i] === true || roles[i] === 'fixed' || (fixedShapes[i] === undefined && isIconPicture(s))));
    const imgNote = fixedPics.length
      ? '<p class="hint" style="margin:2px 0 8px">' + fixedPics.length + ' 个固定图片位将随模板保存（图标/装饰图原样保留，生成时无需重新选图）。</p>'
      : '';
    const elemCardHtml = (s: ShapeInfo, idx: number): string => {
      const defaultRole = defaultRoleForShape(s);
      const canDelete = s.source === 'slide';
      const delTitle = canDelete ? '从 PPT 页面删除该元素（Ctrl+Z 可撤回），并从模板中移除' : '版式/母版元素无法单独删除（会影响所有使用该版式/母版的页面）';
      const varValue = varLabels[idx] || '';
      // 固定勾选：用户手动勾过用勾选值；未操作过时默认跟随类型（几何/无文本元素默认固定）
      const isFixed = fixedShapes[idx] === true || (fixedShapes[idx] === undefined && (defaultRole === 'fixed' || isIconPicture(s)));
      const role = isFixed ? 'fixed' : (roles[idx] || defaultRole);
      const sem = semantics[idx] || {};
      const defaultSemRole = role === 'ai_text' ? defaultSemanticRole(s.textStyle?.size) : '';
      const semRole = sem.semanticRole || (role === 'ai_text' ? (defaultSemRole || 'other') : '');
      return `
      <div class="elem-card">
        <div class="elem-head">
          <span class="elem-idx">${String(idx + 1).padStart(2, '0')}</span>
          <span class="elem-name" title="${escapeAttr(s.name || s.type)}">${escapeHtml(s.name || (s.type === 'table' ? '表格' : s.type))}</span>
          <span class="elem-type">${s.source === 'layout' ? '版式' : s.source === 'master' ? '母版' : ''}</span>
          <span class="flex-spacer"></span>
          <button class="locate" data-idx="${idx}" title="在 PPT 中框出该元素">定位</button>
          <button class="del" data-idx="${idx}" ${canDelete ? '' : 'disabled'} title="${delTitle}">${canDelete ? '删除' : '—'}</button>
        </div>
        <div class="elem-row">
          ${s.type === 'table'
            ? `<select class="role-select" data-idx="${idx}" ${isFixed ? 'disabled' : ''}>
                ${[['fixed', ROLE_LABELS.fixed], ['table', ROLE_LABELS.table]].map(([v, l]) => `<option value="${v}" ${v === role ? 'selected' : ''}>${l}</option>`).join('')}
              </select>`
            : `<select class="role-select" data-idx="${idx}" ${isFixed ? 'disabled' : ''} title="文字=AI/手动生成文本；图片=选择图片填充">
                ${[['ai_text', '文字'], ['ai_image', '图片']].map(([v, l]) => `<option value="${v}" ${v === role ? 'selected' : ''}>${l}</option>`).join('')}
              </select>`}
          <label class="fixed-cb-wrap" title="勾选 = 元素原样保留（固定），不参与生成；不勾选 = 作为文字/图片位参与生成">
            <input type="checkbox" class="fixed-cb" data-idx="${idx}" ${isFixed ? 'checked' : ''} /> 固定
          </label>
          ${recBox(idx)}
          ${(s.type === 'rectangle' || s.type === 'line') ? `
          <select class="shape-type" data-idx="${idx}">
            ${SHAPE_OPTIONS.map(([k, l]) => `<option value="${k}" ${k === (shapeTypes[idx] || s.shapeType || (s.type === 'line' ? 'line' : 'rect')) ? 'selected' : ''}>${l}</option>`).join('')}
          </select>` : ''}
        </div>
        ${s.hasText ? `<input data-idx="${idx}" class="var-label" value="${escapeHtml(varValue)}" placeholder="提示词（文字位，可留空）" style="${(!isFixed && role === 'ai_text') ? '' : 'display:none'}" />` : ''}
        ${s.hasText ? `
        <div class="sem-block" data-idx="${idx}" style="${(!isFixed && role === 'ai_text') ? '' : 'display:none'}">
          <select class="sem-role" data-idx="${idx}" title="语义角色：该文本位在页面中的用途，AI 生成时据此把握文案风格与篇幅">
            ${SEMANTIC_ROLES.map((r) => `<option value="${r.value}" ${r.value === semRole ? 'selected' : ''}>${r.label}</option>`).join('')}
          </select>
          <details class="elem-more"><summary>生成设置</summary>
            <div class="more-body sem-grid">
              <div class="sem-wide sem-char-range" data-idx="${idx}">
                <div class="sem-range-head"><span>字符范围</span><span class="sem-range-val" data-idx="${idx}"></span></div>
                <div class="sem-range-control">
                  <input type="range" min="0" max="300" class="sem-range sem-range-min" data-idx="${idx}" data-field="minChars" step="10" value="${sem.minChars ?? 0}" />
                  <input type="range" min="0" max="300" class="sem-range sem-range-max" data-idx="${idx}" data-field="maxChars" step="10" value="${sem.maxChars ?? 300}" />
                </div>
              </div>
              <label class="sem-check" style="align-items:center;gap:6px"><input type="checkbox" class="sem-translate" data-idx="${idx}" ${sem.translate ? 'checked' : ''} /> 自动翻译副标题（英文）</label>
              <label class="sem-wide">翻译原文来源
                <select class="sem-source" data-idx="${idx}">
                  <option value="theme" ${(sem.translateSource || 'theme') === 'theme' ? 'selected' : ''}>全局主题（向导步骤3顶部输入）</option>
                  ${shapes.map((src, j) => {
                    if (j === idx) return '';
                    const sr = roles[j] || (src.type === 'picture' ? 'ai_image' : 'fixed');
                    if (sr !== 'ai_text' && sr !== 'manual_var') return '';
                    const label = sr === 'ai_text'
                      ? '文本位 ' + (j + 1)
                      : '变量「' + escapeHtml(varLabels[j] || src.name || '变量' + (j + 1)) + '」';
                    // value 存目标形状的数组索引（保存时映射为 shpN），避免存 Office 原始 id 导致引用断链
                    const tsVal = String(sem.translateSource || '');
                    const isSel = tsVal === String(j) || tsVal === 'shp' + j;
                    return '<option value="' + j + '"' + (isSel ? ' selected' : '') + '>' + label + '</option>';
                  }).join('')}
                </select>
              </label>
            </div>
          </details>
        </div>` : ''}
      </div>`;
    };
    // —— 固定元素收纳：勾选「固定」的元素移到所有元素下方，自动折叠收纳（展开可正常修改/取消固定）——
    const isFixedNow = (idx: number): boolean => {
      const s0 = shapes[idx];
      return fixedShapes[idx] === true || (fixedShapes[idx] === undefined && (defaultRoleForShape(s0) === 'fixed' || isIconPicture(s0)));
    };
    const activeCards = shapes.map((s, idx) => isFixedNow(idx) ? '' : elemCardHtml(s, idx)).join('');
    const fixedCards = shapes.map((s, idx) => isFixedNow(idx) ? elemCardHtml(s, idx) : '').join('');
    const fixedCount = shapes.reduce((n, s, idx) => n + (isFixedNow(idx) ? 1 : 0), 0);
    const fixedFold = fixedCount
      ? '<details class="fixed-fold"><summary>固定元素（' + fixedCount + '）' + infoTip('勾选「固定」的元素原样保留、不参与生成（几何图案/母版装饰/无文本元素默认固定）；收纳在此折叠下，展开可修改/取消固定') + '</summary><div class="fixed-fold-body">' + fixedCards + '</div></details>'
      : '';
    listEl.innerHTML = imgNote + analyzeBar + countsLine + (shapes.length ? activeCards + fixedFold : '<p class="hint">页面内没有可保存的形状（仅背景）。</p>') + (unsupported.length ? '<p class="hint">表格/图表/SmartArt 等类型暂不随模板保存，建议先转成图片或文字。</p>' : '');

    // 角色下拉：只绑 role-select（此前用 querySelectorAll('select') 会把形状类型下拉的值
    // 误写进 roles，导致已标记的 role 被覆盖为形状名 —— 修复：按类名精确绑定）
    listEl.querySelectorAll('select.role-select').forEach((sel) => {
      sel.addEventListener('change', () => {
        const idx = Number(sel.getAttribute('data-idx'));
        roles[idx] = (sel as HTMLSelectElement).value;
        fixedShapes[idx] = false; // 手动选择了文字/图片 → 取消固定
        const fixedCb = listEl.querySelector(`input.fixed-cb[data-idx="${idx}"]`) as HTMLInputElement | null;
        if (fixedCb) fixedCb.checked = false;
        userTouchedRoles[idx] = true; // 用户手动设置：规则/AI 推荐不再覆盖
        const block = listEl.querySelector(`.sem-block[data-idx="${idx}"]`) as HTMLElement | null;
        if (block) block.style.display = roles[idx] === 'ai_text' ? '' : 'none';
        // 「提示词」输入框：仅文字位（ai_text）需要；图片/固定等隐藏
        const labelInput = listEl.querySelector(`input.var-label[data-idx="${idx}"]`) as HTMLElement | null;
        if (labelInput) labelInput.style.display = roles[idx] === 'ai_text' ? '' : 'none';
        // 设为文字位时给出合理默认语义角色（按字号启发：≥28 标题 / ≥20 副标题 / 其他正文），用户可改或不指定
        if (roles[idx] === 'ai_text') {
          semantics[idx] = semantics[idx] || {};
          if (!semantics[idx].semanticRole) {
            semantics[idx].semanticRole = defaultSemanticRole(shapes[idx]?.textStyle?.size);
            const roleSel = listEl.querySelector(`select.sem-role[data-idx="${idx}"]`) as HTMLSelectElement | null;
            if (roleSel) roleSel.value = semantics[idx].semanticRole!;
          }
        }
      });
    });
    // 「固定」勾选：勾选 = 原样保留（固定），不参与生成；取消 = 恢复默认文字/图片角色
    listEl.querySelectorAll('input.fixed-cb').forEach((cb) => {
      cb.addEventListener('change', () => {
        const idx = Number(cb.getAttribute('data-idx'));
        const checked = (cb as HTMLInputElement).checked;
        fixedShapes[idx] = checked;
        userTouchedRoles[idx] = true; // 用户手动设置：规则/AI 推荐不再覆盖
        const s = shapes[idx];
        const defaultRole = defaultRoleForShape(s);
        const roleSel = listEl.querySelector(`select.role-select[data-idx="${idx}"]`) as HTMLSelectElement | null;
        if (roleSel) {
          roleSel.disabled = checked;
          if (!checked) {
            // 取消固定：恢复默认角色（保留用户之前手动选择的值）
            if (!roles[idx] || roles[idx] === 'fixed') roles[idx] = defaultRole;
            roleSel.value = roles[idx] || defaultRole;
          }
        }
        const block = listEl.querySelector(`.sem-block[data-idx="${idx}"]`) as HTMLElement | null;
        if (block) block.style.display = (!checked && roles[idx] === 'ai_text') ? '' : 'none';
        const labelInput = listEl.querySelector(`input.var-label[data-idx="${idx}"]`) as HTMLElement | null;
        if (labelInput) labelInput.style.display = (!checked && roles[idx] === 'ai_text') ? '' : 'none';
        // 固定勾选状态变化后整体重渲染：勾选的元素移入下方「固定元素」折叠，取消则移回普通区
        renderList();
      });
    });
    listEl.querySelectorAll('select.shape-type').forEach((sel) => {
      const stSel = sel as HTMLSelectElement;
      stSel.addEventListener('change', () => { shapeTypes[stSel.getAttribute('data-idx')!] = stSel.value; });
    });
    listEl.querySelectorAll('input.var-label').forEach((inp) => {
      inp.addEventListener('change', () => { varLabels[inp.getAttribute('data-idx')!] = (inp as HTMLInputElement).value; });
    });
    // —— 模板语义层 ——
    listEl.querySelectorAll('select.sem-role').forEach((sel) => {
      sel.addEventListener('change', () => {
        const idx = Number(sel.getAttribute('data-idx'));
        userTouchedSemantics[idx] = true; // 用户手动设置：不覆盖
        const v = (sel as HTMLSelectElement).value;
        semantics[idx] = semantics[idx] || {};
        if (v) semantics[idx].semanticRole = v; else delete semantics[idx].semanticRole;
        // —— 提示词自动同步：改角色时，若当前提示词为空或是旧角色的默认提示词，则自动换成新角色的默认提示词 ——
        // （主标题 →「围绕主题输入主标题」；副标题 →「围绕主题输入副标题」；正文 →「围绕主题输入正文」；序号/日期/图片诠释/公式/其他 → 无提示词）
        // 用户手动写过的自定义提示词保留不动。
        {
          const oldPrompt = varLabels[idx] || '';
          const KNOWN_DEFAULTS = ['围绕主题输入主标题', '围绕主题输入副标题', '围绕主题输入正文'];
          const isDefaultOrEmpty = !oldPrompt || KNOWN_DEFAULTS.indexOf(oldPrompt) >= 0;
          let newPrompt = '';
          if (v === 'title') newPrompt = '围绕主题输入主标题';
          else if (v === 'subtitle') newPrompt = '围绕主题输入副标题';
          else if (v === 'body') newPrompt = '围绕主题输入正文';
          if (isDefaultOrEmpty) {
            varLabels[idx] = newPrompt;
            const labelInput = listEl.querySelector('input.var-label[data-idx="' + idx + '"]') as HTMLInputElement | null;
            if (labelInput) labelInput.value = newPrompt;
          }
        }
        // 自动识别：语义角色选「副标题」时自动开启「自动翻译副标题」（英文），默认原文 = 本页主标题（无则第一个文本位/变量）
        if (v === 'subtitle' && !semantics[idx].translate) {
          semantics[idx].translate = true;
          // 默认原文 = 本页第一个 title 角色 AI 文本位（无则第一个文本位/变量）；存数组索引，保存时映射为 shpN
          const tIdx = shapes.findIndex((_, j) => j !== idx && roles[j] === 'ai_text' && semantics[j]?.semanticRole === 'title');
          const tIdx2 = tIdx >= 0 ? tIdx : shapes.findIndex((_, j) => j !== idx && (roles[j] === 'ai_text' || roles[j] === 'manual_var'));
          semantics[idx].translateSource = tIdx2 >= 0 ? String(tIdx2) : 'theme';
          const tcb = listEl.querySelector('input.sem-translate[data-idx="' + idx + '"]') as HTMLInputElement | null;
          if (tcb) tcb.checked = true;
          const srcSel = listEl.querySelector('select.sem-source[data-idx="' + idx + '"]') as HTMLSelectElement | null;
          if (srcSel) srcSel.value = semantics[idx].translateSource;
        }
      });
    });
    const syncCharRange = (idx: number, changedField?: string): void => {
      const minEl = listEl.querySelector(`input.sem-range-min[data-idx="${idx}"]`) as HTMLInputElement | null;
      const maxEl = listEl.querySelector(`input.sem-range-max[data-idx="${idx}"]`) as HTMLInputElement | null;
      const valEl = listEl.querySelector(`.sem-range-val[data-idx="${idx}"]`) as HTMLElement | null;
      if (!minEl || !maxEl) return;
      const toCharStep = (value: number): number => Math.max(0, Math.min(300, Math.round(value / 10) * 10));
      let min = toCharStep(Number(minEl.value) || 0);
      let max = toCharStep(Number(maxEl.value) || 300);
      if (max < 300 && min > max) {
        if (changedField === 'minChars') max = min;
        else min = max;
      }
      minEl.value = String(min);
      maxEl.value = String(max);
      const sem = (semantics[idx] = semantics[idx] || {});
      if (min > 0) sem.minChars = min; else delete sem.minChars;
      if (max > 0 && max < 300) sem.maxChars = max; else delete sem.maxChars;
      if (valEl) valEl.textContent = '最少 ' + (min > 0 ? min + ' 字' : '不设') + ' / 最多 ' + (max < 300 ? max + ' 字' : '不设上限');
    };
    listEl.querySelectorAll('input.sem-range').forEach((inp) => {
      const el = inp as HTMLInputElement;
      const idx = Number(el.getAttribute('data-idx'));
      syncCharRange(idx);
      el.addEventListener('input', () => syncCharRange(idx, el.getAttribute('data-field') || undefined));
      el.addEventListener('change', () => syncCharRange(idx, el.getAttribute('data-field') || undefined));
    });
    // 自动翻译副标题：勾选开关 + 选择翻译原文来源（全局主题 / 本页其他 AI 文本位或手动变量位）
    listEl.querySelectorAll('input.sem-translate').forEach((cb) => {
      cb.addEventListener('change', () => {
        const idx = Number(cb.getAttribute('data-idx'));
        const sem = (semantics[idx] = semantics[idx] || {});
        if ((cb as HTMLInputElement).checked) {
          sem.translate = true;
          // 首次开启时给出合理默认来源：本页第一个标题类 AI 文本位 → 第一个其他 AI 文本位 → 全局主题
          if (!sem.translateSource) {
            const tIdx = shapes.findIndex((_, j) => j !== idx && roles[j] === 'ai_text' && semantics[j]?.semanticRole === 'title');
            const tIdx2 = tIdx >= 0 ? tIdx : shapes.findIndex((_, j) => j !== idx && (roles[j] === 'ai_text' || roles[j] === 'manual_var'));
            sem.translateSource = tIdx2 >= 0 ? String(tIdx2) : 'theme';
            const sel = listEl.querySelector('select.sem-source[data-idx="' + idx + '"]') as HTMLSelectElement | null;
            if (sel) sel.value = sem.translateSource;
          }
        } else {
          delete sem.translate;
        }
      });
    });
    listEl.querySelectorAll('select.sem-source').forEach((sel) => {
      sel.addEventListener('change', () => {
        const idx = Number(sel.getAttribute('data-idx'));
        const sem = (semantics[idx] = semantics[idx] || {});
        const v = (sel as HTMLSelectElement).value;
        if (v) sem.translateSource = v; else delete sem.translateSource;
      });
    });
    // —— AI 分析：接受 / 忽略 / 一键接受高置信度 ——
    listEl.querySelectorAll('.rec-apply').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        if (Number.isFinite(idx) && recommendations[idx]) applyRecommendation(idx);
      });
    });
    listEl.querySelectorAll('.rec-ignore').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        if (Number.isFinite(idx)) ignoreRecommendation(idx);
      });
    });
    const acceptAllBtn = container.querySelector('.accept-all') as HTMLButtonElement | null;
    if (acceptAllBtn) acceptAllBtn.addEventListener('click', acceptAllHighConfidence);
    listEl.querySelectorAll('button.locate').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const idx = Number(btn.getAttribute('data-idx'));
        const s = shapes[idx];
        if (!s) return;
        await refreshHighlightConfig();
        void highlightShapeOnSlide(s.bounds, highlightColor, highlightDuration);
      });
    });
    listEl.querySelectorAll('button.del').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        if (!shapes[idx]) return;
        if (btn.classList.contains('confirm')) {
          void performDelete(idx);
        } else {
          btn.classList.add('confirm');
          btn.textContent = '确认删除？';
          window.setTimeout(() => {
            btn.classList.remove('confirm');
            btn.textContent = '删除';
          }, 3000);
        }
      });
    });
  };

  const performDelete = async (idx: number): Promise<void> => {
    const s = shapes[idx];
    if (!s) return;
    try {
      await deleteShapeOnSlide(s);
      shapes.splice(idx, 1);
      // 平移索引：删除位置之后的角色/形状类型/提示词整体前移
      const shift = (rec: Record<string, string>) => {
        for (let i = idx; i < shapes.length; i++) {
          const v = rec[i + 1];
          if (v === undefined) delete rec[i];
          else rec[i] = v;
        }
        delete rec[shapes.length];
      };
      const shiftFixed = (rec: Record<number, boolean>) => {
        for (let i = idx; i < shapes.length; i++) {
          const v = rec[i + 1];
          if (v === undefined) delete rec[i];
          else rec[i] = v;
        }
        delete rec[shapes.length];
      };
      shift(roles); shift(shapeTypes); shift(varLabels); shiftFixed(fixedShapes);
      // 语义配置随索引整体前移（与角色/提示词一致）
      for (let i = idx; i < shapes.length; i++) {
        const v = semantics[i + 1];
        if (v === undefined) delete semantics[i];
        else semantics[i] = v;
      }
      delete semantics[shapes.length];
      // AI 分析相关状态同步前移（推荐/已应用/已忽略/用户手动标记）
      const shiftRec = (rec: Record<number, unknown>) => {
        for (let i = idx; i < shapes.length; i++) {
          const v = rec[i + 1];
          if (v === undefined) delete rec[i];
          else rec[i] = v;
        }
        delete rec[shapes.length];
      };
      shiftRec(recommendations as unknown as Record<number, unknown>);
      shiftRec(appliedRecs as unknown as Record<number, unknown>);
      shiftRec(ignoredRecs as unknown as Record<number, unknown>);
      shiftRec(userTouchedRoles as unknown as Record<number, unknown>);
      shiftRec(userTouchedSemantics as unknown as Record<number, unknown>);
      renderList();
      if (!shapes.length && !background) {
        saveBtn.style.display = 'none';
      }
    } catch (e) {
      alert(`删除失败：${(e as Error).message}`);
    }
  };

  // —— 编辑模式：从模板库点「编辑」进入，载入已保存模板的全部信息（名称/分类/版本说明/元素/角色/语义/背景）——
  const editId = sessionStorage.getItem('editTemplateId');
  const editFolder = sessionStorage.getItem('editTemplateFolder') || '';
  sessionStorage.removeItem('editTemplateId');
  sessionStorage.removeItem('editTemplateFolder');
  if (editId) {
    void (async () => {
      try {
        const loaded = await Api.getTemplate(editId, editFolder);
        const template = loaded.template;
        // 名称 / 版本说明（当前版本的说明回填，保存时可修改）
        (container.querySelector('#tpl-name') as HTMLInputElement).value = template.name || loaded.name || '';
        const noteInput = container.querySelector('#tpl-note') as HTMLInputElement | null;
        if (noteInput && template.changeNote) noteInput.value = template.changeNote;
        // 分类：等文件夹列表加载完再选中；原分类不在列表（后端未就绪等）则动态补一项
        await folderReady;
        const opt = Array.from(folderSel.options).find((o) => o.value === editFolder);
        if (opt) folderSel.value = editFolder;
        else if (editFolder) {
          const o = document.createElement('option');
          o.value = editFolder; o.textContent = editFolder;
          folderSel.insertBefore(o, folderSel.querySelector('option[value="__new__"]'));
          folderSel.value = editFolder;
        }
        // 元素：模板形状还原为界面数据（与保存格式一一对应，保证完全一致）
        slideSize = template.slideSize || slideSize;
        // 先按位置排序（上到下、左到右，页面→版式→母版），再还原为界面数据；
        // 角色/固定/语义等标记必须按「排序后的新索引」对齐（后续渲染/保存都以 shapes 数组索引为准）
        const srcShapes = (template.shapes || []).map((o, i) => ({ o, i }));
        srcShapes.sort((a, b) => {
          const ra = a.o.source === 'layout' ? 1 : a.o.source === 'master' ? 2 : 0;
          const rb = b.o.source === 'layout' ? 1 : b.o.source === 'master' ? 2 : 0;
          if (ra !== rb) return ra - rb;
          const ba = a.o.bounds || { top: 0, left: 0 }, bb = b.o.bounds || { top: 0, left: 0 };
          if (Math.abs((ba.top || 0) - (bb.top || 0)) > 0.001) return (ba.top || 0) - (bb.top || 0);
          return (ba.left || 0) - (bb.left || 0);
        });
        shapes = srcShapes.map(({ o, i }) => {
          const isPic = o.type === 'picture';
          const isTbl = o.type === 'table';
          const isLine = o.type === 'line';
          return {
            id: o.id || 'shp' + i,
            type: isPic ? 'picture' : isTbl ? 'table' : isLine ? 'line' : (o.type || 'text'),
            name: o.name || '',
            source: (o.source === 'layout' || o.source === 'master') ? o.source : 'slide',
            bounds: o.bounds || { left: 0, top: 0, width: 1, height: 1 },
            hasText: !isPic && !isTbl && !isLine,
            text: o.content,
            textStyle: o.textStyle,
            fill: o.fill,
            line: o.line,
            rotation: o.rotation,
            shapeType: o.shapeType,
            table: o.table,
            imageStyle: o.imageStyle
          } as ShapeInfo;
        });
        // 角色 / 固定 / 提示词 / 形状类型 / 语义层（按排序后的索引对齐）
        for (let i = 0; i < shapes.length; i++) {
          const o = srcShapes[i].o;
          if (!o) continue;
          const role = normalizeRole(o.role || defaultRoleForShape(shapes[i]));
          roles[i] = role;
          fixedShapes[i] = role === 'fixed';
          if (o.prompt) varLabels[i] = o.prompt;
          else if (o.varName) varLabels[i] = o.varName;
          if (o.shapeType) shapeTypes[i] = o.shapeType;
          if (role === 'ai_text') {
            const sem = (semantics[i] = {} as NonNullable<(typeof semantics)[number]>);
            if (o.semanticRole) sem.semanticRole = o.semanticRole;
            if (o.maxChars) sem.maxChars = o.maxChars;
            if (o.minChars) sem.minChars = o.minChars;
            if (o.translate) sem.translate = true;
            if (o.translateSource) sem.translateSource = o.translateSource;
          }
        }
        // 背景：模板中自选图片背景还原；其余（含无背景）按「跟随文档」处理
        const tb = template.background;
        background = tb ? { type: tb.type, source: tb.source || 'slide', color: tb.color, imageDataUrl: tb.imageDataUrl } : null;
        bgUseDefault = !tb || !!tb.source || tb.type !== 'picture' || !tb.imageDataUrl;
        bgOverride = bgUseDefault ? '' : tb!.imageDataUrl!;
        unsupported = [];
        // 页面信息（尺寸 / 背景行）——与「读取」后一致
        const bgSource = background?.source === 'layout' ? '（跟随文档默认背景）' : '';
        const bgDesc = background
          ? (BG_LABELS[background.type] || background.type) + (background.color ? `（${background.color}）` : '') + bgSource
          : '无（默认跟随文档背景）';
        infoEl.innerHTML = `
          <div class="page-info">
            <div class="pi-row"><span class="pi-k">尺寸</span><span class="pi-v">${slideSize.width.toFixed(2)} × ${slideSize.height.toFixed(2)} 英寸</span></div>
            <div class="pi-row">
              <span class="pi-k">背景</span>
              <span class="pi-v" style="display:inline-flex;align-items:center;gap:4px;flex-wrap:wrap">
                <span id="bg-desc" style="font-size:var(--font-aux)">${escapeHtml(bgDesc)}</span>
                <label class="fixed-cb-wrap" style="font-size:var(--font-aux);color:var(--text)">
                  跟随文档 <input type="checkbox" id="bg-default" ${bgUseDefault ? 'checked' : ''} />
                </label>
              </span>
              <span class="flex-spacer"></span>
              <button id="bg-pick" class="ghost" ${bgUseDefault ? 'disabled' : ''}>文件夹选择</button>
              <input type="file" id="bg-file" accept="image/*" style="display:none" />
            </div>
            <p class="pi-hint" id="bg-hint"></p>
            <div id="bg-thumb"></div>
          </div>`;
        infoEl.style.display = '';
        const infoModule = container.querySelector('#slide-info-module') as HTMLElement | null;
        if (infoModule) infoModule.style.display = '';
        // 背景交互（与「读取」后一致）
        const bgDefaultCb = infoEl.querySelector('#bg-default') as HTMLInputElement;
        const bgPickBtn = infoEl.querySelector('#bg-pick') as HTMLButtonElement;
        const bgFileInput = infoEl.querySelector('#bg-file') as HTMLInputElement;
        const bgThumb = infoEl.querySelector('#bg-thumb') as HTMLElement;
        const bgHint = infoEl.querySelector('#bg-hint') as HTMLElement;
        const syncBgUi = (): void => {
          clearErrorText(bgHint);
          bgPickBtn.disabled = bgUseDefault;
          if (bgUseDefault) {
            bgOverride = '';
            bgThumb.innerHTML = '';
            bgThumb.style.display = 'none';
            bgHint.textContent = '';
          }
        };
        bgDefaultCb.addEventListener('change', () => { bgUseDefault = bgDefaultCb.checked; syncBgUi(); });
        bgPickBtn.addEventListener('click', () => bgFileInput.click());
        bgFileInput.addEventListener('change', () => {
          const f = bgFileInput.files?.[0];
          if (!f) return;
          readFileAsDataUrl(f).then((dataUrl) => {
            bgOverride = dataUrl;
            bgThumb.innerHTML = '<img src="' + bgOverride + '" alt="背景预览" style="max-width:160px;max-height:90px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />'
              + '<span style="margin-left:6px">' + escapeHtml(f.name) + '</span>'
              + '<button id="bg-clear" class="ghost" style="margin-left:6px">清除</button>';
            bgThumb.style.display = 'block';
            clearErrorText(bgHint);
            bgHint.textContent = '已选择背景图片：' + escapeHtml(f.name) + '（保存时写入模板，不跟随文档默认）';
            bgThumb.querySelector('#bg-clear')!.addEventListener('click', () => {
              bgUseDefault = true;
              bgDefaultCb.checked = true;
              bgFileInput.value = '';
              syncBgUi();
            });
          }).catch(() => alert('读取图片失败'));
        });
        // 原模板自带自选背景图：直接显示缩略图（文件已随模板保存，无文件名）。
        // 防御：背景数据异常（如损坏的 dataURL）只降级为「跟随文档」，绝不让「页面」信息框整体消失
        if (!bgUseDefault && bgOverride) {
          try {
            bgThumb.innerHTML = '<img src="' + bgOverride + '" alt="背景预览" style="max-width:160px;max-height:90px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />'
              + '<button id="bg-clear" class="ghost" style="margin-left:6px">清除</button>';
            bgThumb.style.display = 'block';
            bgThumb.querySelector('#bg-clear')!.addEventListener('click', () => {
              bgUseDefault = true;
              bgDefaultCb.checked = true;
              syncBgUi();
            });
          } catch {
            bgUseDefault = true;
            bgDefaultCb.checked = true;
            bgOverride = '';
            syncBgUi();
          }
        }
        renderList();
        saveBtn.style.display = 'block';
        hasLoadedTemplateState = true;
        void createTemporaryPreviewSlide(template).catch((e) => {
          showToast('临时预览页插入失败：' + ((e as Error).message || String(e)), 4000);
        });
      } catch (e) {
        // 模板读取失败（已删除 / 路径变化 / 后端未启动）：提示并保持普通模式
        listEl.innerHTML = `<p class="error">载入模板失败：${escapeHtml((e as Error).message || String(e))}</p>`;
      }
    })();
  }
  container.querySelector('#read-slide')!.addEventListener('click', async () => {
    const readBtn = container.querySelector('#read-slide') as HTMLButtonElement;
    const readBtnText = readBtn.textContent || '';
    readBtn.disabled = true;
    readBtn.textContent = '读取中…';
    const readProgress = showProgress(readBtn, '正在准备…');
    try {
      await cleanupTempPreviewSlide();
      const result = await readCurrentSlide((phase) => readProgress.setText(phase));
      readProgress.done();
      shapes = result.shapes;
      slideSize = result.slideSize;
      background = result.background;
      unsupported = result.unsupported;
      // 重新读取后重置背景选择状态（避免上一页的自选背景图/取消勾选状态串到下一页）
      bgUseDefault = true;
      bgOverride = '';
      // 重新读取后清空上一页的角色/形状类型/提示词/语义标记，避免串页
      for (const k of Object.keys(roles)) delete roles[k];
      for (const k of Object.keys(shapeTypes)) delete shapeTypes[k];
      for (const k of Object.keys(varLabels)) delete varLabels[k];
      for (const k of Object.keys(fixedShapes)) delete fixedShapes[Number(k)];
      for (const k of Object.keys(semantics)) delete semantics[Number(k)];
      // 清空上一页的 AI 分析状态
      for (const k of Object.keys(recommendations)) delete recommendations[Number(k)];
      for (const k of Object.keys(appliedRecs)) delete appliedRecs[Number(k)];
      for (const k of Object.keys(ignoredRecs)) delete ignoredRecs[Number(k)];
      for (const k of Object.keys(userTouchedRoles)) delete userTouchedRoles[Number(k)];
      for (const k of Object.keys(userTouchedSemantics)) delete userTouchedSemantics[Number(k)];
      analysisState.running = false; analysisState.done = false; analysisState.aiUsed = false; analysisState.aiError = '';
      // 内置规则分析始终可用（免费、本地、不依赖 AI 服务）；异步执行不阻塞保存主流程。
      // 是否额外调用 AI 服务增强，由配置页「使用 AI 分析模板服务」开关决定（见 runAnalysis）
      void runAnalysis();
    } catch (e) {
      readProgress.done();
      listEl.innerHTML = `<p class="error">读取幻灯片失败：${escapeHtml((e as Error).message || String(e))}</p>`;
      readBtn.disabled = false;
      readBtn.textContent = readBtnText;
      return;
    }
    readBtn.disabled = false;
    readBtn.textContent = readBtnText;
    if (!shapes.length && !background) {
      listEl.innerHTML = '<p class="error">未读取到形状，请确认已选中一页并稍后重试。</p>';
      return;
    }
    hasLoadedTemplateState = true;

    // 后台回读文档 XML 里的精确样式（对齐/下划线等），补齐 Office.js 读不到/读不准的属性。
    // 需读取整个文档 zip（数十 MB），放到后台执行：界面先展示形状列表，不等待；
    // 保存流程也不等它 —— 保存后若回读完成会自动补存精确样式（见保存按钮）。
    // 回读状态更新：infoEl（含 #enrich-status）在 runEnrich 启动后才渲染，
    // 每次更新必须重新 querySelector（不能在开头捕获，否则是 null）。
    const setStatus = (text: string): void => {
      const el = container.querySelector('#enrich-status');
      if (el) el.textContent = text;
    };
    const runEnrich = async (): Promise<void> => {
      let readResult: Awaited<ReturnType<typeof Api.readAllBytes>>;
      // 优先通道：文件直读（Office.js document.url 或 COM 拿磁盘路径 → 后端秒级解析，45MB 也快）
      setStatus('回读状态：正在直读文档文件（秒级）…');
      const shapeBriefs = shapes.map((s) => ({ name: s.name, type: s.type, bounds: s.bounds, textStyle: s.textStyle }));
      let fileReadError = '';
      try {
        // ① Office.js document.url（本地文档返回 file:/// 路径，无 COM 依赖）
        let docPath = getDocumentPathFromOffice();
        if (!docPath) {
          let urlVal = '';
          try { urlVal = String((window as unknown as { Office?: any }).Office?.context?.document?.url || ''); } catch { /* ignore */ }
          fileReadError = 'Office.url 为空（值：' + (urlVal ? urlVal.slice(0, 80) : '（空）') + '）';
        } else {
          fileReadError = '';
        }
        // ② 回退：COM 拿 ActivePresentation.FullName
        if (!docPath) {
          const dp = await Api.getDocPath();
          if (dp.ok && dp.path) docPath = dp.path;
          else fileReadError += '；COM 拿路径失败：' + (dp.error || '');
        }
        if (!docPath) throw new Error(fileReadError || '无法获取文档路径（文档可能尚未保存到磁盘）');
        const idx = await getSelectedSlideIndex().catch(() => 1);
        const sid = await getSelectedSlideId();
        const pf = await Api.parseSlideFile(docPath, idx, shapeBriefs, sid ?? undefined);
        if (!pf.ok) throw new Error(pf.error || '文件直读失败');
        readResult = pf;
      } catch (fileErr) {
        // 文件直读不可用 → 回退 Office.js 慢通道（尽力，含未保存修改）
        setStatus('回读状态：文件直读不可用（' + ((fileErr as Error).message || String(fileErr)) + '），改用 Office.js 通道…');
        try {
          const [zipBytes, slideIndex] = await Promise.all([getDocumentZipBytes(), getSelectedSlideIndex()]);
          readResult = await Api.readAllBytes({ bytes: zipBytes, slideIndex, needBackground: true, shapes: shapeBriefs });
        } catch (e2) {
          throw new Error('读取文档失败：' + ((e2 as Error).message || String(e2)) + '（表格/样式回读已跳过）');
        }
      }
      let enriched = 0;
      readResult.styles.forEach((st, i) => {
        if (st && shapes[i] && shapes[i].textStyle) { shapes[i].textStyle = st; enriched++; }
      });
      if (readResult.background) {
        const bg = readResult.background as TemplateBackground & { source?: SlideBackgroundInfo['source'] };
        background = { type: bg.type, color: bg.color, imageDataUrl: bg.imageDataUrl, source: bg.source || 'slide' };
      }
      // 表格（XML 回读）：Office.js 读不到 GraphicFrame，从文档 zip 解析后并入 shapes
      const tables = readResult.tables || [];
      if (tables.length) {
        // 去重：Office.js 能读到表格外壳（如「Table 1」，type=other 但读不到数据），XML 回读才是真表格。
        // 之前只清 unsupported 警告没删重复形状 → 同一表格在列表出现两次。改为「原位转换」：
        // 命中同一位置的 Office.js 形状时直接改成 table 形状（保留 id 供删除、name 供回读匹配，
        // 索引不变 → styles/imageStyles 数组不错位）。
        const boundsMatch = (b1: { left: number; top: number; width: number; height: number } | undefined, b2: { left: number; top: number; width: number; height: number } | undefined): boolean => {
          if (!b1 || !b2) return false;
          return Math.abs(b1.left - b2.left) < 0.05 && Math.abs(b1.top - b2.top) < 0.05
            && Math.abs(b1.width - b2.width) < 0.05 && Math.abs(b1.height - b2.height) < 0.05;
        };
        for (const tb of tables) {
          const dupIdx = shapes.findIndex((s) => s.type !== 'table' && boundsMatch(s.bounds, tb.bounds));
          if (dupIdx >= 0) {
            const orig = shapes[dupIdx];
            shapes[dupIdx] = {
              ...orig, // 保留 id（删除功能）、name（回读匹配）等
              type: 'table', hasText: false, text: undefined, textStyle: undefined,
              bounds: tb.bounds, table: tb.table
            };
          } else {
            shapes.push({
              id: 'tbl' + shapes.length, type: 'table', name: tb.name || '表格',
              source: 'slide', hasText: false,
              bounds: tb.bounds, table: tb.table
            });
          }
        }
        // 表格合并后重跑规则：type 从 other → table，规则应推荐「表格位」（只补表格推荐，不覆盖 AI/已接受/已忽略）
        const recs2 = analyzeShapesByRules(shapes, slideSize);
        for (const r2 of recs2) {
          if (r2.recommendedRole === 'table' && !appliedRecs[r2.idx] && !ignoredRecs[r2.idx]) recommendations[r2.idx] = r2;
        }
        // 表格已可保存/生成：从「无法保存」列表移除（条目格式如「Table 1（Table，位于页面）」），
        // 并就地更新界面上已渲染的提示行（此前只改数组、UI 不刷新导致仍显示 Table 警告）
        unsupported = unsupported.filter((u) => !/（Table，位于/.test(u));
        const unsupLine = container.querySelector('#unsupported-line');
        if (unsupLine) {
          if (unsupported.length) unsupLine.textContent = '无法随模板保存/生成：' + unsupported.join('、');
          else unsupLine.remove();
        }
        showToast('已识别 ' + tables.length + ' 个表格（作为表格位保存）✓', 2500);
      }
      // 图片样式（柔化边缘矩形/圆角/裁剪等）并入 shapes[i].imageStyle
      let imgStyled = 0;
      (readResult.imageStyles || []).forEach((st, i) => {
        if (st && shapes[i]) { shapes[i].imageStyle = st; imgStyled++; }
      });
      // 「背景」行不再追加「（已回读 N 个元素的精确样式）」后缀（用户指定删除）
      enrichDone = true;
      const debugInfo = (readResult as { debug?: { slidePath?: string; hasTbl?: boolean; slideCount?: number; error?: string; tablePages?: number[] } }).debug;
      const debugText = tables.length ? '' : (debugInfo
        ? (debugInfo.error || '解析页 ' + (debugInfo.slidePath || '?') + ' 含表格结构: ' + (debugInfo.hasTbl ? '是' : '否') + '（共 ' + (debugInfo.slideCount ?? '?') + ' 页）' + (debugInfo.tablePages && debugInfo.tablePages.length ? '；全文档含表格的页：第 ' + debugInfo.tablePages.join('、') + ' 页' : ''))
        : '');
      setStatus('回读状态：样式 ' + enriched + ' 个' + (tables.length ? ' / 表格 ' + tables.length + ' 个 ✓' : ' / 未识别到表格' + (debugText ? '（' + debugText + '）' : '')) + (readResult.imageStyles?.length ? ' / 图片样式 ' + readResult.imageStyles.length + ' 个' : ''));
      renderList(); // 表格（XML 回读）可能已并入 shapes，刷新列表
      if (!tables.length && unsupported.some((u) => /（Table，位于/.test(u))) {
        showToast('未识别到表格数据（文档 zip 解析无表格；表格可能无法保存）', 4000);
      }
    };
    enrichTask = (async () => {
      try {
        await runEnrich();
      } catch (e) {
        // 文档读取/后台服务偶发失败：自动重试一次
        try {
          await new Promise((r) => setTimeout(r, 1200));
          await runEnrich();
        } catch (e2) {
          enrichFailed = true;
          const msg = (e2 as Error)?.message || String(e2);
          void Api.debugRead({ phase: 'save-enrich', error: msg, stack: (e2 as Error)?.stack }).catch(() => {});
          const bgLine = container.querySelector('#bg-desc');
          if (bgLine) bgLine.textContent += '（样式回读失败：' + escapeHtml(msg) + '，保存将使用 Office.js 实时值）';
          setStatus('回读状态：失败（' + msg + '）');
          showToast('样式/表格回读失败：' + msg + '（表格可能无法识别为表格位）', 5000);
        }
      }
    })();
    // 读取定位框颜色与显示时长（配置界面可修改）
    void refreshHighlightConfig();

    const bgSource = background?.source === 'layout' ? '（跟随文档默认背景）' : '';
    const bgDesc = background
      ? (BG_LABELS[background.type] || background.type) + (background.color ? `（${background.color}）` : '') + bgSource
      : '无（默认跟随文档背景；旧版 PowerPoint 可能不支持背景 API）';
    const unsupportedHtml = unsupported.length
      ? `<p class="error" id="unsupported-line">无法随模板保存/生成：${unsupported.map(escapeHtml).join('、')}</p>`
      : '';
    infoEl.innerHTML = `
      <div class="page-info">
        <div class="pi-row"><span class="pi-k">尺寸</span><span class="pi-v">${slideSize.width.toFixed(2)} × ${slideSize.height.toFixed(2)} 英寸</span></div>
        <div class="pi-row">
          <span class="pi-k">背景</span>
          <span class="pi-v" style="display:inline-flex;align-items:center;gap:4px;flex-wrap:wrap">
            <span id="bg-desc" style="font-size:var(--font-aux)">${escapeHtml(bgDesc)}</span>
            <label class="fixed-cb-wrap" style="font-size:var(--font-aux);color:var(--text)">
              跟随文档 <input type="checkbox" id="bg-default" checked />
            </label>
          </span>
          <span class="flex-spacer"></span>
          <button id="bg-pick" class="ghost" disabled>文件夹选择</button>
          <input type="file" id="bg-file" accept="image/*" style="display:none" />
        </div>
        <p class="pi-hint" id="bg-hint"></p>
        <div id="bg-thumb"></div>
        ${unsupportedHtml}
      </div>`;
    infoEl.style.display = '';
    const infoModule = container.querySelector('#slide-info-module') as HTMLElement | null;
    if (infoModule) infoModule.style.display = '';

    // 背景选项交互：默认勾选「跟随文档」；取消后可用「文件夹选择」选本地图片，可清除恢复默认
    const bgDefaultCb = infoEl.querySelector('#bg-default') as HTMLInputElement;
    const bgPickBtn = infoEl.querySelector('#bg-pick') as HTMLButtonElement;
    const bgFileInput = infoEl.querySelector('#bg-file') as HTMLInputElement;
    const bgThumb = infoEl.querySelector('#bg-thumb') as HTMLElement;
    const bgHint = infoEl.querySelector('#bg-hint') as HTMLElement;
    const syncBgUi = (): void => {
      clearErrorText(bgHint); // 背景已合法（沿用默认）时清除错误提示
      bgPickBtn.disabled = bgUseDefault;
      if (bgUseDefault) {
        bgOverride = '';
        bgThumb.innerHTML = '';
        bgThumb.style.display = 'none';
        bgHint.textContent = ''; // 默认不再显示「跟随文档」说明文字（用户指定彻底删除）
      }
    };
    bgDefaultCb.addEventListener('change', () => {
      bgUseDefault = bgDefaultCb.checked;
      syncBgUi();
    });
    bgPickBtn.addEventListener('click', () => bgFileInput.click());
    bgFileInput.addEventListener('change', () => {
      const f = bgFileInput.files?.[0];
      if (!f) return;
      readFileAsDataUrl(f).then((dataUrl) => {
        bgOverride = dataUrl;
        bgThumb.innerHTML = '<img src="' + bgOverride + '" alt="背景预览" style="max-width:160px;max-height:90px;border:1px solid #ddd;border-radius:4px;vertical-align:middle" />'
          + '<span style="margin-left:6px">' + escapeHtml(f.name) + '</span>'
          + '<button id="bg-clear" class="ghost" style="margin-left:6px">清除</button>';
        bgThumb.style.display = 'block';
        clearErrorText(bgHint); // 已选背景，清除错误提示
        bgHint.textContent = '已选择背景图片：' + escapeHtml(f.name) + '（保存时写入模板，不跟随文档默认）';
        bgThumb.querySelector('#bg-clear')!.addEventListener('click', () => {
          bgUseDefault = true;
          bgDefaultCb.checked = true;
          bgFileInput.value = '';
          syncBgUi();
        });
      }).catch(() => alert('读取图片失败'));
    });

    renderList();
    saveBtn.style.display = 'block';
  });

  saveBtn.addEventListener('click', async () => {
    const nameInput = container.querySelector('#tpl-name') as HTMLInputElement;
    const name = nameInput.value.trim();
    if (!name) {
      markInputError(nameInput, '！需要输入内容！请输入模板名称');
      return;
    }
    clearInputError(nameInput);
    // 版本说明（可选）：作为本次保存的新版本说明
    const changeNote = (container.querySelector('#tpl-note') as HTMLInputElement | null)?.value.trim() || undefined;
    // 分类文件夹：默认「未分类」；选「新建文件夹」时用输入框内容
    const folderSel = container.querySelector('#tpl-folder') as HTMLSelectElement | null;
    const folderNew = container.querySelector('#tpl-folder-new') as HTMLInputElement | null;
    let folder = '';
    if (folderSel && folderSel.value !== '') {
      folder = folderSel.value === '__new__'
        ? ((folderNew?.value || '').trim() || '')
        : folderSel.value;
      if (folderSel.value === '__new__' && !folder) {
        markInputError(folderNew, '！需要输入内容！请输入新文件夹名称');
        return;
      }
      clearInputError(folderNew);
    }

    // —— 重名检查（用户指定：同名模板不再静默创建新版本；编辑自己的模板才允许覆盖）——
    // 显示方式与「未填写模板名称」一致（红框 + 红字提示）
    const editingId = editId; // 编辑模式标记（进入保存页时已消费）
    let confirmOverwrite = false;
    let existingTemplate: { name: string; template: TemplateDoc } | null = null;
    try { existingTemplate = await Api.getTemplate(name, folder); } catch { /* 不存在或读取失败：按新模板保存 */ }
    // —— 编辑模式保存方式弹窗（用户明确要求）：编辑后不能直接覆盖当前版本，必须选择 ——
    // 覆盖：替换当前版本（历史版本仍保留，可随时恢复）；存为新版本：保留当前版本，另存为新版本；取消：留在界面继续编辑。
    if (editingId) {
      if (existingTemplate) {
        // 判断「是不是正在编辑的模板」：优先按模板 id（=目录名）比对，避免编辑时改名撞上
        // 另一个同名模板被误判为自己而覆盖错误对象；极旧模板无 id 时回退 name+folder 判定。
        const tId = existingTemplate.template?.id;
        const isSelf = tId ? tId === editId : (existingTemplate.template?.name === name && folder === editFolder);
        if (!isSelf) {
          markInputError(nameInput, '！模板名称已存在！请换个名称，或到模板库点「编辑」修改已有模板');
          return;
        }
      }
      // 改名成全新模板名（existingTemplate 为空）：直接保存 = 新建模板（旧模板不受影响），无需弹窗
      if (existingTemplate) {
        const choice = await showModal({
          title: '保存方式',
          message: '已编辑模板，选择保存方式：\n· 覆盖：替换当前版本（历史版本仍保留，可随时恢复）\n· 存为新版本：保留当前版本，另存为新版本',
          buttons: [
            { id: 'overwrite', label: '覆盖', kind: 'primary' },
            { id: 'newversion', label: '存为新版本', kind: 'secondary' },
            { id: 'cancel', label: '取消', kind: 'danger' }
          ]
        });
        if (choice !== 'overwrite' && choice !== 'newversion') return;
        confirmOverwrite = choice === 'overwrite';
      }
    } else if (existingTemplate) {
      // 非编辑模式：同名模板已存在 → 报错（不再静默创建新版本，也不允许覆盖他人模板）
      markInputError(nameInput, '！模板名称已存在！请换个名称，或到模板库点「编辑」修改已有模板');
      return;
    }
    // —— 同名模板标记继承（P1-C）：保存同名模板时自动复用旧模板的角色标记 ——
    // 用户改完一页重存时，新形状按「位置重叠 + 尺寸相近 + 名称相同」匹配旧模板形状，
    // 把 role / semanticRole / prompt / varName 预填到新形状，免去重新逐个标注。
    // 仅在用户尚未手动设置过该形状角色时生效；不匹配的形状保持当前默认/推荐。
    // 编辑模式（从模板库「编辑」进入）已完整载入旧模板角色，无需继承，直接跳过。
    if (folderSel && !editId) {
      try {
        const existing = await Api.getTemplate(name, folder);
        const oldShapes: TemplateShape[] = (existing?.template?.shapes || []);
        if (oldShapes.length && shapes.length) {
          let inherited = 0;
          const usedNew = new Set<number>(); // 已匹配的新形状索引
          const matchScore = (o: TemplateShape, n: typeof shapes[number]): number => {
            let score = 0;
            const ob = o.bounds || { left: 0, top: 0, width: 0, height: 0 };
            const nb = n.bounds || { left: 0, top: 0, width: 0, height: 0 };
            // 位置重叠（中心点距离归一化）
            const ocx = ob.left + ob.width / 2, ocy = ob.top + ob.height / 2;
            const ncx = nb.left + nb.width / 2, ncy = nb.top + nb.height / 2;
            const dist = Math.hypot(ocx - ncx, ocy - ncy) / Math.max(slideSize.width, slideSize.height);
            if (dist < 0.05) score += 4; else if (dist < 0.15) score += 2; else if (dist < 0.3) score += 1;
            // 尺寸相近
            const dw = ob.width > 0 ? Math.abs(ob.width - nb.width) / ob.width : 1;
            const dh = ob.height > 0 ? Math.abs(ob.height - nb.height) / ob.height : 1;
            if (dw < 0.1 && dh < 0.1) score += 2; else if (dw < 0.3 && dh < 0.3) score += 1;
            // 名称相同
            if (o.name && n.name && o.name === n.name) score += 3;
            // 类型一致
            if ((o.type === 'picture') === (n.type === 'picture')) score += 1;
            return score;
          };
          // 贪心：对每个旧形状，找分数最高的未占用新形状
          const pairs: { o: TemplateShape; n: number; score: number }[] = [];
          for (const o of oldShapes) {
            let bestN = -1, bestScore = 0;
            shapes.forEach((n, ni) => {
              if (usedNew.has(ni)) return;
              const sc = matchScore(o, n);
              if (sc > bestScore) { bestScore = sc; bestN = ni; }
            });
            if (bestN >= 0 && bestScore >= 4) {
              usedNew.add(bestN);
              pairs.push({ o, n: bestN, score: bestScore });
            }
          }
          for (const { o, n } of pairs) {
            if (userTouchedRoles[n]) continue; // 用户已手动设置过 → 不覆盖
            if (!roles[n] || roles[n] === defaultRoleForShape(shapes[n])) {
              const orole = normalizeRole(o.role || 'fixed'); // manual_var → ai_text
              roles[n] = orole;
              fixedShapes[n] = orole === 'fixed'; // 旧模板固定元素 → 勾选固定
              if (orole === 'ai_text') {
                const sem = (semantics[n] = semantics[n] || {});
                if (o.semanticRole) sem.semanticRole = o.semanticRole;
                if (o.maxChars) sem.maxChars = o.maxChars;
                if (o.minChars) sem.minChars = o.minChars;
                if (o.prompt) varLabels[n] = o.prompt;
                else if (o.varName) varLabels[n] = o.varName; // 旧手动变量位：varName 当提示词
              }
              inherited++;
            }
          }
          if (inherited) {
            showToast('已继承同名模板 ' + inherited + ' 个元素的角色标记 ✓（可手动调整）', 2500);
          }
        }
      } catch { /* 同名模板不存在或读取失败：静默，按新模板保存 */ }
    }

    // 校验：提示词可选（向导中可手动输入或仅用主题生成），不再强制填写
    shapes.forEach((s, idx) => {
      if (fixedShapes[idx]) return; // 固定元素无需提示词
      const role = roles[idx] || defaultRoleForShape(s);
      if (role !== 'ai_text') return;
      const labelInput = container.querySelector(`input.var-label[data-idx="${idx}"]`) as HTMLInputElement | null;
      if (labelInput) clearInputError(labelInput);
    });
    if (unsupported.length) {
      const choice = await showModal({
        title: '提示',
        message: `以下元素无法随模板保存/生成：\n${unsupported.join('\n')}`,
        buttons: [
          { id: 'cancel', label: '取消', kind: 'secondary' },
          { id: 'save', label: '仍然保存', kind: 'primary' }
        ]
      });
      if (choice !== 'save') return;
    }

    // 背景：勾选「跟随文档」→ 保存当前文档/母版解析到的背景；未勾选 → 必须已自选图片
    const bgDefaultCb = container.querySelector('#bg-default') as HTMLInputElement;
    bgUseDefault = bgDefaultCb ? bgDefaultCb.checked : true;
    if (!bgUseDefault && !bgOverride) {
      markErrorText(container.querySelector('#bg-hint'), '！需要输入内容！请先选择一张背景图片，或勾选「跟随文档」。');
      alert('请先点击「文件夹选择」选择一张背景图，或勾选「跟随文档」。');
      return;
    }
    clearErrorText(container.querySelector('#bg-hint'));
    const finalBg = await buildTemplateBackground();

    // 保存前清洗 textStyle：剔除 false / 默认值，减小模板 JSON 体积。
    // 注意 margin 全 0 必须保留（生成端需显式写 0 边距，否则 PowerPoint 用默认 0.1"/0.05"）。
    const cleanTextStyle = (ts: typeof shapes[number]['textStyle']) => {
      if (!ts) return ts;
      const out: Record<string, unknown> = { ...ts };
      for (const k of ['bold', 'italic', 'strikethrough', 'doubleStrikethrough', 'subscript', 'superscript']) {
        if (out[k] === false) delete out[k];
      }
      if (out.autoFit === 'none') delete out.autoFit;
      if (out.wordWrap === true) delete out.wordWrap;
      if (out.align === 'left') delete out.align;   // 生成端缺省即左对齐
      if (out.valign === 'top') delete out.valign;  // 生成端缺省即顶部
      if (out.underline === false) delete out.underline;
      return out as typeof shapes[number]['textStyle'];
    };

    const buildShapesOut = () => shapes.map((s, idx) => {
      // 固定勾选优先：勾选 = 固定元素（原样保留）；否则按下拉选择（文字/图片）或类型默认
      const isFixed = fixedShapes[idx] === true || (fixedShapes[idx] === undefined && (defaultRoleForShape(s) === 'fixed' || isIconPicture(s)));
      const role = isFixed ? 'fixed' : normalizeRole(roles[idx] || defaultRoleForShape(s));
      const labelInput = container.querySelector(`input.var-label[data-idx="${idx}"]`) as HTMLInputElement | null;
      return {
        id: `shp${idx}`,
        type: s.type,
        name: s.name,
        role,
        source: s.source,
        shapeType: (s.type === 'rectangle' || s.type === 'line') ? (shapeTypes[idx] || s.shapeType || (s.type === 'line' ? 'line' : 'rect')) : undefined,
        rotation: s.rotation,
        bounds: s.bounds,
        textStyle: cleanTextStyle(s.textStyle),
        fill: s.fill,
        line: s.line,
        imageStyle: s.type === 'picture' && s.imageStyle
          ? (Object.keys(s.imageStyle).length === 1 && s.imageStyle.imageDataUrl ? undefined
            : (() => { const { imageDataUrl, ...rest } = s.imageStyle; return Object.keys(rest).length ? rest : undefined; })())
          : undefined,
        table: s.type === 'table' ? s.table : undefined,
        prompt: role === 'ai_text' ? (labelInput?.value || '') : undefined,
        varName: role === 'manual_var' ? (labelInput?.value || `变量${idx + 1}`) : undefined,
        // 图片随模板保存（P2 优化）：固定图片位（图标/装饰/Logo）把图片本体 dataURL 存入 content，
        // 生成端 slideBuilder 对 fixed+picture+content=data: 原样渲染；AI 图片位不存本体（生成时由用户选图替换）
        content: role === 'fixed' ? (s.type === 'picture' ? (s.imageStyle?.imageDataUrl || undefined) : (s.hasText ? s.text : undefined)) : undefined,
        // 模板语义层：仅 AI 文本位写入语义字段；其余角色（fixed/图片/变量）不强制拥有
        ...(role === 'ai_text' ? {
          semanticRole: semantics[idx]?.semanticRole || undefined,
          maxChars: semantics[idx]?.maxChars || undefined,
          minChars: semantics[idx]?.minChars || undefined,
          translate: semantics[idx]?.translate || undefined,
          // translateSource 存的是来源形状的数组索引（或 'theme'）；保存时映射为稳定 id shpN
          // （兼容旧数据：已存 shpN 或其它字符串则原样保留，向导端有容错）
          translateSource: (() => {
            const ts = semantics[idx]?.translateSource;
            if (ts === undefined || ts === '' || ts === 'theme') return ts || undefined;
            const n = Number(ts);
            return (Number.isInteger(n) && n >= 0) ? 'shp' + n : ts;
          })()
        } : {})
      };
    });

    saveBtn.disabled = true;
    saveBtn.textContent = '保存中…';
    const saveProgress = showProgress(saveBtn, '正在保存模板…');
    const restoreSaveBtn = () => { saveBtn.disabled = false; saveBtn.textContent = '保存'; };
    // 注意：保存流程不等待样式回读（读整个文档 zip 可能很慢），也绝不调用 slide.getImageAsBase64()
    // （它在复杂/大文档上会让 PowerPoint 任务窗格一起冻结）。保存立即完成；
    // 回读完成后在后台用精确样式重新保存模板并补齐预览图，失败不影响模板本身。
    const savedWithStyles = enrichDone;
    try {
      // 预览图 = 模板结构示意图（与模板库展示一致）。不再截取 PowerPoint 当前页：
      // 编辑模板时当前页往往不是模板对应页，且截取失败会残留旧预览（如改角色后仍是旧占位图）。
      const { renderTemplateDiagram } = await import('../lib/templateDiagram.js');
      const shapesOut = buildShapesOut();
      const diagram = renderTemplateDiagram({ schemaVersion: 1, name, slideSize, background: finalBg, shapes: shapesOut });
      await Api.saveTemplate({
        name,
        folder,
        template: { schemaVersion: 1, name, slideSize, background: finalBg, shapes: shapesOut },
        preview: diagram || '',
        changeNote, // 版本说明（可选）
        updateCurrent: confirmOverwrite ? true : undefined // 编辑模式确认覆盖 = 替换当前版本
      });
      saveProgress.done();
      showToast('保存成功 ✓');
      if (enrichFailed) showToast('已保存，但样式回读失败（可能缺少居中/下划线等精确样式）', 3000);
      void (async () => {
        // 等样式回读结束（若尚未完成）；失败则直接放弃补存
        if (enrichTask) { try { await enrichTask; } catch { return; } }
        if (!enrichDone || savedWithStyles) return;
        // 用回读到的精确样式覆盖保存（模板 JSON 以最后一次保存为准）
        try {
          // 补存不产生新版本：修正当前版本（updateCurrent），避免一次保存操作产生多个版本
          await Api.saveTemplate({ name, folder, template: { schemaVersion: 1, name, slideSize, background: finalBg, shapes: buildShapesOut() }, preview: '', updateCurrent: true });
        } catch { /* 补存失败不影响已保存的模板 */ }
      })();
      allowLeaveWithoutConfirm = true;
      await cleanupTempPreviewSlide();
      setRouteLeaveGuard('#save', null);
      location.hash = '#library';
    } catch (e) {
      saveProgress.done();
      alert(`保存失败：${(e as Error).message}`);
    } finally {
      restoreSaveBtn();
    }
  });
}
