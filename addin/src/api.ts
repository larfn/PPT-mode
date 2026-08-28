// 表格适配结果类型（引擎 tableModel.ts 输出；buildSlideBase64 可选携带，向后兼容旧调用）
import type { FitResult } from './lib/tableModel.js';
export type { FitResult, FitCell } from './lib/tableModel.js';

const BASE_URL = '/api';

// P2-E 安全加固：后端启用鉴权时，所有 API 请求需带 X-Auth-Token（一次性 token）。
// 从同源端点 /api/runtime 引导获取并缓存；后端未启用（旧版/测试）时返回空串、不带头，行为不变。
let cachedToken: string | null = null;

async function getAuthToken(): Promise<string> {
  if (cachedToken !== null) return cachedToken;
  try {
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), 5000);
    const res = await fetch('/api/runtime', { signal: controller.signal });
    window.clearTimeout(timer);
    if (res.ok) {
      const body = await res.json().catch(() => null);
      cachedToken = body?.token || '';
    } else {
      cachedToken = '';
    }
  } catch { /* 后端未启用鉴权/不可达：不带头，等后续请求报错 */ cachedToken = ''; }
  return cachedToken ?? '';
}

async function request<T>(path: string, init?: RequestInit, timeoutMs?: number): Promise<T> {
  const controller = new AbortController();
  const timer = timeoutMs ? window.setTimeout(() => controller.abort(), timeoutMs) : undefined;
  try {
    const token = await getAuthToken();
    const res = await fetch(`${BASE_URL}${path}`, {
      ...init,
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { 'X-Auth-Token': token } : {}),
        ...(init?.headers || {})
      }
    });
    if (!res.ok) {
      if (res.status === 413) throw new Error('文档过大，超出本地服务请求上限');
      const body = await res.json().catch(() => ({ error: res.statusText }));
      throw new Error(body.error || `HTTP ${res.status}`);
    }
    return res.json() as Promise<T>;
  } catch (e) {
    if ((e as Error)?.name === 'AbortError') throw new Error('请求超时，请重试');
    throw e;
  } finally {
    if (timer) window.clearTimeout(timer);
  }
}

export interface TemplateMeta { id: string; name: string; folder: string; preview: string; updatedAt: string | null; version?: number; hasPreview?: boolean; }

// 统一图片搜索结果模型（Image Provider 层定义；前端只依赖这些字段）
// 套版（Deck）：有序模板引用序列 + 每页填参规格
export interface DeckPageSpec {
  templateId: string;
  templateFolder?: string;
  templateVersion?: string;   // 固定版本（如 'v3'），保证可复现
  variables?: Record<string, string>; // 变量默认值（向导中可覆盖）
  image?: { prompt?: string; provider?: string };  // 图片位搜索提示（可选）
  text?: { instruction?: string };                 // 文字生成指令（可选，覆盖模板默认）
}
export interface DeckDoc {
  schemaVersion: number;
  name: string;
  slideSize?: { width: number; height: number };
  pages: DeckPageSpec[];
  createdAt?: string;
  updatedAt?: string;
}

// 表格位（XML 回读）：结构 + 单元格样式 + 尺寸；生成时按保存长高自适应重建
export interface TableCellSpec {
  row: number; col: number;
  rowspan?: number; colspan?: number;
  text?: string;
  textStyle?: { font?: string; eaFont?: string; size?: number; bold?: boolean; italic?: boolean; color?: string; align?: string };
  fill?: string;
  valign?: string;
  margin?: { top?: number; right?: number; bottom?: number; left?: number }; // 磅（a:tcMar）
  border?: { left?: { width?: number; color?: string }; right?: { width?: number; color?: string }; top?: { width?: number; color?: string }; bottom?: { width?: number; color?: string } };
}
export interface TableSpec {
  rows: number; cols: number;
  colWidths?: number[];
  rowHeights?: number[];
  tblPr?: { firstRow?: boolean; bandRow?: boolean; tableStyleId?: string };
  cells: TableCellSpec[];
}

export interface ImageResult {
  id: string;
  thumbnailUrl: string;
  imageUrl: string;
  width: number | null;
  height: number | null;
  source: string;
  sourceUrl: string | null;
  title: string;
  author: string | null;
  license: string | null;
  mimeType: string | null;
  query: string;
  provider: string;
}
// —— 图源管理（默认源 + 自定义导入源）——
// 一个图源 = JSON 定义：endpoint 支持 {query}/{count}/{page}/{start}/{key} 占位符；
// resultsPath / fields 支持点分路径与数组下标（如 data.items、imageinfo[0].url）
export interface ImageSourceDef {
  id: string;
  name: string;
  endpoint: string;
  headers?: Record<string, string>;
  cookies?: Record<string, string>;
  resultsPath?: string;
  fields?: Record<string, string>;
  key?: string;
  enabled?: boolean;
  allowPrivate?: boolean;
  note?: string;
  keyRequired?: boolean;
  builtin?: boolean;
  preset?: boolean;
}
export interface ImageSourcesResponse {
  builtins: ImageSourceDef[];
  custom: ImageSourceDef[];
  presets: ImageSourceDef[];
}
export interface ImportSourcesResult {
  added: { id: string }[];
  skipped: { id?: string; reason: string }[];
}
export interface TemplateFolder { name: string; count: number; }
export interface TemplateShape {
  id: string; type: string; role: string; source?: string; name?: string;
  shapeType?: string; rotation?: number;
  bounds: { left: number; top: number; width: number; height: number };
  textStyle?: {
    font?: string; size?: number; bold?: boolean; italic?: boolean; color?: string; align?: string;
    valign?: string; underline?: boolean; strikethrough?: boolean; doubleStrikethrough?: boolean;
    subscript?: boolean; superscript?: boolean;
    margin?: { top: number; right: number; bottom: number; left: number };
    autoFit?: string; wordWrap?: boolean;
    eaFont?: string;
  };
  fill?: { type?: string; color?: string };
  line?: { color?: string; weight?: number; visible?: boolean };
  imageStyle?: {
    shape?: string;
    softEdgeEmu?: number;
    srcRect?: { l: number; t: number; r: number; b: number };
    spPrXml?: string; // 完整图片样式：spPr 除 xfrm 的原样 XML（边框/阴影/反射/辉光/柔化边缘/棱台/三维旋转/形状）
    blipAttrs?: string; // a:blip 除 r:embed 的属性（透明度等）
    blipKids?: string; // a:blip 子元素（图片更正/艺术效果等）
    imageDataUrl?: string; // 图片本体（dataURL）：保存模板时随固定图片位写入 shape.content，生成端原样渲染
  };
  prompt?: string; varName?: string; placeholder?: string; content?: string;
  table?: TableSpec; // 表格位（XML 回读）
  // —— 模板语义层（Semantic Template Layer，全部可选，旧模板缺省即默认行为）——
  // 语义角色：title/subtitle/body/bullet/caption/label/data/conclusion/page_number/tag/other
  semanticRole?: string;
  // 内容类型：自由文本描述（如「数字指标」「公司名」「日期」），0/缺失 = 不限制
  contentType?: string;
  // 是否必填（AI 文本位是否允许留空）
  required?: boolean;
  // 最多字符数（0 或缺失 = 不限制）
  maxChars?: number;
  // 最多行数（0 或缺失 = 不限制）
  maxLines?: number;
  // 最少字符数（0 或缺失 = 不要求）
  minChars?: number;
  // 建议长度（字符数，0 或缺失 = 无建议）
  preferredLength?: number;
  // 生成指令：给 AI 的补充要求（如「用一句话概括」）
  generationInstruction?: string;
  // 自动翻译副标题：true = 该 AI 文本位生成时把 translateSource 所指原文翻译为英文副标题，而非自由发挥
  translate?: boolean;
  // 翻译原文来源形状 id；特殊值 'theme' = 全局主题（向导步骤3顶部输入框）
  translateSource?: string;
}
export interface TemplateBackground {
  type: 'none' | 'solid' | 'picture' | 'gradient' | 'pattern' | 'unsupported';
  source?: 'slide' | 'layout' | 'master';
  color?: string;
  imageDataUrl?: string;
}
export interface TemplateDoc {
  schemaVersion: number; name: string;
  id?: string; // 后端保存时写入的目录名（safeId(name)），编辑模式用于识别「正在编辑的模板」
  slideSize: { width: number; height: number };
  background?: TemplateBackground;
  shapes: TemplateShape[];
  // —— 模板版本控制（可选，旧模板缺省）——
  version?: number;
  versionId?: string;
  createdAt?: string;
  changeNote?: string;
}

// —— 部署一致性诊断（设置页「关于/系统诊断」 + 启动自检）——
export interface VersionStatus {
  ok: boolean;
  name: string;
  appVersion: string;      // 统一版本号 YYYY.MM.DD.NN（前后端同版本）
  apiVersion: string;      // API 契约版本（破坏性变更 +1）
  builtAt: string;
  frontend: { version: string; sizeKB: number | null };   // sizeKB = 实际被后端伺服的主 JS 大小
  backend: { version: string; exeSizeKB: number | null }; // exeSizeKB 仅打包 exe 模式有值
  mcp: { running: boolean; pid: number | null; lastSeenAt: string | null };
  port: number | null;     // 后端实际监听端口（3788 被占时可能 +1）
}
export interface RuntimeStatus { ok: boolean; port: number | null; token: string; startedAt: string | null; }

export interface VersionInfo {
  version: number;
  versionId: string;
  createdAt: string | null;
  updatedAt: string | null;
  changeNote: string | null;
  isCurrent: boolean;
}

export const Api = {
  getAppVersion: () => request<VersionStatus>('/version'),
  getRuntime: () => request<RuntimeStatus>('/runtime'),
  listTemplates: () => request<TemplateMeta[]>('/templates'),
  listFolders: () => request<TemplateFolder[]>('/templates/folders'),
  getTemplate: (id: string, folder = '') => request<{ name: string; template: TemplateDoc }>(
    '/templates/' + encodeURIComponent(id) + '?folder=' + encodeURIComponent(folder)),
  // 默认创建新版本（v1→v2…）；updateCurrent=true 修正当前版本（后台补存样式/预览，不产生新版本）
  saveTemplate: (payload: { name: string; folder?: string; template: TemplateDoc; preview: string; changeNote?: string; updateCurrent?: boolean }) =>
    request<{ id: string; name: string; version?: number; versionId?: string; created?: boolean }>('/templates', { method: 'POST', body: JSON.stringify(payload) }, 60000),
  deleteTemplate: (id: string, folder = '') => request<{ ok: boolean }>(
    '/templates/' + encodeURIComponent(id) + '?folder=' + encodeURIComponent(folder), { method: 'DELETE' }),
  // —— 回收站（P2-F）：删除模板 = 移入 .回收站，可恢复 / 彻底删除 / 清空 ——
  listRecycleBin: () =>
    request<{ items: { entryId: string; id: string; folder: string; name: string; deletedAt: string | null; preview: string }[] }>('/templates/recycle'),
  restoreRecycle: (entryId: string) =>
    request<{ ok: boolean; id: string; folder: string }>('/templates/recycle/restore', { method: 'POST', body: JSON.stringify({ entryId }) }),
  purgeRecycle: (entryId: string) =>
    request<{ ok: boolean }>('/templates/recycle/purge', { method: 'POST', body: JSON.stringify({ entryId }) }),
  emptyRecycleBin: () =>
    request<{ ok: boolean; removed: number }>('/templates/recycle', { method: 'DELETE' }),
  // —— 模板版本控制 ——
  listVersions: (id: string, folder = '') =>
    request<{ currentVersion?: number; currentVersionId?: string; versions: VersionInfo[] }>(
      '/templates/' + encodeURIComponent(id) + '/versions?folder=' + encodeURIComponent(folder)),
  getVersion: (id: string, folder: string, versionId: string) =>
    request<{ version: TemplateDoc; previewUrl: string | null; isCurrent: boolean }>(
      '/templates/' + encodeURIComponent(id) + '/versions/' + encodeURIComponent(versionId) + '?folder=' + encodeURIComponent(folder)),
  restoreVersion: (id: string, folder: string, versionId: string) =>
    request<{ ok: boolean; version: number; versionId: string }>(
      '/templates/' + encodeURIComponent(id) + '/restore?folder=' + encodeURIComponent(folder),
      { method: 'POST', body: JSON.stringify({ versionId }) }),
  setCurrentVersion: (id: string, folder: string, versionId: string) =>
    request<{ ok: boolean }>(
      '/templates/' + encodeURIComponent(id) + '/set-current?folder=' + encodeURIComponent(folder),
      { method: 'POST', body: JSON.stringify({ versionId }) }),
  deleteVersion: (id: string, folder: string, versionId: string) =>
    request<{ ok: boolean }>(
      '/templates/' + encodeURIComponent(id) + '/versions/' + encodeURIComponent(versionId) + '?folder=' + encodeURIComponent(folder),
      { method: 'DELETE' }),
  getConfig: () => request<any>('/config'),
  saveConfig: (cfg: any) => request<{ ok: boolean }>('/config', { method: 'PUT', body: JSON.stringify(cfg) }),
  searchImages: (query: string, count = 9, page = 1, provider = 'baidu_page') =>
    request<{ images: ImageResult[]; page: number; hasMore: boolean; provider?: string; providerName?: string | null; error?: { code: string; message: string } }>(
      '/images/search', { method: 'POST', body: JSON.stringify({ provider, query, count, page }) }),
  getImageProviders: () =>
    request<{ providers: { id: string; name: string }[] }>('/images/providers'),
  // —— 图源管理（自定义源增删改 / 批量导入 / 测试 / 预置模板）——
  getImageSources: () =>
    request<ImageSourcesResponse>('/images/sources'),
  saveImageSource: (source: ImageSourceDef) =>
    request<{ ok: boolean; source?: ImageSourceDef }>('/images/sources', { method: 'POST', body: JSON.stringify({ source }) }),
  deleteImageSource: (id: string) =>
    request<{ ok: boolean }>('/images/sources/' + encodeURIComponent(id), { method: 'DELETE' }),
  importImageSources: (sources: ImageSourceDef[]) =>
    request<{ ok: boolean; added: { id: string }[]; skipped: { id?: string; reason: string }[] }>(
      '/images/sources/import', { method: 'POST', body: JSON.stringify({ sources }) }),
  testImageSource: (source: ImageSourceDef, query?: string) =>
    request<{ ok: boolean; resultsCount: number; error?: { code: string; message: string } | null; sample?: { title?: string; imageUrl?: string; thumbnailUrl?: string } | null }>(
      '/images/sources/test', { method: 'POST', body: JSON.stringify({ source, query: query || '猫' }) }),
  downloadImage: (url: string, provider?: string) =>
    request<{ taskId: string }>('/images/download', { method: 'POST', body: JSON.stringify({ url, provider }) }),
  getDownloadStatus: (taskId: string) =>
    request<{ done: boolean; error?: string; received: number; total: number | null; fileName: string; filePath: string; dataUrl: string }>(`/images/download/${encodeURIComponent(taskId)}`),
  openDownloadDir: () =>
    request<{ ok: boolean; dir: string }>('/images/open-downloads', { method: 'POST', body: JSON.stringify({}) }),
  // constraints: 模板语义层约束（可选）；clean: 输出模式清洗规则（plain/maxChars/maxLines，可选）；
  // 旧调用不带这些字段时行为完全不变（后端仍做 Markdown 语法兜底清洗）
  generateText: (systemPrompt: string, userPrompt: string, constraints?: {
    semanticRole?: string; contentType?: string; required?: boolean;
    maxChars?: number; maxLines?: number; minChars?: number; preferredLength?: number;
    generationInstruction?: string;
  }, clean?: { plain?: boolean; maxChars?: number; maxLines?: number }) =>
    request<{ text: string }>('/text/generate', { method: 'POST', body: JSON.stringify({ systemPrompt, userPrompt, constraints, clean }) }),
  buildSlideBase64: (payload: { template: TemplateDoc; images?: Record<string, string>; imageDataUrl?: string; texts: Record<string, string>; vars: Record<string, string>; tableData?: Record<string, string[][]>; tables?: Record<string, FitResult> }) =>
    request<{ base64: string }>('/slides/build', { method: 'POST', body: JSON.stringify(payload) }),
  exportDebugSlide: (base64: string) =>
    request<{ ok: boolean; filePath: string }>('/slides/export-debug', { method: 'POST', body: JSON.stringify({ base64 }) }),
  // 大文档回退通道：COM 拿文档磁盘路径 → 后端直读文件解析（绕开 Office.js 慢通道）
  getDocPath: () =>
    request<{ ok: boolean; path?: string; error?: string }>('/slides/doc-path'),
  parseSlideFile: (filePath: string, slideIndex: number, shapes?: { name: string; type?: string; bounds: { left: number; top: number; width: number; height: number }; textStyle?: TemplateShape['textStyle'] }[], slideId?: number) =>
    request<{ ok: boolean; background?: TemplateBackground | null; styles: (TemplateShape['textStyle'] | null)[]; imageStyles?: (TemplateShape['imageStyle'] | null)[]; tables?: { name?: string; bounds: { left: number; top: number; width: number; height: number }; table: TableSpec }[]; error?: string; debug?: { slidePath?: string; hasGraphicFrame?: boolean; hasTbl?: boolean; slideCount?: number; error?: string; tablePages?: number[] } }>(
      '/slides/parse-file', { method: 'POST', body: JSON.stringify({ path: filePath, slideIndex, shapes, slideId }) }, 60000),
  debugRead: (payload: unknown) =>
    request<{ ok: boolean }>('/slides/debug-read', { method: 'POST', body: JSON.stringify(payload) }),
  // AI 待写队列（MCP/ChatGPT 生成的页面）
  aiPendingList: () =>
    request<{ id: string; templateName: string; createdAt: string; written: boolean }[]>('/ai/pending'),
  aiPendingGet: (id: string) =>
    request<{ id: string; templateName: string; createdAt: string; written: boolean; base64: string }>('/ai/pending/' + encodeURIComponent(id)),
  aiPendingWrite: (id: string) =>
    request<{ ok: boolean }>('/ai/pending/' + encodeURIComponent(id) + '/write', { method: 'POST', body: JSON.stringify({}) }),
  aiPendingDelete: (id: string) =>
    request<{ ok: boolean }>('/ai/pending/' + encodeURIComponent(id), { method: 'DELETE' }),
  aiPendingClearAll: () =>
    request<{ ok: boolean; removed: number }>('/ai/pending', { method: 'DELETE' }),
  // 二进制上传文档 zip（省 33% 数据 + 免 64MB JSON 解析）：body = [4 字节大端 shapes JSON 长度][shapes JSON][zip 字节]
  // —— 套版（Deck）：整份/多页生成 ——
  listDecks: () =>
    request<{ id: string; name: string; folder: string; pageCount: number; preview: string; updatedAt: string | null }[]>('/decks'),
  saveDeck: (payload: { name: string; folder?: string; deck: DeckDoc; preview?: string }) =>
    request<{ id: string; name: string }>('/decks', { method: 'POST', body: JSON.stringify(payload) }),
  getDeck: (id: string, folder = '') =>
    request<{ name: string; deck: DeckDoc }>('/decks/' + encodeURIComponent(id) + '?folder=' + encodeURIComponent(folder)),
  deleteDeck: (id: string, folder = '') =>
    request<{ ok: boolean }>('/decks/' + encodeURIComponent(id) + '?folder=' + encodeURIComponent(folder), { method: 'DELETE' }),
  listDeckRecycleBin: () =>
    request<{ items: { entryId: string; id: string; folder: string; name: string; deletedAt: string | null; preview: string; pageCount: number }[] }>('/decks/recycle'),
  restoreDeckRecycle: (entryId: string) =>
    request<{ ok: boolean; id: string; folder: string }>('/decks/recycle/restore', { method: 'POST', body: JSON.stringify({ entryId }) }),
  purgeDeckRecycle: (entryId: string) =>
    request<{ ok: boolean }>('/decks/recycle/purge', { method: 'POST', body: JSON.stringify({ entryId }) }),
  emptyDeckRecycleBin: () =>
    request<{ ok: boolean; removed: number }>('/decks/recycle', { method: 'DELETE' }),
  buildDeck: (pages: { templateId: string; templateFolder?: string; templateVersion?: string; texts?: Record<string, string>; variables?: Record<string, string>; imageDataUrl?: string; tableData?: Record<string, string[][]>; tables?: Record<string, FitResult> }[]) =>
    request<{ ok: boolean; base64?: string; pageCount?: number; pageResults?: { index: number; ok: boolean; error?: string }[]; error?: string }>(
      '/decks/build', { method: 'POST', body: JSON.stringify({ pages }) }, 90000),

  // AI 自动模板分析（可选增强；失败返回 ok:false，前端回退规则分类）
  // 性能埋点批量上报（静默失败；后端记入环形缓冲 + 超预算 warn 日志）
  logPerf: (entries: { op: string; ms: number; meta?: Record<string, unknown> }[]) =>
    request<{ ok: boolean; recorded: number }>('/perf/log', { method: 'POST', body: JSON.stringify({ entries }) }, 5000),
  analyzeShapes: (shapes: { shapeId: string; name?: string; type?: string; fontSize?: number; bold?: boolean; text?: string; left?: number; top?: number; width?: number; height?: number; source?: string }[]) =>
    request<{ ok: boolean; recommendations?: { shapeId: string; recommendedRole: string; recommendedSemanticRole?: string; confidence: number; reason: string; suggestedPrompt?: string; suggestedConstraints?: Record<string, unknown> }[]; error?: string }>(
      '/analyze', { method: 'POST', body: JSON.stringify({ shapes }) }, 30000),
  readAllBytes: (payload: { bytes: Uint8Array; slideIndex?: number; needBackground?: boolean; shapes: { name: string; type?: string; bounds: { left: number; top: number; width: number; height: number }; textStyle?: TemplateShape['textStyle'] }[] }) => {
    const q = new URLSearchParams();
    if (payload.slideIndex !== undefined) q.set('slideIndex', String(payload.slideIndex));
    q.set('needBackground', payload.needBackground ? '1' : '0');
    const shapesJson = JSON.stringify(payload.shapes);
    const shapesBytes = new TextEncoder().encode(shapesJson);
    const head = new Uint8Array(4);
    new DataView(head.buffer).setUint32(0, shapesBytes.length, false);
    const body = new Uint8Array(4 + shapesBytes.length + payload.bytes.length);
    body.set(head, 0);
    body.set(shapesBytes, 4);
    body.set(payload.bytes, 4 + shapesBytes.length);
    return request<{ background?: TemplateBackground | null; styles: (TemplateShape['textStyle'] | null)[]; imageStyles?: (TemplateShape['imageStyle'] | null)[]; tables?: { name?: string; bounds: { left: number; top: number; width: number; height: number }; table: TableSpec }[] }>(
      '/slides/read-all?' + q.toString(),
      { method: 'POST', body: new Blob([body]), headers: { 'Content-Type': 'application/octet-stream' } },
      90000
    );
  }
};

// 解析「默认搜图供应商」：读取配置并校验该 provider 仍存在于注册表（内置或启用的自定义源）
// 校验失败/后端不可用 → 回退 baidu_page（保持旧行为）
let imageProvidersCache: Promise<string[]> | null = null;
export async function getDefaultImageProvider(): Promise<string> {
  try {
    const cfg = await Api.getConfig();
    const want = cfg.image?.provider;
    if (want) {
      const ids = await (imageProvidersCache ||= Api.getImageProviders()
        .then((r) => r.providers.map((p) => p.id))
        .catch(() => []));
      if (ids.includes(want)) return want;
    }
  } catch { /* 后端未就绪：使用默认 */ }
  return 'baidu_page';
}
