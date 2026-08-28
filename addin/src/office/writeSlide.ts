import { Api } from '../api.js';
import type { TemplateDoc } from '../api.js';
import type { FitResult } from '../api.js'; // 表格适配结果类型（引擎输出）
import { perfAsync } from '../lib/perf.js';

export interface PendingSlide {
  templateId: string;
  folder: string; // 模板所在分类文件夹（可能为空 = 未分类）
  template?: TemplateDoc; // 本次生成用模板副本（如向导临时背景），不写回模板库
  // 每个 AI 图片位各自的图片 dataURL（多图）；兼容旧调用传 imageDataUrl 单值
  images?: Record<string, string>;
  imageDataUrl?: string;
  texts: Record<string, string>;
  vars: Record<string, string>;
  tableData?: Record<string, string[][]>; // 表格位数据（逐格编辑/粘贴 CSV/AI 生成）：{ [shapeId]: string[][] }
  tables?: Record<string, FitResult>; // 表格位适配结果（排版引擎 fit 输出）：{ [shapeId]: FitResult }，可选，缺省走旧路径
  clearSrcRectFor?: string[]; // 已人工裁剪的 ai_image 位 shape id：生成前清除其旧 srcRect（防双重裁剪）
}

export type WriteStage = '读取模板' | '生成页面' | '插入PPT';

export interface WriteOptions {
  onStage?: (stage: WriteStage) => void;
  insertAtStart?: boolean; // 兜底：不指定目标页，插入到文档开头
}

// 带类型的写入错误：kind 用于界面给出针对性操作（重试 / 插入到开头 / 复制错误 / 导出诊断）
export class WriteError extends Error {
  kind: 'selection' | 'insert' | 'timeout' | 'unsupported' | 'server' | 'unknown';
  detail: string;
  exportBase64?: string; // 插入失败时附上生成的文件，便于导出诊断
  constructor(kind: WriteError['kind'], message: string, detail = '', exportBase64?: string) {
    super(message);
    this.name = 'WriteError';
    this.kind = kind;
    this.detail = detail;
    this.exportBase64 = exportBase64;
  }
}

// 给耗时操作加超时：任何一步卡住都会变成可见的错误，而不是无限等待
function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${what}超时（超过 ${Math.round(ms / 1000)} 秒无响应）`)), ms);
  });
  return Promise.race([p, timeout]).finally(() => { if (timer) clearTimeout(timer); });
}

// 官方要求 targetSlideId 形如 "267#" / "#763315295" / "267#763315295"。
// 纯数字 ID（官方 getSelectedDataAsync 返回 nnn 幻灯片 ID）→ 追加 "#"；
// 若已含 "#" 或为 GUID 形式则原样/前缀处理，避免拼出非法格式。
function normalizeTargetId(id: string): string {
  const s = String(id || '').trim();
  if (!s) return '';
  if (s.includes('#')) return s;
  if (/^[0-9]+$/.test(s)) return `${s}#`;
  return `#${s}`;
}

// 获取当前选中的幻灯片：官方推荐 Common API（返回 nnn 数字 ID + 1-based index），
// 比 PowerPoint.getSelectedSlides().getItemAt(0).id 更可靠（后者可能返回创建 ID 导致 targetSlideId 非法）。
function getSelectedSlideInfo(): Promise<{ id: string; index: number }> {
  return new Promise((resolve, reject) => {
    const office = (window as unknown as { Office?: any }).Office;
    if (!office?.context?.document?.getSelectedDataAsync) {
      reject(new WriteError('selection', '当前 PowerPoint 不支持读取选中页。请先选中一页幻灯片后重试。'));
      return;
    }
    office.context.document.getSelectedDataAsync(office.CoercionType.SlideRange, (result: any) => {
      if (result.status === office.AsyncResultStatus.Failed) {
        reject(new WriteError('selection',
          `无法读取选中的幻灯片：${result.error?.message || '未知错误'}\n请先在左侧缩略图中点击选中一页后重试。`));
        return;
      }
      const slides = result.value?.slides;
      if (!slides || !slides.length) {
        reject(new WriteError('selection', '没有检测到选中的幻灯片页。请先在左侧缩略图中点击选中一页，然后重试。'));
        return;
      }
      resolve({ id: String(slides[0].id), index: Number(slides[0].index) });
    });
  });
}

function isInvalidTargetError(e: unknown): boolean {
  const err = e as { code?: string; message?: string };
  const msg = `${err?.message || ''} ${err?.code || ''}`.toLowerCase();
  return msg.includes('invalidargument') || msg.includes('slidenotfound');
}

function isPowerPointApiAvailable(set: string, major: number): boolean {
  const office = (window as unknown as { Office?: any }).Office;
  return !!office?.context?.requirements?.isSetSupported?.(set, major);
}

// 方案 A：带 targetSlideId 插入（官方格式 "267#"），插入到选中页之后
async function insertWithTarget(base64: string, targetId: string, formatting: PowerPoint.InsertSlideFormatting): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.insertSlidesFromBase64(base64, {
      formatting,
      targetSlideId: targetId
    });
    await context.sync();
  });
}

// 文件兼容性测试：PowerPoint.createPresentation 会在新窗口打开该文件，
// 若文件本身有效即可打开——用来区分「文件问题」还是「插入 API 问题」
export async function testFileOpenable(base64: string): Promise<void> {
  await PowerPoint.createPresentation(base64);
}

// 方案 B：不带 targetSlideId（插入到文档开头），再用 Slide.moveTo 定位到选中页之后
// selectedIndex1Based 为插入前选中页的 1-based 序号；插入后新页在 0-based 0，
// 选中页变为 0-based (index1Based)，新页 moveTo(index1Based) 即紧随其后。
async function insertAtStartAndMove(base64: string, selectedIndex1Based: number): Promise<void> {
  const canMove = isPowerPointApiAvailable('PowerPointApi', 1.8);
  await PowerPoint.run(async (context) => {
    const pres = context.presentation;
    pres.insertSlidesFromBase64(base64); // 裸调用：不传任何 options（默认 KeepSourceFormatting）
    await context.sync();
    if (canMove && Number.isFinite(selectedIndex1Based) && selectedIndex1Based > 0) {
      const inserted = pres.slides.getItemAt(0);
      inserted.moveTo(selectedIndex1Based); // 0-based 位置 = 原 1-based 序号
      await context.sync();
    }
  });
}

async function insertIntoPowerPoint(base64: string, insertAtStart: boolean): Promise<void> {
  const office = (window as unknown as { Office?: any }).Office;
  const supported = office?.context?.requirements?.isSetSupported?.('PowerPointApi', 1.2);
  if (supported === false) {
    throw new WriteError('unsupported',
      '当前 PowerPoint 不支持「插入幻灯片」功能（需要 PowerPointApi 1.2）。\n请升级 PowerPoint 后重试，或换一台电脑使用。');
  }
  // 尝试记录：把每次尝试的结果写进错误详情，便于定位
  const attempts: string[] = [];
  const record = (name: string, e: unknown) => {
    const err = e as { code?: string; message?: string };
    attempts.push(name + ' → ' + (err?.code || err?.message || String(e)).slice(0, 80));
  };
  try {
    if (insertAtStart) {
      await withTimeout(insertAtStartAndMove(base64, -1), 30000, '插入 PPT');
      return;
    }
    // 1) 官方方案：getSelectedDataAsync 拿选中页 ID（先试 UseDestinationTheme，再试 KeepSourceFormatting）
    const selected = await getSelectedSlideInfo();
    const targetId = normalizeTargetId(selected.id);
    const tryTarget = async (f: PowerPoint.InsertSlideFormatting) => {
      await withTimeout(insertWithTarget(base64, targetId, f), 30000, '插入 PPT');
    };
    try {
      await tryTarget(PowerPoint.InsertSlideFormatting.useDestinationTheme);
      return;
    } catch (e) {
      record('带目标页(UseDestinationTheme)', e);
      if (!isInvalidTargetError(e)) throw e;
      try {
        await tryTarget(PowerPoint.InsertSlideFormatting.keepSourceFormatting);
        return;
      } catch (e2) {
        record('带目标页(KeepSourceFormatting)', e2);
        if (!isInvalidTargetError(e2)) throw e2;
      }
    }
    // 2) 裸调用兜底：无任何 options + 移动定位
    try {
      await withTimeout(insertAtStartAndMove(base64, selected.index), 30000, '插入 PPT');
      return;
    } catch (e) {
      record('裸调用(开头+移动)', e);
      if (!isInvalidTargetError(e)) throw e;
    }
    throw new WriteError('insert',
      'PowerPoint 拒绝导入生成的页面（三种方式均报 InvalidArgument）。\n可点击「测试文件能否打开」判断是文件还是 API 的问题；或「导出诊断文件」手动拖入 PowerPoint 测试。',
      attempts.join('；'), base64);
  } catch (e) {
    if (e instanceof WriteError) throw e;
    const err = e as { code?: string; message?: string; debugInfo?: string; name?: string };
    const msg = `${err?.message || ''} ${err?.code || ''}`.toLowerCase();
    if (msg.includes('超时')) {
      throw new WriteError('timeout',
        '插入 PPT 超时（PowerPoint 长时间没有响应）。\n请稍等片刻观察演示文稿中是否已出现新页面；若没有，可重试，或完全退出并重新打开 PowerPoint。',
        `PowerPoint.run 超过 30 秒未完成`, base64);
    }
    if (msg.includes('itemnotfound') || msg.includes('index out of') || msg.includes('no slide') || msg.includes('selected')) {
      throw new WriteError('selection',
        '没有检测到选中的幻灯片页。请先在左侧缩略图中点击选中一页，然后重试。',
        `${err?.code || ''} ${err?.message || ''}`.trim());
    }
    if (isInvalidTargetError(e)) {
      // 两种方案都失败：大概率是生成的文件 PowerPoint 无法导入
      throw new WriteError('insert',
        'PowerPoint 无法导入生成的页面（InvalidArgument）。\n可点击「导出诊断文件」，把文件手动拖入 PowerPoint 测试导入；\n或返回上一步重新生成后重试；仍失败请完全退出并重开 PowerPoint。',
        `${err?.code || ''} ${err?.message || ''}`.trim(), base64);
    }
    throw new WriteError('insert',
      `PPT 插入失败${err?.code ? `（${err.code}）` : ''}：${err?.message || '未知错误'}\n可点击「重试」；若反复失败，请完全退出并重开 PowerPoint。`,
      err?.debugInfo || `${err?.code || ''} ${err?.message || ''}`.trim(), base64);
  }
}

// 直接插入已生成好的 base64 页面（AI 待写队列场景：MCP 生成 → 任务窗格写入）
export async function insertSlideBase64(base64: string, opts: { insertAtStart?: boolean } = {}): Promise<void> {
  await insertIntoPowerPoint(base64, opts.insertAtStart === true);
}

export async function insertTemporarySlideBase64(base64: string): Promise<string> {
  const selected = await getSelectedSlideInfo().catch(() => ({ id: '', index: -1 }));
  const canMove = isPowerPointApiAvailable('PowerPointApi', 1.8);
  return PowerPoint.run(async (context) => {
    const pres = context.presentation;
    pres.insertSlidesFromBase64(base64);
    await context.sync();
    const inserted = pres.slides.getItemAt(0);
    if (canMove && Number.isFinite(selected.index) && selected.index > 0) {
      inserted.moveTo(selected.index);
    }
    inserted.load('id');
    await context.sync();
    return inserted.id;
  });
}

export async function deleteSlideById(slideId: string): Promise<void> {
  if (!slideId) return;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemOrNullObject(slideId);
    await context.sync();
    if (slide.isNullObject) return;
    slide.delete();
    await context.sync();
  });
}

export async function writePendingSlide(p: PendingSlide, opts: WriteOptions = {}): Promise<void> {
  const { onStage, insertAtStart } = opts;
  try {
    onStage?.('读取模板');
    let template;
    try {
      if (p.template) {
        template = JSON.parse(JSON.stringify(p.template));
      } else {
        const res = await withTimeout(Api.getTemplate(p.templateId, p.folder || ''), 15000, '读取模板');
        template = res.template;
      }
    } catch (e) {
      throw new WriteError('server', `读取模板失败：${(e as Error).message}`, String((e as Error)?.stack || ''));
    }
    onStage?.('生成页面');
    let base64 = '';
    try {
      // 人工裁剪过的图片位：图片本身已是裁剪结果，清除模板旧 srcRect 防双重裁剪（圆角/阴影保留）
      if (p.clearSrcRectFor?.length) {
        template = JSON.parse(JSON.stringify(template));
        for (const s of template.shapes) {
          if (p.clearSrcRectFor.includes(s.id) && s.imageStyle) delete s.imageStyle.srcRect;
        }
      }
      // 兼容兜底：若未显式传 imageDataUrl，取第一个图片位作为单图兜底（旧后端只认 imageDataUrl 时至少显示第一张图）
      const fallbackImage = p.imageDataUrl || (p.images ? (Object.values(p.images)[0] || '') : '');
      const built = await withTimeout(perfAsync('slideBuild', Api.buildSlideBase64({
        template,
        images: p.images,
        imageDataUrl: fallbackImage,
        texts: p.texts,
        vars: p.vars,
        tableData: p.tableData,
        tables: p.tables
      })), 60000, '生成页面');
      base64 = built.base64;
    } catch (e) {
      throw new WriteError('server', `生成页面失败：${(e as Error).message}\n可返回上一步重新生成后重试。`, String((e as Error)?.stack || ''));
    }
    onStage?.('插入PPT');
    await perfAsync('slideInsert', insertIntoPowerPoint(base64, insertAtStart === true), { sizeKB: Math.round(base64.length / 1024) });
  } catch (e) {
    if (e instanceof WriteError) throw e;
    throw new WriteError('unknown', `写入失败：${(e as Error)?.message || String(e)}`, String((e as Error)?.stack || ''));
  }
}
