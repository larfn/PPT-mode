import { Api } from '../api.js';
import { perfRecord } from '../lib/perf.js';

export interface ShapeInfo {
  id: string; type: string; name: string;
  source?: 'slide' | 'layout' | 'master';
  // 占位符类型（Office.js placeholderFormat.type 归一化）：title/ctrTitle/subTitle/body/pic/tbl/sldNum/dt/ftr/hdr…
  // 角色分析的最高优先级信号（Office 官方语义），无占位符时为 undefined
  phType?: string;
  bounds: { left: number; top: number; width: number; height: number };
  hasText: boolean; text?: string;
  textStyle?: {
    font?: string; size?: number; bold?: boolean; italic?: boolean; color?: string; align?: string;
    valign?: string; underline?: boolean; strikethrough?: boolean; doubleStrikethrough?: boolean;
    subscript?: boolean; superscript?: boolean;
    margin?: { top: number; right: number; bottom: number; left: number };
    autoFit?: string; wordWrap?: boolean;
    eaFont?: string; // 东亚字体（XML 回读，中文实际显示字体；latin 存 font）
  };
  fill?: { type?: string; color?: string };
  line?: { color?: string; weight?: number; visible?: boolean };
  rotation?: number;
  shapeType?: string;
  // 表格位（XML 回读补齐）：Office.js 读不到 GraphicFrame，由 read-all 的 tables 合并
  table?: { rows: number; cols: number; colWidths?: number[]; rowHeights?: number[]; tblPr?: { firstRow?: boolean; bandRow?: boolean }; cells: { row: number; col: number; rowspan?: number; colspan?: number; text?: string; textStyle?: Record<string, unknown>; fill?: string; valign?: string }[] };
  // 图片样式（XML 回读补齐）：几何形状（如 roundRect）/ 柔化边缘（EMU）/ 裁剪（千分比 srcRect）
  imageStyle?: {
    shape?: string;
    softEdgeEmu?: number;
    srcRect?: { l: number; t: number; r: number; b: number };
    spPrXml?: string; // 完整图片样式：spPr 除 xfrm 的原样 XML（边框/阴影/反射/辉光/柔化边缘/棱台/三维旋转/形状）
    blipAttrs?: string; // a:blip 除 r:embed 的属性（透明度等）
    blipKids?: string; // a:blip 子元素（图片更正/艺术效果等）
    imageDataUrl?: string; // 图片本体（dataURL）：保存模板时随固定图片位写入 shape.content，生成端原样渲染
  };
}

export interface SlideBackgroundInfo {
  type: 'none' | 'solid' | 'picture' | 'gradient' | 'pattern' | 'unsupported';
  source?: 'slide' | 'layout' | 'master';
  color?: string;
  imageDataUrl?: string;
}

export interface ReadSlideResult {
  slideSize: { width: number; height: number };
  background: SlideBackgroundInfo | null;
  shapes: ShapeInfo[];
  unsupported: string[];
}

// PowerPoint.js 的形状位置属性单位是磅（points），模板按设计文档统一存英寸（pt/72）。
const POINTS_PER_INCH = 72;

// textFrame 对无文本形状（图片/线条/表格等）会抛 InvalidArgument，因此只探测
// 明确支持文本的形状类型；其余类型按形状类型安全跳过。
const TEXT_SUPPORTING_TYPES = new Set(['TextBox', 'Placeholder', 'GeometricShape', 'Callout']);

// 生成阶段无法重建的元素类型（表格/图表/SmartArt 等），保存时提示用户。
const NON_REBUILDABLE_TYPES = new Set(['Table', 'Chart', 'SmartArt', 'Diagram', 'Media', 'Ole', 'ContentApp', 'Ink', 'Model3D', 'Graphic']);

const SOURCE_LABELS: Record<string, string> = { slide: '页面', layout: '版式', master: '母版' };

function toShapeType(type: string): string {
  if (type === 'TextBox') return 'text';
  if (type === 'Image') return 'picture';
  if (type === 'Line') return 'line';
  if (type === 'GeometricShape') return 'rectangle';
  return 'other';
}

// horizontalAlignment 在个别 PowerPoint 客户端版本返回非字符串值（可能是数字枚举或枚举对象），统一归一化为小写字符串。
function normalizeAlign(value: unknown): string | undefined {
  if (typeof value === 'number') {
    // 部分 Office 客户端把水平对齐枚举返回为数字（Word 兼容序号）：0=左 1=中 2=右 3=两端对齐 4/5=分散对齐
    const numeric: Record<number, string> = { 0: 'left', 1: 'center', 2: 'right', 3: 'justify', 4: 'justify', 5: 'justify' };
    return numeric[value];
  }
  let v: string | undefined;
  if (typeof value === 'string') {
    v = value.toLowerCase();
  } else if (value && typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    for (const key of ['Left', 'Center', 'Right', 'Justify', 'Distributed', 'ThaiDistributed']) {
      if (obj[key] === key || obj[key] === key.toLowerCase()) { v = key.toLowerCase(); break; }
    }
  }
  if (!v) return undefined;
  if (v === 'distributed' || v === 'thaiddistributed' || v === 'justified') return 'justify';
  if (v === 'centered') return 'center';
  return v; // left / center / right / justify
}

// Office.js placeholderFormat.type（如 'Title'/'Body'/'Picture'/'SlideNumber'/'DateAndTime'…）归一化为
// 统一的 OOXML 风格命名（与 server/src/readStyles.js 的 p:ph type 命名一致，供角色分析使用）
function normalizePhType(value: unknown): string | undefined {
  if (typeof value !== 'string' || !value.trim()) return undefined;
  const map: Record<string, string> = {
    title: 'title', centertitle: 'ctrTitle', subtitle: 'subTitle', body: 'body',
    verticalbody: 'vertBody', picture: 'pic', table: 'tbl', chart: 'chart',
    diagram: 'dgm', media: 'media', clipart: 'clipArt', slidenumber: 'sldNum',
    dateandtime: 'dt', footer: 'ftr', header: 'hdr', object: 'obj',
    content: 'content', contentvertical: 'vertContent', verticalobject: 'vertObj', orgchart: 'orgChart'
  };
  return map[value.trim().toLowerCase()];
}

function normalizeVAlign(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const v = value.toLowerCase();
  if (v === 'middle' || v === 'middlecentered') return 'middle';
  if (v === 'bottom' || v === 'bottomcentered') return 'bottom';
  return 'top';
}

function normalizeAutoFit(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const v = value.toLowerCase();
  if (v.includes('texttofit') || v.includes('shrink')) return 'shrink';
  if (v.includes('shapetofit') || v.includes('resize')) return 'resize';
  return 'none';
}

// Office.js 无法读取具体几何类型（箭头/椭圆/菱形等），只能读到 GeometricShape。
// 根据默认形状名（如「右箭头 5」「Oval 6」）猜测最可能的类型，供保存界面下拉选择。
// 连接符（如「直接箭头连接符 20」）按宽高判断方向：竖直→下箭头、水平→右箭头。
function guessShapeType(name: string, officeType: string, bounds?: { width: number; height: number }): string {
  const n = name || '';
  const isVertical = bounds ? bounds.height > bounds.width + 0.001 : false;
  if (/下箭头|down ?arrow/i.test(n)) return 'downArrow';
  if (/上箭头|up ?arrow/i.test(n)) return 'upArrow';
  if (/左箭头|left ?arrow/i.test(n)) return 'leftArrow';
  if (/右箭头|right ?arrow/i.test(n)) return 'rightArrow';
  if (/左右箭头|双箭|双向|left.?right|leftright/i.test(n)) return 'leftRightArrow';
  if (/上下箭头|up.?down/i.test(n)) return 'upDownArrow';
  if (officeType === 'Line') {
    if (/箭头|arrow|连接符|connector/i.test(n)) return isVertical ? 'downArrow' : 'rightArrow';
    return 'line';
  }
  if (/箭头|arrow/i.test(n)) return isVertical ? 'downArrow' : 'rightArrow';
  if (/圆角|rounded/i.test(n)) return 'roundRect';
  if (/椭圆|ellipse|oval|圆形|circle|^圆/i.test(n)) return 'ellipse';
  if (/菱形|diamond/i.test(n)) return 'diamond';
  if (/三角|triangle/i.test(n)) return 'triangle';
  if (/五边|pentagon/i.test(n)) return 'pentagon';
  if (/六边|hexagon/i.test(n)) return 'hexagon';
  if (/梯形|trapezoid/i.test(n)) return 'trapezoid';
  if (/平行四边|parallelogram/i.test(n)) return 'parallelogram';
  if (/燕尾|chevron/i.test(n)) return 'chevron';
  if (/心形|heart/i.test(n)) return 'heart';
  if (/星|star/i.test(n)) return 'star5';
  if (/云|cloud/i.test(n)) return 'cloud';
  if (/环形|donut|annulus/i.test(n)) return 'donut';
  if (/弧形|blockarc/i.test(n)) return 'blockArc';
  return 'rect';
}

interface PendingShape {
  shape: PowerPoint.Shape;
  offsetX: number;
  offsetY: number;
  source: 'slide' | 'layout' | 'master';
  textFrame: PowerPoint.TextFrame | null;
}

export async function readCurrentSlide(onProgress?: (phase: string) => void): Promise<ReadSlideResult> {
  const t0 = performance.now();
  try {
  return await PowerPoint.run(async (context) => {
    let phase = '选中幻灯片';
    const setPhase = (p: string) => { phase = p; onProgress?.(p); };
    try {
      const pageSetup = context.presentation.pageSetup;
      pageSetup.load(['slideWidth', 'slideHeight']);
      const slides = context.presentation.getSelectedSlides();
      slides.load('items');
      await context.sync();
      if (!slides.items.length) {
        throw new Error('未检测到选中的幻灯片，请先在左侧幻灯片缩略图中点击选中一页');
      }
      const slide = slides.getItemAt(0);
      setPhase('加载形状列表');
      const slideW = pageSetup.slideWidth / POINTS_PER_INCH;
      const slideH = pageSetup.slideHeight / POINTS_PER_INCH;

      const pending: PendingShape[] = [];
      const unsupported: string[] = [];

      // 递归收集形状；组合形状（group）展开为独立元素，子元素坐标按相对 group 偏移累加。
      const collect = async (shape: PowerPoint.Shape, offsetX: number, offsetY: number, source: 'slide' | 'layout' | 'master'): Promise<void> => {
        const type = String(shape.type);
        if (type === 'Group') {
          const group = shape.group;
          group.shapes.load('items');
          await context.sync();
          for (const child of group.shapes.items) {
            child.load(['id', 'name', 'type', 'left', 'top', 'width', 'height', 'rotation']);
            child.fill.load(['type', 'foregroundColor']);
            child.lineFormat.load(['color', 'weight', 'visible']);
            if (String(child.type) === 'Placeholder') child.placeholderFormat.load('type');
          }
          await context.sync();
          for (const child of group.shapes.items) {
            await collect(child, offsetX + shape.left, offsetY + shape.top, source);
          }
          return;
        }
        if (NON_REBUILDABLE_TYPES.has(type)) {
          unsupported.push(`${shape.name || type}（${type}，位于${SOURCE_LABELS[source] || source}）`);
        }
        let textFrame: PowerPoint.TextFrame | null = null;
        if (TEXT_SUPPORTING_TYPES.has(type)) {
          try {
            // isNullObject is auto-filled after sync; loading it throws InvalidArgument.
            textFrame = shape.textFrame;
          } catch (e) {
            unsupported.push(`${shape.name || type}（文本框架读取失败，位于${SOURCE_LABELS[source] || source}）`);
          }
        }
        pending.push({ shape, offsetX, offsetY, source, textFrame });
      };

      // 从形状集合批量加载基础属性并收集。skipPlaceholderTypes 用于跳过与页面实例重复的占位符。
      const collectFrom = async (
        shapes: PowerPoint.ShapeCollection,
        source: 'slide' | 'layout' | 'master',
        skipPlaceholderTypes: Set<string> | null,
        trackPlaceholderTypes?: Set<string>
      ): Promise<void> => {
        shapes.load('items');
        await context.sync();
        for (const shape of shapes.items) {
          shape.load(['id', 'name', 'type', 'left', 'top', 'width', 'height', 'rotation']);
          shape.fill.load(['type', 'foregroundColor']);
          shape.lineFormat.load(['color', 'weight', 'visible']);
          if (String(shape.type) === 'Placeholder') shape.placeholderFormat.load('type');
        }
        await context.sync();
        for (const shape of shapes.items) {
          const type = String(shape.type);
          if (type === 'Placeholder') {
            const pt = shape.placeholderFormat ? String(shape.placeholderFormat.type || '') : '';
            if (trackPlaceholderTypes && pt) trackPlaceholderTypes.add(pt);
            if (skipPlaceholderTypes && skipPlaceholderTypes.has(pt)) continue;
          }
          await collect(shape, 0, 0, source);
        }
      };



      // 1) 页面形状（含占位符实例），同时记录页面上已有的占位符类型
      setPhase('加载页面形状');
      const pagePlaceholderTypes = new Set<string>();
      await collectFrom(slide.shapes, 'slide', null, pagePlaceholderTypes);

      // 2) 版式形状（模板自带的装饰文本框等；页面已实例化的占位符跳过，避免重复）
      setPhase('加载版式形状');
      try {
        await collectFrom(slide.layout.shapes, 'layout', pagePlaceholderTypes);
      } catch (e) {
        unsupported.push(`版式元素读取失败：${(e as Error)?.message || String(e)}`);
      }

      // 3) 母版形状（Slide 直接提供 slideMaster 导航属性，直达该页版式所属母版）
      setPhase('加载母版形状');
      try {
        const master = slide.slideMaster;
        await collectFrom(master.shapes, 'master', pagePlaceholderTypes);
      } catch (e) {
        unsupported.push(`母版元素读取失败：${(e as Error)?.message || String(e)}`);
      }

      setPhase('加载文本内容与样式');
      // textFrame.isNullObject 是自动填充属性，需在读取前先 sync 一次
      // （改为 slide.slideMaster 直连后，母版查找中的 sync 已不存在，必须补上）。
      await context.sync();
      // 文本样式批量加载：每批一次 sync（Office.js 往返开销大，逐形状 sync 是读取卡顿主因）。
      // 批内失败时回退逐形状读取，保留单个文本框失败不影响其他元素的容错。
      const textTargets = pending.filter((p) => p.textFrame && !p.textFrame.isNullObject);
      const textEntries: { shape: PowerPoint.Shape; textRange: PowerPoint.TextRange; font: PowerPoint.ShapeFont; paragraph: PowerPoint.ParagraphFormat; frame: PowerPoint.TextFrame }[] = [];
      const BATCH_SIZE = 8;
      const loadOneText = async (p: PendingShape): Promise<boolean> => {
        const textRange = p.textFrame!.textRange;
        textRange.load(['text']);
        textRange.font.load(['name', 'size', 'bold', 'italic', 'color', 'underline', 'strikethrough', 'doubleStrikethrough', 'subscript', 'superscript']);
        textRange.paragraphFormat.load('horizontalAlignment');
        p.textFrame!.load(['verticalAlignment', 'leftMargin', 'rightMargin', 'topMargin', 'bottomMargin', 'autoSizeSetting', 'wordWrap']);
        await context.sync();
        if (textRange.text) {
          textEntries.push({ shape: p.shape, textRange, font: textRange.font, paragraph: textRange.paragraphFormat, frame: p.textFrame! });
        }
        return true;
      };
      for (let i = 0; i < textTargets.length; i += BATCH_SIZE) {
        const batch = textTargets.slice(i, i + BATCH_SIZE);
        try {
          for (const p of batch) {
            const textRange = p.textFrame!.textRange;
            textRange.load(['text']);
            textRange.font.load(['name', 'size', 'bold', 'italic', 'color', 'underline', 'strikethrough', 'doubleStrikethrough', 'subscript', 'superscript']);
            textRange.paragraphFormat.load('horizontalAlignment');
            p.textFrame!.load(['verticalAlignment', 'leftMargin', 'rightMargin', 'topMargin', 'bottomMargin', 'autoSizeSetting', 'wordWrap']);
          }
          await context.sync();
          for (const p of batch) {
            if (p.textFrame!.textRange.text) {
              textEntries.push({ shape: p.shape, textRange: p.textFrame!.textRange, font: p.textFrame!.textRange.font, paragraph: p.textFrame!.textRange.paragraphFormat, frame: p.textFrame! });
            }
          }
        } catch {
          for (const p of batch) {
            try {
              await loadOneText(p);
            } catch (e) {
              unsupported.push(`${p.shape.name || '文本框'}（文本样式读取失败，位于${SOURCE_LABELS[p.source] || p.source}：${(e as Error)?.message || String(e)}）`);
            }
          }
        }
      }

      setPhase('读取页面背景');
      // 背景默认跟随文档默认背景：优先当前页背景；页面未单独设置时回退到版式、再回退到母版。
      const readFill = async (fill: PowerPoint.SlideBackgroundFill, source: 'slide' | 'layout' | 'master'): Promise<SlideBackgroundInfo | null> => {
        fill.load('type');
        await context.sync();
        const fillType = String(fill.type);
        if (fillType === 'Solid') {
          const solid = fill.getSolidFillOrNullObject();
          solid.load('color');
          await context.sync();
          if (!solid.isNullObject && solid.color) return { type: 'solid', source, color: solid.color };
          return null;
        }
        if (fillType === 'PictureOrTexture') {
          // 注意：不读取 imageBase64 —— 该 API 会把整张背景图转成 base64 传回（大图很慢且占内存），
          // 且模板背景现在由「沿用默认 / 自选图片」决定，不再需要这份数据。
          return { type: 'picture', source };
        }
        if (fillType === 'Gradient') return { type: 'gradient', source };
        if (fillType === 'Pattern') return { type: 'pattern', source };
        return null; // none / unsupported
      };
      let background: SlideBackgroundInfo | null = null;
      try {
        background = await readFill(slide.background.fill, 'slide');
        if (!background) {
          const layout = slide.layout;
          background = await readFill(layout.background.fill, 'layout');
        }
        if (!background) {
          const masters = context.presentation.slideMasters;
          masters.load('items');
          await context.sync();
          for (const master of masters.items) {
            background = await readFill(master.background.fill, 'master');
            if (background) break;
          }
        }
      } catch {
        background = null;
      }

      const rawDebug: Record<string, unknown>[] = [];
      const shapesOut = pending.map((p) => {
        const { shape, offsetX, offsetY, source, textFrame } = p;
        const entry = textEntries.find((t) => t.shape === shape);
        const type = String(shape.type);
        const fillType = String(shape.fill?.type || '');
        const info: ShapeInfo = {
          id: String(shape.id),
          type: toShapeType(type),
          name: shape.name || '',
          source,
          bounds: {
            left: (shape.left + offsetX) / POINTS_PER_INCH,
            top: (shape.top + offsetY) / POINTS_PER_INCH,
            width: shape.width / POINTS_PER_INCH,
            height: shape.height / POINTS_PER_INCH
          },
          hasText: !!(entry && entry.textRange.text),
          phType: type === 'Placeholder' && shape.placeholderFormat
            ? normalizePhType((shape.placeholderFormat as { type?: unknown }).type)
            : undefined,
          rotation: shape.rotation || undefined,
          fill: fillType && fillType !== 'NoFill' ? { type: fillType, color: shape.fill.foregroundColor || undefined } : undefined,
          line: shape.lineFormat?.visible ? { color: shape.lineFormat.color || undefined, weight: shape.lineFormat.weight ?? undefined, visible: true } : undefined
        };
        if (type === 'GeometricShape' || type === 'Line') {
          info.shapeType = guessShapeType(shape.name || '', type, { width: shape.width, height: shape.height });
        }
        if (entry && entry.textRange.text) {
          rawDebug.push({
            name: info.name, type: info.type, source: info.source,
            rawUnderline: entry.font.underline, rawAlign: entry.paragraph.horizontalAlignment,
            rawStrike: entry.font.strikethrough, rawDoubleStrike: entry.font.doubleStrikethrough,
            rawSup: entry.font.superscript, rawSub: entry.font.subscript,
            rawFont: entry.font.name, rawSize: entry.font.size, rawBold: entry.font.bold, rawItalic: entry.font.italic, rawColor: entry.font.color
          });
          const frame = entry.frame;
          info.text = entry.textRange.text;
          info.textStyle = {
            font: entry.font.name ?? undefined,
            size: entry.font.size ?? undefined,
            bold: entry.font.bold ?? undefined,
            italic: entry.font.italic ?? undefined,
            color: entry.font.color ?? undefined,
            underline: !!entry.font.underline && entry.font.underline !== 'None' ? true : undefined,
            strikethrough: entry.font.strikethrough ?? undefined,
            doubleStrikethrough: entry.font.doubleStrikethrough ?? undefined,
            subscript: entry.font.subscript ?? undefined,
            superscript: entry.font.superscript ?? undefined,
            align: normalizeAlign(entry.paragraph.horizontalAlignment),
            valign: normalizeVAlign(frame?.verticalAlignment),
            margin: frame ? {
              left: (frame.leftMargin ?? 0) / POINTS_PER_INCH,
              right: (frame.rightMargin ?? 0) / POINTS_PER_INCH,
              top: (frame.topMargin ?? 0) / POINTS_PER_INCH,
              bottom: (frame.bottomMargin ?? 0) / POINTS_PER_INCH
            } : undefined,
            autoFit: frame ? normalizeAutoFit(frame.autoSizeSetting) : undefined,
            wordWrap: frame ? !!frame.wordWrap : undefined
          };
        }
        return info;
      });

      // 调试日志：记录读取阶段的原始样式值，便于排查属性保存问题（失败不影响主流程）
      void Api.debugRead({ slideW, slideH, shapes: rawDebug }).catch(() => {});

      // 读取顺序稳定化：同一来源组内按「上到下、左到右」排序（Office.js 返回的是 z-order，
      // 可能出现竖向 1、3、2 的乱序）；跨来源仍保持 页面 → 版式 → 母版 的分组顺序。
      const sourceRank: Record<string, number> = { slide: 0, layout: 1, master: 2 };
      shapesOut.sort((a, b) => {
        const ra = sourceRank[a.source || 'slide'] ?? 0;
        const rb = sourceRank[b.source || 'slide'] ?? 0;
        if (ra !== rb) return ra - rb;
        const da = a.bounds?.top ?? 0, db = b.bounds?.top ?? 0;
        if (Math.abs(da - db) > 0.001) return da - db;
        return (a.bounds?.left ?? 0) - (b.bounds?.left ?? 0);
      });

      return {
        slideSize: { width: slideW, height: slideH },
        background,
        shapes: shapesOut,
        unsupported
      };
    } catch (e) {
      const msg = (e as Error)?.message || String(e);
      throw new Error(`[${phase}] ${msg}`);
    }
  });
  } finally {
    perfRecord('pageRead', performance.now() - t0);
  }
}
