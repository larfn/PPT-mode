import { Api } from '../api.js';

const POINTS_PER_INCH = 72;

export interface HighlightConfig {
  color: string;
  durationMs: number; // 框高亮停留时长（0-500ms，不含淡出）
}

// 读取配置中的定位框颜色与显示时长；读取失败时退回默认（红色、0.5 秒）。
export async function getHighlightConfig(): Promise<HighlightConfig> {
  try {
    const cfg = await Api.getConfig();
    const h = (cfg && cfg.highlight) || {};
    const durationMs = typeof h.durationMs === 'number' ? Math.max(0, Math.min(500, h.durationMs)) : 500;
    return { color: h.color || '#FF0000', durationMs };
  } catch {
    return { color: '#FF0000', durationMs: 500 };
  }
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function setTransparency(shapeId: string, t: number): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const rect = slide.shapes.getItem(shapeId);
      rect.lineFormat.transparency = t;
      await context.sync();
    });
  } catch {
    // 形状已被删除或页面切换，中断淡入淡出
    throw new Error('shape-lost');
  }
}

// 在 PowerPoint 当前选中页上，用指定颜色画一个临时矩形框住该元素（激光笔效果）：
// 立即显示 → 停留 durationMs → 分步淡出 → 自动删除，不留痕迹。
export async function highlightShapeOnSlide(
  bounds: { left: number; top: number; width: number; height: number },
  color: string,
  durationMs = 500
): Promise<void> {
  let shapeId: string | null = null;
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const left = Math.max(bounds.left * POINTS_PER_INCH, 0);
      const top = Math.max(bounds.top * POINTS_PER_INCH, 0);
      const width = Math.max(bounds.width * POINTS_PER_INCH, 6);
      const height = Math.max(bounds.height * POINTS_PER_INCH, 6);
      const rect = slide.shapes.addGeometricShape('RoundRectangle', { left, top, width, height });
      rect.fill.clear();
      rect.lineFormat.color = color;
      rect.lineFormat.weight = 3.5;
      rect.lineFormat.transparency = 0;
      rect.load('id');
      await context.sync();
      shapeId = rect.id;
    });
  } catch {
    return; // 画框失败（如未选中页面）直接忽略
  }
  if (!shapeId) return;

  try {
    // 停留（时长可配置）
    await sleep(Math.max(0, Math.min(500, durationMs)));
    // 淡出（约 450ms）
    for (const t of [0.15, 0.3, 0.45, 0.6, 0.74, 0.86, 0.94, 1]) {
      await setTransparency(shapeId, t);
      await sleep(55);
    }
  } catch {
    // 淡入淡出中断（如切换页面）：仍尝试删除，避免残留空框
  }
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const rect = slide.shapes.getItem(shapeId!);
      rect.delete();
      await context.sync();
    });
  } catch {
    // 形状已不存在，忽略
  }
}
