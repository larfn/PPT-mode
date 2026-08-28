// 图片裁剪几何计算（纯函数，可独立单测）
// 模型：图片以 transform: translate(tx,ty) scale(s) 显示在固定比例的裁剪框（stage）内，
//       s 为显示缩放（相对原图像素），tx/ty 为图片左上角在 stage 中的偏移（CSS px，可负）。
// 约束：图片必须完全覆盖裁剪框（cover），否则露白；s 有下限，tx/ty 有边界。

export interface CropState {
  naturalW: number; // 原图宽（px）
  naturalH: number; // 原图高（px）
  frameW: number;   // 裁剪框宽（显示 px）
  frameH: number;   // 裁剪框高（显示 px）
  scale: number;    // 显示缩放（>= coverScale）
  tx: number;       // 图片左上角在框内 X 偏移（px）
  ty: number;
}

// 覆盖裁剪框所需的最小缩放（不露白）
export function coverScale(naturalW: number, naturalH: number, frameW: number, frameH: number): number {
  if (!naturalW || !naturalH || !frameW || !frameH) return 1;
  return Math.max(frameW / naturalW, frameH / naturalH);
}

// 缩放图片到「恰好覆盖裁剪框」（适应裁剪框 / 重置）
export function fitCrop(state: CropState): CropState {
  const scale = coverScale(state.naturalW, state.naturalH, state.frameW, state.frameH);
  return {
    ...state,
    scale,
    tx: (state.frameW - state.naturalW * scale) / 2,
    ty: (state.frameH - state.naturalH * scale) / 2
  };
}

// 边界钳制：scale 不低于覆盖所需；tx/ty 保证图片覆盖整个裁剪框（不露白）
export function clampCrop(state: CropState): CropState {
  const minScale = coverScale(state.naturalW, state.naturalH, state.frameW, state.frameH);
  const scale = Math.max(minScale, state.scale);
  const imgW = state.naturalW * scale;
  const imgH = state.naturalH * scale;
  // tx 范围：图片左边缘 <= 框左边缘(0) 且 图片右边缘 >= 框右边缘(frameW)
  const txMin = state.frameW - imgW; // 图片左边缘最多移到 frameW - imgW（即右边缘贴框右）
  const txMax = 0;                    // 图片左边缘不能超过框左
  const tyMin = state.frameH - imgH;
  const tyMax = 0;
  const tx = Math.min(txMax, Math.max(txMin, state.tx));
  const ty = Math.min(tyMax, Math.max(tyMin, state.ty));
  return { ...state, scale, tx, ty };
}

// 从当前显示状态计算「原图上的裁剪矩形」（像素），用于 canvas drawImage 与校验
export function cropRectFromState(state: CropState): { sx: number; sy: number; sw: number; sh: number } {
  const c = clampCrop(state); // 先钳制，保证结果合法
  const sx = -c.tx / c.scale;
  const sy = -c.ty / c.scale;
  const sw = c.frameW / c.scale;
  const sh = c.frameH / c.scale;
  return { sx, sy, sw, sh };
}

// 根据目标宽高比计算裁剪框显示尺寸（受最大宽/高限制）
export function stageSizeForRatio(ratio: number, maxW: number, maxH: number): { w: number; h: number } {
  const r = ratio > 0 ? ratio : 1;
  let w = maxW;
  let h = w / r;
  if (h > maxH) {
    h = maxH;
    w = h * r;
  }
  return { w, h };
}

// 输出尺寸上限：裁剪区像素长边 cap，保持比例（防超大图内存爆炸）
export function outputSize(sw: number, sh: number, maxEdge = 2048): { w: number; h: number } {
  if (!sw || !sh) return { w: 0, h: 0 };
  const longest = Math.max(sw, sh);
  if (longest <= maxEdge) return { w: Math.round(sw), h: Math.round(sh) };
  const k = maxEdge / longest;
  return { w: Math.round(sw * k), h: Math.round(sh * k) };
}
