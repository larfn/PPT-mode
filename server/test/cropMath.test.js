// 图片裁剪几何计算单测（Node 24 type-stripping 直接 require 前端 TS 纯函数）
const { test } = require('node:test');
const assert = require('node:assert');
const {
  coverScale, fitCrop, clampCrop, cropRectFromState, stageSizeForRatio, outputSize
} = require('../../addin/src/lib/cropMath.ts');

test('coverScale: 取宽/高覆盖所需的最大值（不露白）', () => {
  // 框 320x240（4:3），图 1920x1080（16:9）
  assert.ok(Math.abs(coverScale(1920, 1080, 320, 240) - 240 / 1080) < 1e-9, '高度决定缩放');
  // 图比框窄/矮时返回 1 的等效（scale >= 1 不放大缩小于1）
  assert.ok(coverScale(100, 100, 300, 200) >= 3);
  assert.equal(coverScale(0, 100, 300, 200), 1, '非法尺寸返回 1');
  assert.equal(coverScale(100, 100, 0, 200), 1);
});

test('fitCrop: 恰好覆盖且居中，不露白', () => {
  const s = fitCrop({ naturalW: 1920, naturalH: 1080, frameW: 320, frameH: 240, scale: 1, tx: 0, ty: 0 });
  const imgW = 1920 * s.scale, imgH = 1080 * s.scale;
  assert.ok(imgW >= 320 && imgH >= 240, '覆盖裁剪框');
  // 居中：左右留白对称（tx 为负，图片左边缘在框左外）
  assert.ok(Math.abs(s.tx + (imgW - 320) / 2) < 1e-6, '水平居中');
  assert.ok(Math.abs(s.ty + (imgH - 240) / 2) < 1e-6, '垂直居中');
});

test('cropRectFromState: 比例与裁剪框一致、无变形、不越界', () => {
  const base = { naturalW: 1920, naturalH: 1080, frameW: 320, frameH: 240, scale: 1, tx: 0, ty: 0 };
  const s = fitCrop(base);
  const r = cropRectFromState(s);
  assert.ok(Math.abs(r.sw / r.sh - 320 / 240) < 1e-6, '裁剪比例 = 模板图片位比例（无变形）');
  assert.ok(r.sx >= 0 && r.sy >= 0, '源坐标非负');
  assert.ok(r.sx + r.sw <= 1920 + 1e-6, '不越出图片右边界');
  assert.ok(r.sy + r.sh <= 1080 + 1e-6, '不越出图片下边界');
  // 裁剪面积 = 框/scale^2
  assert.ok(Math.abs(r.sw * r.sh - (320 * 240) / (s.scale * s.scale)) < 1e-3);
});

test('clampCrop: 拖动越界被钳制（图片必须覆盖裁剪框，不露白）', () => {
  const base = fitCrop({ naturalW: 1920, naturalH: 1080, frameW: 320, frameH: 240, scale: 1, tx: 0, ty: 0 }); // 精确 cover 状态
  // 拖出右边界（tx=100 > 0）→ 钳到 0
  let c = clampCrop({ ...base, tx: 100, ty: 0 });
  assert.equal(c.tx, 0);
  // 拖出左边界（tx 小于最小值 frameW-imgW）→ 钳到最小值
  const imgW = 1920 * base.scale;
  c = clampCrop({ ...base, tx: -9999, ty: 0 });
  assert.ok(Math.abs(c.tx - (320 - imgW)) < 1e-6, 'tx 钳到最左');
  // 垂直同理
  const imgH = 1080 * base.scale;
  c = clampCrop({ ...base, tx: 0, ty: 9999 });
  assert.equal(c.ty, 0);
  c = clampCrop({ ...base, tx: 0, ty: -9999 });
  assert.ok(Math.abs(c.ty - (240 - imgH)) < 1e-6);
  // 钳制后仍不露白
  const r = cropRectFromState(c);
  assert.ok(r.sx >= 0 && r.sy >= 0 && r.sx + r.sw <= 1920 + 1e-6 && r.sy + r.sh <= 1080 + 1e-6);
});

test('clampCrop: 缩放不能低于覆盖所需（缩太小时自动回弹）', () => {
  const base = { naturalW: 1920, naturalH: 1080, frameW: 320, frameH: 240, scale: 0.01, tx: -50, ty: -50 };
  const c = clampCrop(base);
  assert.ok(c.scale >= coverScale(1920, 1080, 320, 240), 'scale 不小于 coverScale');
  const r = cropRectFromState(c);
  assert.ok(r.sw * r.sh > 0 && r.sx + r.sw <= 1920 + 1e-6);
});

test('极端宽高比：极宽图与极窄图都能 cover 且比例保持', () => {
  // 极宽图 10000x100（100:1）
  const s1 = fitCrop({ naturalW: 10000, naturalH: 100, frameW: 400, frameH: 300, scale: 1, tx: 0, ty: 0 });
  const r1 = cropRectFromState(s1);
  assert.ok(Math.abs(r1.sw / r1.sh - 400 / 300) < 1e-6);
  assert.ok(r1.sx >= 0 && r1.sx + r1.sw <= 10000 + 1e-6);
  // 极窄图 100x10000
  const s2 = fitCrop({ naturalW: 100, naturalH: 10000, frameW: 400, frameH: 300, scale: 1, tx: 0, ty: 0 });
  const r2 = cropRectFromState(s2);
  assert.ok(Math.abs(r2.sw / r2.sh - 400 / 300) < 1e-6);
  assert.ok(r2.sy >= 0 && r2.sy + r2.sh <= 10000 + 1e-6);
});

test('cropRectFromState: 用户缩放后裁剪区域相应缩小（放大图片=看到更少内容）', () => {
  const base = { naturalW: 1920, naturalH: 1080, frameW: 320, frameH: 240, scale: 1, tx: 0, ty: 0 };
  const fit = fitCrop(base);
  const rFit = cropRectFromState(fit);
  const zoomed = clampCrop({ ...fit, scale: fit.scale * 2 });
  const rZoom = cropRectFromState(zoomed);
  assert.ok(rZoom.sw < rFit.sw, '放大后裁剪区更小（看到更少内容）');
  assert.ok(Math.abs(rZoom.sw / rZoom.sh - rFit.sw / rFit.sh) < 1e-6, '比例始终不变');
});

test('stageSizeForRatio: 受最大宽高限制，比例保持', () => {
  const s1 = stageSizeForRatio(16 / 9, 520, 340);
  assert.ok(Math.abs(s1.w / s1.h - 16 / 9) < 1e-6);
  assert.ok(s1.w <= 520 && s1.h <= 340);
  const s2 = stageSizeForRatio(1, 520, 340);
  assert.equal(s2.w, 340); assert.equal(s2.h, 340); // 高受限
  const s3 = stageSizeForRatio(3 / 4, 520, 340);
  assert.ok(s3.w <= 520 && s3.h <= 340);
  assert.ok(Math.abs(s3.w / s3.h - 0.75) < 1e-6);
  const s0 = stageSizeForRatio(0, 520, 340);
  assert.equal(s0.w, 340, '非法比例回退 1:1，受最大高限制'); assert.equal(s0.h, 340);
});

test('outputSize: 超长边按比例缩小，短图原样', () => {
  const big = outputSize(4000, 3000, 2048);
  assert.equal(big.w, 2048);
  assert.ok(Math.abs(big.w / big.h - 4000 / 3000) < 1e-6, '比例保持');
  assert.ok(big.w <= 2048 && big.h <= 2048);
  const small = outputSize(800, 600, 2048);
  assert.deepEqual(small, { w: 800, h: 600 });
  assert.deepEqual(outputSize(0, 0), { w: 0, h: 0 });
  // 极宽裁剪区
  const wide = outputSize(10000, 100, 2048);
  assert.equal(wide.w, 2048);
  assert.ok(Math.abs(wide.w / wide.h - 100) < 6, 'Math.round 舍入容差内保持近似比例：' + wide.w + '/' + wide.h);
});
