// 模板图片位人工裁剪编辑器（纯前端，canvas 裁剪）
//  - 裁剪框比例固定为模板图片位宽高比；用户拖动/缩放的是图片显示区域
//  - 确定：把选中区域裁成新图（走原生成链路，imageStyle 圆角/阴影照常应用）
//  - 取消：不改动调用方已有的图片选择结果
import {
  CropState, fitCrop, clampCrop, cropRectFromState, stageSizeForRatio, outputSize, coverScale
} from './cropMath.js';

export interface CropEditorOptions {
  imageDataUrl: string;
  frameRatio: number;      // 模板图片位宽高比（width/height，英寸比 = 像素比）
  frameSizeLabel: string;  // 如「模板图片位：320 × 180 px」
  maxFrameW?: number;      // 显示最大宽（默认 520）
  maxFrameH?: number;      // 显示最大高（默认 340）
}

export interface CropEditorResult {
  canceled: boolean;
  dataUrl?: string;        // 裁剪后图片（canceled=false 时）
}

export function openImageCropEditor(opts: CropEditorOptions): Promise<CropEditorResult> {
  return new Promise((resolve) => {
    // 加载项窗口较小时（相对 PPT 页面比例小）自动收缩裁剪框，避免图片/弹窗超出可视区
    const vw = window.innerWidth || document.documentElement.clientWidth || 800;
    const vh = window.innerHeight || document.documentElement.clientHeight || 600;
    const availW = Math.max(200, vw - 48);   // 弹窗左右边距
    const availH = Math.max(160, vh - 220);  // 标题/元信息/按钮占位
    const maxW = Math.min(opts.maxFrameW || 520, availW);
    const maxH = Math.min(opts.maxFrameH || 340, availH);
    const { w: frameW, h: frameH } = stageSizeForRatio(opts.frameRatio || 1, maxW, maxH);

    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal-box crop-modal">
        <h3 class="modal-title">✂ 裁剪图片（按模板图片位比例）</h3>
        <div class="crop-meta">
          <span class="crop-meta-item">${escapeText(opts.frameSizeLabel)}</span>
          <span class="crop-meta-item" id="crop-img-size">图片原尺寸：读取中…</span>
          <span class="crop-zoom">
            <button class="secondary crop-zoom-btn" id="crop-zoom-out" title="缩小">−</button>
            <button class="secondary crop-zoom-btn" id="crop-zoom-in" title="放大">＋</button>
          </span>
        </div>
        <div class="crop-stage" id="crop-stage" style="width:${frameW}px;height:${frameH}px">
          <div class="crop-stage-inner">
            <img id="crop-img" alt="待裁剪图片" draggable="false" />
          </div>
          <div class="crop-frame" style="width:${frameW}px;height:${frameH}px">
            <span class="crop-frame-label">模板图片位 ${Math.round(opts.frameRatio * 100) / 100}:1</span>
          </div>
          <p class="crop-tip">拖动图片调整位置；滚轮或 ＋/− 缩放；裁剪框内即最终效果</p>
        </div>
        <div class="crop-actions">
          <button class="secondary" id="crop-fit">适应裁剪框</button>
          <button class="secondary" id="crop-reset">重置</button>
          <button class="secondary" id="crop-cancel">取消</button>
          <button class="primary" id="crop-ok">确定</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    const stage = overlay.querySelector('#crop-stage') as HTMLElement;
    const imgEl = overlay.querySelector('#crop-img') as HTMLImageElement;
    const imgSizeEl = overlay.querySelector('#crop-img-size') as HTMLElement;
    const done = (result: CropEditorResult) => { overlay.remove(); resolve(result); };

    const img = new Image();
    img.onload = () => {
      const naturalW = img.naturalWidth || 1;
      const naturalH = img.naturalHeight || 1;
      imgSizeEl.textContent = `图片原尺寸：${naturalW} × ${naturalH} px`;
      // 初始：恰好覆盖裁剪框并居中
      let state: CropState = fitCrop({ naturalW, naturalH, frameW, frameH, scale: 1, tx: 0, ty: 0 });
      const minScale = coverScale(naturalW, naturalH, frameW, frameH);
      const maxScale = Math.max(minScale * 12, 8);

      const apply = (): void => {
        state = clampCrop(state);
        // 关键：缩放必须真正作用到图片本身（transform scale），且用 scale→translate 顺序，
        // 否则 translate 偏移会被 scale 放大、且视觉上缩放变成「移动」。
        imgEl.style.width = naturalW + 'px';
        imgEl.style.height = naturalH + 'px';
        // 顺序必须是 translate → scale：cropMath 的模型是「先平移(屏幕px)再缩放」，
        // 写成 scale→translate 会让平移量被 scale 放大，导致预览区域与实际裁剪结果不符。
        imgEl.style.transform = `translate(${state.tx}px, ${state.ty}px) scale(${state.scale})`;
        imgEl.style.transformOrigin = '0 0';
        imgEl.style.position = 'absolute';
        imgEl.style.left = '0';
        imgEl.style.top = '0';
      };

      // 以 stage 中心缩放
      const zoom = (factor: number): void => {
        const centerX = frameW / 2;
        const centerY = frameH / 2;
        const imgW = naturalW * state.scale;
        const imgH = naturalH * state.scale;
        const cx = state.tx + imgW / 2; // 图片中心（stage 坐标）
        const cy = state.ty + imgH / 2;
        const scale = Math.min(maxScale, Math.max(minScale, state.scale * factor));
        // 保持中心不变
        state = {
          ...state, scale,
          tx: cx - (naturalW * scale) / 2,
          ty: cy - (naturalH * scale) / 2
        };
        apply();
      };

      // 拖动（Pointer Events，兼容鼠标/触摸）
      let dragging = false;
      let lastX = 0, lastY = 0;
      stage.addEventListener('pointerdown', (e) => {
        dragging = true;
        lastX = e.clientX; lastY = e.clientY;
        stage.setPointerCapture(e.pointerId);
      });
      stage.addEventListener('pointermove', (e) => {
        if (!dragging) return;
        state = { ...state, tx: state.tx + (e.clientX - lastX), ty: state.ty + (e.clientY - lastY) };
        lastX = e.clientX; lastY = e.clientY;
        apply();
      });
      const endDrag = () => { dragging = false; };
      stage.addEventListener('pointerup', endDrag);
      stage.addEventListener('pointercancel', endDrag);

      // 滚轮缩放
      stage.addEventListener('wheel', (e) => {
        e.preventDefault();
        zoom(e.deltaY < 0 ? 1.1 : 1 / 1.1);
      }, { passive: false });

      overlay.querySelector('#crop-zoom-in')!.addEventListener('click', () => zoom(1.15));
      overlay.querySelector('#crop-zoom-out')!.addEventListener('click', () => zoom(1 / 1.15));
      overlay.querySelector('#crop-fit')!.addEventListener('click', () => { state = fitCrop(state); apply(); });
      overlay.querySelector('#crop-reset')!.addEventListener('click', () => { state = fitCrop(state); apply(); });
      overlay.querySelector('#crop-cancel')!.addEventListener('click', () => done({ canceled: true }));

      // 确定：canvas 裁出选中区域（输出长边 ≤2048，防超大图内存爆炸）
      overlay.querySelector('#crop-ok')!.addEventListener('click', () => {
        try {
          const { sx, sy, sw, sh } = cropRectFromState(state);
          const out = outputSize(sw, sh, 2048);
          if (!out.w || !out.h) { done({ canceled: true }); return; }
          const canvas = document.createElement('canvas');
          canvas.width = out.w; canvas.height = out.h;
          const ctx = canvas.getContext('2d');
          if (!ctx) { done({ canceled: true }); return; }
          ctx.drawImage(img, sx, sy, sw, sh, 0, 0, out.w, out.h);
          const isPng = opts.imageDataUrl.startsWith('data:image/png');
          const dataUrl = canvas.toDataURL(isPng ? 'image/png' : 'image/jpeg', 0.92);
          done({ canceled: false, dataUrl });
        } catch {
          done({ canceled: true });
        }
      });

      // 关闭：点击遮罩（弹窗外区域）等同取消
      overlay.addEventListener('click', (e) => { if (e.target === overlay) done({ canceled: true }); });

      apply();
      imgEl.src = opts.imageDataUrl;
    };
    img.onerror = () => done({ canceled: true });
    img.src = opts.imageDataUrl;
  });
}

function escapeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
