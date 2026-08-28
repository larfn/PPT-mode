// 图片压缩工具：把大图 dataURL 压成小尺寸 JPEG。
// 模板库缩略图只需几百像素宽，全尺寸 PNG（可能 1-2MB）会撑大模板目录与上传流量。
export async function compressImageDataUrl(dataUrl: string, maxWidth = 480, quality = 0.82): Promise<string> {
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error('image load failed'));
      i.src = dataUrl;
    });
    const scale = Math.min(1, maxWidth / Math.max(img.naturalWidth, 1));
    const w = Math.max(1, Math.round(img.naturalWidth * scale));
    const h = Math.max(1, Math.round(img.naturalHeight * scale));
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    if (!ctx) return dataUrl;
    ctx.drawImage(img, 0, 0, w, h);
    return canvas.toDataURL('image/jpeg', quality);
  } catch {
    return dataUrl; // 压缩失败原样返回，不阻塞流程
  }
}
