// 返回当前选中页的 PNG dataUrl（如 "data:image/png;base64,..."）；任何失败或超时都返回空串，不阻塞保存流程。
// PowerPoint 对大文档执行 getImageAsBase64 可能很慢，给 8 秒上限，超时则放弃预览直接保存。
export async function captureSlidePreview(): Promise<string> {
  const timeout = new Promise<string>((resolve) => { window.setTimeout(() => resolve(''), 8000); });
  const work = (async () => {
    try {
      return await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const result = slide.getImageAsBase64();
        await context.sync();
        const value = result.value;
        if (!value) return '';
        return value.startsWith('data:') ? value : `data:image/png;base64,${value}`;
      });
    } catch {
      return '';
    }
  })();
  return Promise.race([work, timeout]);
}