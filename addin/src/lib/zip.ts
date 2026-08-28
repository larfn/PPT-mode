// Office.js 文档 zip 读取工具。
// 关键坑：getSliceAsync 每个分片是独立 base64（4MB 分片各自编码、末尾带 padding），
// 直接 join('') 会破坏数据（JSZip 报 Can't find end of central directory），必须按字节拼接。

// 单个分片 base64 解码为字节
export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

export function concatBytes(chunks: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let pos = 0;
  for (const c of chunks) { out.set(c, pos); pos += c.length; }
  return out;
}

// 通过 Office.js 获取整个文档（含未保存修改）的 zip 字节，用于后台回读精确样式。
// 大文档（数十 MB）读取可能很慢，给 60 秒上限；超时/失败只影响样式回读，不阻塞保存。
export function getDocumentZipBytes(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    let settled = false;
    let timer: number;
    const finish = (fn: () => void) => { if (settled) return; settled = true; if (timer) window.clearTimeout(timer); fn(); };
    timer = window.setTimeout(() => finish(() => reject(new Error('读取文档超时（文档过大，超过 120 秒）'))), 120000);
    Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        finish(() => reject(new Error(result.error?.message || '获取文档失败')));
        return;
      }
      const file = result.value;
      const chunks: Uint8Array[] = [];
      let offset = 0;
      const readSlice = () => {
        file.getSliceAsync(offset, (sliceResult) => {
          if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
            finish(() => file.closeAsync(() => reject(new Error('读取文档分片失败'))));
            return;
          }
          const raw = (sliceResult.value as { data: unknown }).data;
          const b64 = typeof raw === 'string' ? raw : String(raw || '');
          chunks.push(base64ToBytes(b64));
          offset++;
          if (offset < file.sliceCount) readSlice();
          else finish(() => file.closeAsync(() => resolve(concatBytes(chunks))));
        });
      };
      readSlice();
    });
  });
}
