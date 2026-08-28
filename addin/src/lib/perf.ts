// 前端性能埋点：关键操作计时 → 批量 POST /api/perf/log（后端环形缓冲 + 超预算 warn 日志）。
// 用户决策：性能指标只写日志、不展示界面；此处全部静默失败，绝不影响主流程。
import { Api } from '../api.js';

export interface PerfEntry {
  op: string;
  ms: number;
  meta?: Record<string, unknown>;
}

const pending: PerfEntry[] = [];
let flushing = false;
let lastFlush = 0;

// 记录一条耗时（累计到 5 条或 3 秒后自动上报）
export function perfRecord(op: string, ms: number, meta?: Record<string, unknown>): void {
  const entry: PerfEntry = { op, ms: Math.round(ms), ...(meta ? { meta } : {}) };
  pending.push(entry);
  const now = Date.now();
  if (pending.length >= 5 || (pending.length > 0 && now - lastFlush > 3000)) void flushPerf();
}

// 计时一个 Promise（生成/读取/插入等）
export async function perfAsync<T>(op: string, p: Promise<T>, meta?: Record<string, unknown>): Promise<T> {
  const t0 = performance.now();
  try {
    return await p;
  } finally {
    perfRecord(op, performance.now() - t0, meta);
  }
}

// 立即上报（页面卸载/长任务结束时调用；失败静默）
export function flushPerf(): void {
  if (flushing || !pending.length) return;
  flushing = true;
  const batch = pending.splice(0, pending.length);
  lastFlush = Date.now();
  Api.logPerf(batch).catch(() => {}).finally(() => { flushing = false; });
}