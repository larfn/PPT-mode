import { defineConfig } from 'vite';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

// 统一版本号：由 scripts/release.js 每次发布时写入根目录 release.json（vite 构建前已生成）；
// 未发布（开发模式 / 单独 npm run build）回退 0.0.0-dev。
let appVersion = '0.0.0-dev';
let builtAt = '';
try {
  const rel = JSON.parse(fs.readFileSync(fileURLToPath(new URL('../release.json', import.meta.url)), 'utf8'));
  if (rel.version) appVersion = rel.version;
  if (rel.builtAt) builtAt = rel.builtAt;
} catch { /* 未发布 */ }

export default defineConfig({
  root: '.',
  define: {
    __APP_VERSION__: JSON.stringify(appVersion),
    __APP_BUILT_AT__: JSON.stringify(builtAt)
  },
  server: {
    host: 'localhost',
    port: 3000,
    https: false,
    proxy: {
      '/api': 'http://127.0.0.1:3788'
    }
  },
  build: { outDir: 'dist', emptyOutDir: true }
});