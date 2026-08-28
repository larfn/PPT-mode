import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addinRoot = path.resolve(testDir, '..');

async function loadModule(entry, outfileName) {
  const outdir = await mkdtemp(path.join(tmpdir(), 'ppt-addin-preview-'));
  await build({
    entryPoints: [path.join(addinRoot, entry)],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    outfile: path.join(outdir, outfileName),
    logLevel: 'silent',
  });
  const mod = await import(pathToFileURL(path.join(outdir, outfileName)).href);
  return { mod, cleanup: () => rm(outdir, { recursive: true, force: true }) };
}

test('renderPreview scales the whole slide instead of leaving font size fixed', async () => {
  const { mod, cleanup } = await loadModule('src/previewPanel.ts', 'previewPanel.mjs');
  try {
    const html = mod.renderPreview({
      slideSize: { width: 13.333, height: 7.5 },
      shapes: [{
        id: 'title',
        role: 'ai_text',
        type: 'text',
        bounds: { left: 0.5, top: 0.8, width: 4, height: 0.6 },
        textStyle: { size: 24, bold: true },
      }],
    }, {}, { title: '子系统设计思路' }, {});
    assert.match(html, /<svg\b/);
    assert.match(html, /viewBox="0 0 1280 720"/);
    assert.match(html, /foreignObject/);
    assert.match(html, /font-size:32\.0px/);
  } finally {
    await cleanup();
  }
});

test('quality check suppresses mild text overflow when PowerPoint autoFit is enabled', async () => {
  const { mod, cleanup } = await loadModule('src/lib/qualityCheck.ts', 'qualityCheck.mjs');
  try {
    const template = {
      slideSize: { width: 13.333, height: 7.5 },
      shapes: [{
        id: 'body',
        role: 'ai_text',
        type: 'text',
        bounds: { left: 7.8, top: 1.4, width: 3.0, height: 3.5 },
        textStyle: { size: 20, autoFit: 'shrink', margin: { left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 } },
      }],
    };
    const text = '1.采用摆臂式越障机构，实现跨越防振锤等典型障碍物。\n2.前后轮组交替越障。\n3.中部刚性平台承担主体安装，同时考虑减重与空间布置。\n4.越障过程中中央夹臂提供辅助支撑，需考虑运动约束方向问题。';
    const report = await mod.runQualityChecks(template, {}, { body: text }, {}, {});
    assert.equal(report.issues.some((issue) => issue.category === '文本溢出'), false);
  } finally {
    await cleanup();
  }
});
