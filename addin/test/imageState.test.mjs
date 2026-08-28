import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addinRoot = path.resolve(testDir, '..');

async function loadModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'ppt-addin-image-state-'));
  await build({
    entryPoints: [path.join(addinRoot, 'src/pages/wizard/imageState.ts')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    outfile: path.join(outdir, 'imageState.mjs'),
    logLevel: 'silent',
  });
  const mod = await import(pathToFileURL(path.join(outdir, 'imageState.mjs')).href);
  return { mod, cleanup: () => rm(outdir, { recursive: true, force: true }) };
}

test('default image state uses configured page size when valid', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    assert.equal(mod.normalizeImagePageSize(12), 12);
    assert.equal(mod.defaultImgState(12).pageSize, 12);
  } finally {
    await cleanup();
  }
});

test('invalid page size falls back to default', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    assert.equal(mod.normalizeImagePageSize(3), 9);
    assert.equal(mod.normalizeImagePageSize(20), 9);
    assert.equal(mod.normalizeImagePageSize('abc'), 9);
  } finally {
    await cleanup();
  }
});

test('thumbnail falls back to original image url', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    const img = { imageUrl: 'https://example.com/full.jpg', thumbnailUrl: '' };
    assert.equal(mod.thumbnailUrlOf(img), 'https://example.com/full.jpg');
  } finally {
    await cleanup();
  }
});
