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
  const outdir = await mkdtemp(path.join(tmpdir(), 'ppt-addin-settings-'));
  await build({
    entryPoints: [path.join(addinRoot, 'src/pages/settings.ts')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    outfile: path.join(outdir, 'settings.mjs'),
    logLevel: 'silent',
  });
  const mod = await import(pathToFileURL(path.join(outdir, 'settings.mjs')).href);
  return { mod, cleanup: () => rm(outdir, { recursive: true, force: true }) };
}

test('custom source preview url prefers thumbnail and falls back to image url', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    assert.equal(
      mod.previewImageUrlOf({ thumbnailUrl: 'https://example.com/thumb.jpg', imageUrl: 'https://example.com/full.jpg' }),
      'https://example.com/thumb.jpg',
    );
    assert.equal(
      mod.previewImageUrlOf({ thumbnailUrl: '', imageUrl: 'https://example.com/full.jpg' }),
      'https://example.com/full.jpg',
    );
    assert.equal(mod.previewImageUrlOf(null), '');
  } finally {
    await cleanup();
  }
});
