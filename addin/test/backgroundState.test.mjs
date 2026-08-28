import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addinRoot = path.resolve(testDir, '..');

async function loadModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'ppt-addin-bg-state-'));
  await build({
    entryPoints: [path.join(addinRoot, 'src/pages/wizard/backgroundState.ts')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    outfile: path.join(outdir, 'backgroundState.mjs'),
    logLevel: 'silent',
  });
  const mod = await import(pathToFileURL(path.join(outdir, 'backgroundState.mjs')).href);
  return { mod, cleanup: () => rm(outdir, { recursive: true, force: true }) };
}

test('default wizard background follows the template document background', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    const template = { background: { type: 'solid', color: '#123456' } };
    const state = mod.defaultWizardBackgroundState(template);
    assert.equal(state.followDocument, true);
    assert.deepEqual(mod.resolveWizardBackground(template, state), template.background);
  } finally {
    await cleanup();
  }
});

test('custom wizard background overrides the template background without mutating it', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    const template = { background: { type: 'solid', color: '#123456' } };
    const state = { followDocument: false, customImageDataUrl: 'data:image/png;base64,AAAA' };
    const resolved = mod.resolveWizardBackground(template, state);
    assert.deepEqual(resolved, { type: 'picture', imageDataUrl: 'data:image/png;base64,AAAA' });
    assert.deepEqual(template.background, { type: 'solid', color: '#123456' });
  } finally {
    await cleanup();
  }
});
