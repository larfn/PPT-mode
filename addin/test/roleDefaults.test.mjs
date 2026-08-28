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
  const outdir = await mkdtemp(path.join(tmpdir(), 'ppt-addin-role-defaults-'));
  await build({
    entryPoints: [path.join(addinRoot, 'src/lib/roleDefaults.ts')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    outfile: path.join(outdir, 'roleDefaults.mjs'),
    logLevel: 'silent',
  });
  const mod = await import(pathToFileURL(path.join(outdir, 'roleDefaults.mjs')).href);
  return { mod, cleanup: () => rm(outdir, { recursive: true, force: true }) };
}

test('table shape defaults to table role instead of fixed', async () => {
  const { mod, cleanup } = await loadModule();
  try {
    assert.equal(mod.defaultRoleForShape({ type: 'table', hasText: false }), 'table');
  } finally {
    await cleanup();
  }
});
