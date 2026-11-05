// The shipped advisory scan covers the code a package build inlines, not only `dependencies`.
import { expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { bundledDependencies } from './bundled-dependencies.mjs';

const root = join(import.meta.dir, '..');

test('core inlines the HarfBuzz runtime from a devDependency', () => {
  expect(bundledDependencies(join(root, 'packages/core'))).toContain('harfbuzzjs');
  expect(bundledDependencies(join(root, 'packages/ooxml-core'))).toEqual([]);
});

test('a noExternal entry that is not a package name string is refused', () => {
  const dir = mkdtempSync(join(tmpdir(), 'bundled-deps-'));
  try {
    writeFileSync(join(dir, 'tsup.config.ts'), "export default { noExternal: [/^x/, 'y'] };\n");
    expect(() => bundledDependencies(dir)).toThrow(/not a package name string/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('the shipped advisory closure includes the inlined HarfBuzz at its locked version', () => {
  const run = Bun.spawnSync(['node', 'scripts/check-shipped-advisories.mjs', '--list-closure'], {
    cwd: root,
  });
  expect(run.exitCode).toBe(0);
  const closure = run.stdout.toString().split('\n');
  const manifest = JSON.parse(readFileSync(join(root, 'packages/core/package.json'), 'utf8')) as {
    devDependencies: Record<string, string>;
  };
  // core pins the exact HarfBuzz build (HARFBUZZ_VERSION), so the manifest names the locked version.
  const pinned = manifest.devDependencies.harfbuzzjs;
  expect(closure).toContain(`harfbuzzjs@${pinned}\t@pptx-editor.dev/core > harfbuzzjs`);
});
