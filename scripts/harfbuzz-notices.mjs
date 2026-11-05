// The JavaScript metafile cannot identify code inside the copied WebAssembly binary.
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';

const coreRoot = resolve(import.meta.dirname, '../packages/core');
export const HARFBUZZ_MANIFEST = JSON.parse(
  readFileSync(resolve(coreRoot, 'licenses/harfbuzz-manifest.json'), 'utf8')
);

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');

/** Fail before publishing if a dependency upgrade bypasses the notice inventory. */
export function validateHarfBuzzNotices(options = {}) {
  const root = options.coreRoot ?? coreRoot;
  const manifest = JSON.parse(
    readFileSync(resolve(root, 'licenses/harfbuzz-manifest.json'), 'utf8')
  );
  const dependencyRoot =
    options.dependencyRoot ??
    resolve(dirname(createRequire(resolve(root, 'package.json')).resolve('harfbuzzjs')), '..');
  const dependency = JSON.parse(readFileSync(resolve(dependencyRoot, 'package.json'), 'utf8'));
  const declared = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
  if (
    dependency.name !== manifest.package.name ||
    dependency.version !== manifest.package.version ||
    declared.devDependencies.harfbuzzjs !== manifest.package.version
  ) {
    throw new Error(
      'HarfBuzz version changed. Update the component licenses and manifest before publishing.'
    );
  }
  for (const file of [
    resolve(dependencyRoot, 'dist/harfbuzz.wasm'),
    resolve(root, 'dist/harfbuzz.wasm'),
  ]) {
    if (sha256(readFileSync(file)) !== manifest.package.wasmSha256) {
      throw new Error(`HarfBuzz binary differs from the notice manifest: ${file}`);
    }
  }
  for (const notice of manifest.notices) {
    if (sha256(readFileSync(resolve(root, 'licenses', notice.file))) !== notice.sha256) {
      throw new Error(`HarfBuzz notice differs from the source manifest: ${notice.file}`);
    }
  }
}

export const HARFBUZZ_NOTICE =
  '## HarfBuzz\n\n' +
  '`dist/harfbuzz.wasm` contains HarfBuzz, distributed by `harfbuzzjs` (its license is below). ' +
  'The component versions, source URLs, and notice checksums are in `licenses/harfbuzz-manifest.json`. ' +
  'The following files contain the component copyright and permission notices. ' +
  'The source and runtime notices also cover optional upstream components.\n\n' +
  HARFBUZZ_MANIFEST.notices
    .map((notice) => `- ${notice.component}: \`licenses/${notice.file}\``)
    .join('\n') +
  '\n';
