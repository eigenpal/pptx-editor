import { afterEach, expect, test } from 'bun:test';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import {
  HARFBUZZ_MANIFEST,
  HARFBUZZ_NOTICE,
  validateHarfBuzzNotices,
} from './harfbuzz-notices.mjs';

const core = resolve(import.meta.dir, '../packages/core');
const temporary: string[] = [];

afterEach(() => {
  for (const directory of temporary.splice(0)) rmSync(directory, { recursive: true, force: true });
});

function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'harfbuzz-notices-'));
  temporary.push(root);
  const coreRoot = join(root, 'core');
  const dependencyRoot = join(root, 'dependency');
  mkdirSync(join(coreRoot, 'dist'), { recursive: true });
  mkdirSync(join(dependencyRoot, 'dist'), { recursive: true });
  cpSync(join(core, 'licenses'), join(coreRoot, 'licenses'), { recursive: true });
  cpSync(join(core, 'package.json'), join(coreRoot, 'package.json'));
  cpSync(join(core, 'node_modules/harfbuzzjs/package.json'), join(dependencyRoot, 'package.json'));
  for (const destination of [coreRoot, dependencyRoot]) {
    cpSync(
      join(core, 'node_modules/harfbuzzjs/dist/harfbuzz.wasm'),
      join(destination, 'dist/harfbuzz.wasm')
    );
  }
  return { coreRoot, dependencyRoot };
}

test('the installed binary has matching component notices', () => {
  expect(() => validateHarfBuzzNotices(fixture())).not.toThrow();
  const notices = HARFBUZZ_MANIFEST.notices
    .map((entry: { file: string }) => readFileSync(join(core, 'licenses', entry.file), 'utf8'))
    .join('\n');
  for (const copyright of [
    'Grigori Goronzy',
    'Copyright (c) Microsoft Corporation',
    'Emscripten authors',
    'Rich Felker',
    'Unicode, Inc.',
  ])
    expect(notices).toContain(copyright);
  for (const entry of HARFBUZZ_MANIFEST.notices) {
    expect(HARFBUZZ_NOTICE).toContain(`licenses/${entry.file}`);
  }
});

test('a dependency upgrade requires a new notice inventory', () => {
  const paths = fixture();
  const file = join(paths.dependencyRoot, 'package.json');
  const dependency = JSON.parse(readFileSync(file, 'utf8'));
  dependency.version = '99.0.0';
  writeFileSync(file, JSON.stringify(dependency));
  expect(() => validateHarfBuzzNotices(paths)).toThrow('HarfBuzz version changed');
});

test('a replaced release binary cannot reuse the old notice inventory', () => {
  const paths = fixture();
  writeFileSync(join(paths.coreRoot, 'dist/harfbuzz.wasm'), 'different binary');
  expect(() => validateHarfBuzzNotices(paths)).toThrow('HarfBuzz binary differs');
});

test('an altered permission notice fails the release check', () => {
  const paths = fixture();
  writeFileSync(join(paths.coreRoot, 'licenses/HarfBuzz-Microsoft-MIT.txt'), 'MIT');
  expect(() => validateHarfBuzzNotices(paths)).toThrow('HarfBuzz notice differs');
});

test('a missing component notice fails the release check', () => {
  const paths = fixture();
  rmSync(join(paths.coreRoot, 'licenses/HarfBuzz-source-notices.txt'));
  expect(() => validateHarfBuzzNotices(paths)).toThrow();
});
