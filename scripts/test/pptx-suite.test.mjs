import { expect, test } from 'bun:test';
import { resolveSpecifier } from './changed.mjs';
import { DOM_PRELOAD, bunArgumentsFor } from './pptx-suite.mjs';

test('only the adapter tests preload a DOM', () => {
  expect(bunArgumentsFor('packages/react/src/__tests__/adapter.test.tsx')).toEqual([
    '--preload',
    DOM_PRELOAD,
  ]);
  expect(bunArgumentsFor('packages/vue/test/adapter.test.ts')).toEqual(['--preload', DOM_PRELOAD]);
  for (const file of [
    'packages/core/src/layout/__tests__/a.test.ts',
    'packages/ooxml-core/src/__tests__/zip.test.ts',
    'packages/editor-api/src/__tests__/a.test.ts',
    'scripts/test/pptx-suite.test.mjs',
  ]) {
    expect(bunArgumentsFor(file)).toEqual([]);
  }
});

// This file runs in the DOM-free group, as every engine test does. A global `document` here
// means the runner preloaded happy-dom for a file that must not have it.
test('a DOM-free file runs with no global document', () => {
  expect(typeof document).toBe('undefined');
  expect(typeof window).toBe('undefined');
});

test('both workspace scopes resolve to their packages', () => {
  expect(resolveSpecifier('packages/core/src/a.ts', '@eigenpal/ooxml-core')).toBe(
    'packages/ooxml-core/src/index.ts'
  );
  expect(resolveSpecifier('packages/react/src/a.ts', '@pptx-editor.dev/core/editor')).toBe(
    'packages/core/src/editor/index.ts'
  );
});
