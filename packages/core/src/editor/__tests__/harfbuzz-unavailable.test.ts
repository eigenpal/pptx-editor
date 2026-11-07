// A text engine that does not load (a missing `harfbuzz.wasm`, or a page policy that blocks
// WebAssembly): the deck still opens, with the fixed measurer, and the failure reports through
// `onFontError`, as DOCX Editor degrades to its fixed measurer.
//
// The module mock needs its own process: `bun run test` gives each file one.

import { expect, mock, test } from 'bun:test';

// A runtime of another version is refused, as a binary that does not load is.
mock.module('harfbuzzjs', () => ({ versionString: () => '0.0.0' }));

test('the deck opens with the fixed measurer and reports the failure', async () => {
  const { createPptxEditor } = await import('../index.ts');
  const { buildPptx } = await import('../../../../../scripts/fixtures/builder.ts');
  const errors: string[] = [];
  const editor = createPptxEditor({ onFontError: (e) => void errors.push(e.message) });
  expect((await editor.load(buildPptx({ slides: [{ shapes: '' }] }))).ok).toBe(true);
  expect(editor.snapshot().status).toBe('ready');
  for (let i = 0; i < 100 && editor.fontMeasurement().resolving; i += 1)
    await new Promise((r) => setTimeout(r, 5));
  expect(editor.fontMeasurement()).toEqual({ measurer: 'fixed', resolving: false });
  expect(errors).toHaveLength(1);
  expect(errors[0]).toContain('HarfBuzz');
  editor.destroy();
});
