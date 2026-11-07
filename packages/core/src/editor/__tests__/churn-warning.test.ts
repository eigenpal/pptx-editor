// A root prop made anew on each render warns once in development, with the fix.

import { afterEach, expect, test } from 'bun:test';
import { churnWarning, resetDevWarnings } from '../dev-warnings.ts';

const warn = console.warn;
afterEach(() => {
  console.warn = warn;
  resetDevWarnings();
});

test('ten changes a second are quiet; the eleventh warns once, with the fix', () => {
  const printed: string[] = [];
  console.warn = (message: string) => void printed.push(message);
  const changed = churnWarning('PptxEditor.Root', 'document');
  for (let i = 0; i < 10; i += 1) changed();
  expect(printed).toEqual([]);
  for (let i = 0; i < 5; i += 1) changed();
  expect(printed).toHaveLength(1);
  expect(printed[0]).toContain('PptxEditor.Root got new `document` bytes 11 times');
  expect(printed[0]).toContain('Keep the bytes in state');
});

test('the fonts warning names useFonts', () => {
  const printed: string[] = [];
  console.warn = (message: string) => void printed.push(message);
  const changed = churnWarning('PptxEditorRoot', 'fonts');
  for (let i = 0; i < 11; i += 1) changed();
  expect(printed[0]).toContain('useFonts(...)');
});
