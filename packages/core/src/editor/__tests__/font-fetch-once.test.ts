// A packaged font file loads once per open: the engine loads the packaged substitutes of the
// deck's families itself, and a host origin that serves the same files (`packagedFonts()`)
// skips them.

import { expect, test } from 'bun:test';
import { packagedFonts } from '@docx-editor.dev/fonts';
import { createPptxEditor } from '../index.ts';
import { missingFontsDeck, settle } from './font-test-helpers.ts';

test('packagedFonts() fetches no file that the engine loads itself', async () => {
  const fetched: string[] = [];
  const counting = (async (input: RequestInfo | URL) => {
    fetched.push(String(input));
    return new Response(null, { status: 404 });
  }) as typeof fetch;
  const editor = createPptxEditor({ fonts: packagedFonts({ fetcher: counting }) });
  expect((await editor.load(missingFontsDeck())).ok).toBe(true);
  await settle(10);
  // The deck names Calibri and Arial: the engine loads their substitutes.
  const families = editor.session!.substituteFamilies ?? [];
  expect(families).toContain('Calibri');
  expect(families).toContain('Arial');
  const twice = fetched.filter((url) => /Carlito|LiberationSans-/.test(url));
  expect(twice).toEqual([]);
  editor.destroy();
});
