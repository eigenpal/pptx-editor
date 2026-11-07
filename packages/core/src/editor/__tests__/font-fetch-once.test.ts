// A packaged font file loads once per open: the engine loads the packaged substitutes of the
// deck's families itself, and a host origin that serves the same files (`packagedFonts()`)
// skips them.

import { expect, test } from 'bun:test';
import { packagedFonts } from '@docx-editor.dev/fonts';
import { createPptxEditor } from '../index.ts';
import { customFonts } from '../custom-fonts.ts';
import { composeFontOrigins, defineFontResolver } from '../font-resolver.ts';
import { fontBytes, missingFontsDeck, settle } from './font-test-helpers.ts';

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

test('a host origin that serves the real Calibri and Arial still loads them', async () => {
  const fetched: string[] = [];
  const serving = (async (input: RequestInfo | URL) => {
    const url = String(input);
    fetched.push(url);
    if (!url.startsWith('/fonts/')) return new Response(null, { status: 404 });
    return new Response(fontBytes('LiberationSans-Regular.ttf').slice().buffer);
  }) as typeof fetch;
  const sources = ['Calibri', 'Arial'].map((family) => ({
    url: `/fonts/${family}.ttf`,
    family,
    weight: 400,
    style: 'normal' as const,
  }));
  const editor = createPptxEditor({
    fonts: defineFontResolver((request) =>
      composeFontOrigins(
        [customFonts({ sources, fetcher: serving }), packagedFonts({ fetcher: serving })],
        request
      )
    ),
  });
  expect((await editor.load(missingFontsDeck())).ok).toBe(true);
  await settle(10);
  expect(fetched).toContain('/fonts/Calibri.ttf');
  expect(fetched).toContain('/fonts/Arial.ttf');
  expect(fetched.filter((url) => /Carlito|LiberationSans-/.test(url))).toEqual([]);
  editor.destroy();
});
