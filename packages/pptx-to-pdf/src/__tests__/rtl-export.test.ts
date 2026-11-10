/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { openPresentation, deckFontRequest } from '@pptx-editor.dev/core/editor';
import { readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx, shape, para, run } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const fonts = [
  {
    family: 'Noto Sans Arabic',
    origin: 'test',
    bold: false,
    italic: false,
    bytes: new Uint8Array(
      readFileSync(
        new URL('../../../../fixtures/fonts/NotoSansArabic-Regular.ttf', import.meta.url)
      )
    ),
  },
];
function deck(text: string, family = 'Noto Sans Arabic') {
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: 914400,
          y: 914400,
          cx: 7315200,
          cy: 1828800,
          text: para(
            run(text, '', `<a:latin typeface="Arial"/><a:cs typeface="${family}"/>`),
            '<a:pPr rtl="1" algn="r"/>'
          ),
        }),
      },
    ],
  });
}

for (const text of [
  'مرحبا بالعالم الحقيقي',
  'السَّلَامُ عَلَيْكُمْ',
  'سلام دنیا',
  'مرحبا hello 123 بالعالم',
]) {
  test(`Arabic export preserves searchable words: ${text}`, async () => {
    const result = await exportPdf(deck(text), { fonts, fidelityPolicy: 'best-effort' });
    expect(result.diagnostics.some((d) => d.code === 'text.missing-glyph')).toBe(false);
    const pdf = await getDocument({ data: result.bytes.slice(), useSystemFonts: false }).promise;
    try {
      const items = (await (await pdf.getPage(1)).getTextContent()).items;
      const extracted = items.map((i) => ('str' in i ? i.str : '')).join('');
      if (!text.includes('hello')) expect(extracted.normalize('NFC')).toBe(text.normalize('NFC'));
      else for (const word of text.split(' ')) expect(extracted).toContain(word);
    } finally {
      await pdf.loadingTask.destroy();
    }
  });
}

test('Arabic fallback is requested from text even when the declared family is Latin', () => {
  const read = readPptx(deck('مرحبا', 'Arial'));
  if (!read.ok) throw new Error('fixture');
  expect(deckFontRequest(read.pkg).families).toContain('Amiri');
});

test('layout retains contextual glyphs and their mark offsets', async () => {
  const opened = await openPresentation(deck('السَّلَامُ عَلَيْكُمْ'), { fonts });
  if (!opened.ok) throw new Error('fixture');
  try {
    const item = opened.session.layout(0)?.items.find((item) => item.kind === 'shape');
    if (!item || item.kind !== 'shape') throw new Error('text shape');
    const glyphs = item.text!.lines.flatMap((line) => line.runs.flatMap((run) => run.glyphs ?? []));
    expect(glyphs.length).toBeGreaterThan(0);
    expect(glyphs.some((g) => g.text === '' && g.y !== 0)).toBe(true);
    expect(glyphs.every((g) => Number.isFinite(g.x) && Number.isFinite(g.y))).toBe(true);
  } finally {
    opened.session.dispose();
  }
});

test('Google font intake verifies hashes and keeps caller fonts first', async () => {
  const { googleExportFonts } = await import('../google-fonts.ts');
  const { sha256FontBytes } = await import('@pptx-editor.dev/core/layout');
  const bytes = fonts[0].bytes;
  const face = {
    id: 'arabic',
    request: { family: 'Amiri', weight: 400, style: 'normal' as const },
    faceIndex: 0,
    bytes,
    hash: sha256FontBytes(bytes),
  };
  const supplied = { ...fonts[0], family: 'Amiri' };
  expect(await googleExportFonts(deck('مرحبا'), undefined, [], { sources: [face] })).toHaveLength(
    1
  );
  expect(
    await googleExportFonts(deck('مرحبا'), undefined, [supplied], { sources: [face] })
  ).toEqual([]);
  await expect(
    googleExportFonts(deck('مرحبا'), undefined, [], {
      sources: [{ ...face, hash: 'sha256:incorrect' }],
    })
  ).rejects.toThrow('admission');
  const controller = new AbortController();
  controller.abort();
  expect(
    await googleExportFonts(deck('مرحبا'), controller.signal, [], { sources: [face] })
  ).toEqual([]);
});

test('CJK text requests an open fallback even when no East Asian font is declared', () => {
  const read = readPptx(deck('日本語', 'Arial'));
  if (!read.ok) throw new Error('fixture');
  expect(deckFontRequest(read.pkg).families).toContain('Noto Sans JP');
});

test('Google Fonts opt-in accepts only booleans', async () => {
  await expect(
    exportPdf(deck('مرحبا'), { googleFonts: 'true' as unknown as boolean })
  ).rejects.toThrow('googleFonts must be boolean');
});

test('Arabic fallback preserves spaces across Latin-font space runs', async () => {
  const result = await exportPdf(deck('مرحبا بالعالم الحقيقي', 'Arial'), {
    fonts,
    fidelityPolicy: 'best-effort',
  });
  const task = getDocument({ data: result.bytes.slice(), useSystemFonts: false });
  try {
    const pdf = await task.promise;
    const items = (await (await pdf.getPage(1)).getTextContent()).items;
    expect(items.map((i) => ('str' in i ? i.str : '')).join('')).toBe('مرحبا بالعالم الحقيقي');
  } finally {
    await task.destroy();
  }
});
