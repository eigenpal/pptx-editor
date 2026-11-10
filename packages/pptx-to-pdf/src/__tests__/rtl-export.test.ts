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
  'לא שלום',
  'שָׁלוֹם עוֹלָם',
  'لا للا لآ لَّا',
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
      else expect(extracted).toBe('بالعالم hello 123 مرحبا');
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

test('font downloads cancelled during session open return a typed refusal', async () => {
  const { spyOn } = await import('bun:test');
  const { openDocumentForExport } = await import('../index.ts');
  const controller = new AbortController();
  let began!: () => void;
  const started = new Promise<void>((resolve) => {
    began = resolve;
  });
  const delayedFetch = Object.assign(
    (
      _input: Parameters<typeof globalThis.fetch>[0],
      init?: Parameters<typeof globalThis.fetch>[1]
    ): Promise<Response> => {
      began();
      return new Promise((_resolve, reject) => {
        const signal = init?.signal;
        if (signal?.aborted) reject(signal.reason);
        else signal?.addEventListener('abort', () => reject(signal.reason), { once: true });
      });
    },
    { preconnect: globalThis.fetch.preconnect }
  );
  const fetch = spyOn(globalThis, 'fetch').mockImplementation(delayedFetch);
  try {
    const opening = openDocumentForExport(deck('مرحبا'), {
      googleFonts: true,
      signal: controller.signal,
    });
    await started;
    controller.abort(new Error('client disconnected'));
    expect(await opening).toEqual({ ok: false, reason: 'aborted' });
  } finally {
    controller.abort();
    fetch.mockRestore();
  }
});

for (const [language, family] of [
  ['zh-MO', 'Noto Sans TC'],
  ['yue-Hant', 'Noto Sans TC'],
  ['cmn-Hans', 'Noto Sans SC'],
  ['ko-KR', 'Noto Sans KR'],
]) {
  test(`CJK fallback follows ${language}`, () => {
    const bytes = buildPptx({
      slides: [
        {
          shapes: shape({
            id: 2,
            x: 0,
            y: 0,
            cx: 7315200,
            cy: 1828800,
            text: para(run('漢字', `altLang="${language}"`)),
          }),
        },
      ],
    });
    const read = readPptx(bytes);
    if (!read.ok) throw new Error('fixture');
    const request = deckFontRequest(read.pkg);
    expect(request.families).toContain(family);
    expect(request.families).not.toContain('Noto Sans JP');
  });
}

test('a full declared-font list still requests Arabic coverage', () => {
  const text = Array.from({ length: 80 }, (_, i) =>
    run('مرحبا', '', `<a:cs typeface="Unknown ${i}"/>`)
  ).join('');
  const read = readPptx(
    buildPptx({
      slides: [
        { shapes: shape({ id: 2, x: 0, y: 0, cx: 7315200, cy: 1828800, text: para(text) }) },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  const request = deckFontRequest(read.pkg);
  expect(request.families.length).toBeLessThanOrEqual(64);
  expect(request.families).toContain('Amiri');
});

test('a refused shaping operation retains text for nominal paint', async () => {
  const { shapeRtlRun } = await import('../../../core/src/layout/text/shaped-run.ts');
  const { createFixedMeasurer } = await import('@pptx-editor.dev/core/layout');
  const opened = await openPresentation(deck('مرحبا'), { fixedMeasurer: true });
  if (!opened.ok) throw new Error('fixture');
  try {
    const item = opened.session.layout(0)?.items.find((i) => i.kind === 'shape');
    if (!item || item.kind !== 'shape') throw new Error('shape');
    const run = item.text!.lines.flatMap((line) => line.runs).find((r) => r.rtl)!;
    expect(run.text).toBe('مرحبا');
    expect(shapeRtlRun(run, { ...createFixedMeasurer(), shape: () => [] })).toBe(run);
  } finally {
    opened.session.dispose();
  }
});

test('Arabic script fonts use fallback punctuation instead of omitting it', async () => {
  const result = await exportPdf(deck('مرحبا (123) بالعالم!'), {
    fonts,
    fidelityPolicy: 'best-effort',
  });
  expect(result.diagnostics.some((d) => d.code === 'text.missing-glyph')).toBe(false);
  const task = getDocument({ data: result.bytes.slice(), useSystemFonts: false });
  try {
    const pdf = await task.promise;
    const items = (await (await pdf.getPage(1)).getTextContent()).items;
    const text = items.map((i) => ('str' in i ? i.str : '')).join('');
    expect(text).toContain('123');
    expect(text).toContain('(');
    expect(text).toContain(')');
    expect(text.match(/ /g)).toHaveLength(2);
  } finally {
    await task.destroy();
  }
});

test('Arabic gradient text and its shadow keep one searchable copy of the marks', async () => {
  const text = 'السَّلَامُ عَلَيْكُمْ';
  const gradient =
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="336699"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';
  const bytes = buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: 914400,
          y: 914400,
          cx: 7315200,
          cy: 1828800,
          effects:
            '<a:effectLst><a:outerShdw blurRad="0" dist="25400" dir="2700000"><a:srgbClr val="888888"/></a:outerShdw></a:effectLst>',
          text: para(
            run(text, '', `${gradient}<a:cs typeface="Noto Sans Arabic"/>`),
            '<a:pPr rtl="1" algn="r"/>'
          ),
        }),
      },
    ],
  });
  const result = await exportPdf(bytes, { fonts, fidelityPolicy: 'best-effort' });
  expect(result.diagnostics.some((d) => d.code === 'text.missing-glyph')).toBe(false);
  const task = getDocument({ data: result.bytes.slice(), useSystemFonts: false });
  try {
    const pdf = await task.promise;
    const page = await pdf.getPage(1);
    const items = (await page.getTextContent()).items;
    expect(
      items
        .map((i) => ('str' in i ? i.str : ''))
        .join('')
        .normalize('NFC')
    ).toBe(text.normalize('NFC'));
    const { fnArray } = await page.getOperatorList();
    const { OPS } = await import('pdfjs-dist/legacy/build/pdf.mjs');
    expect(fnArray).toContain(OPS.shadingFill);
    expect(fnArray).toContain(OPS.paintFormXObjectBegin);
  } finally {
    await task.destroy();
  }
});

test('CJK fallback inherits the presentation language when runs omit it', () => {
  const bytes = buildPptx({
    defaultLanguage: 'zh-Hant',
    slides: [
      {
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 7315200,
          cy: 1828800,
          text: para(run('漢字').replace('lang="en-US"', '')),
        }),
      },
    ],
  });
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('fixture');
  const request = deckFontRequest(read.pkg);
  expect(request.families).toContain('Noto Sans TC');
  expect(request.families).not.toContain('Noto Sans JP');
});
