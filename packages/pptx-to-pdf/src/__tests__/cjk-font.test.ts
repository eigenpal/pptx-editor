/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The East Asian fallback face of PDF export: Noto Sans CJK JP from the optional
// `@docx-editor.dev/fonts-cjk`, used only for decks with East Asian text.

import { describe, expect, test } from 'bun:test';
import { readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { CJK_FAMILY, cjkFallbackFont, hasEastAsianText, importCjkFace } from '../cjk-font.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const deck = (text: string, ea = ''): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run(text, 'altLang="ja-JP"', ea ? `<a:ea typeface="${ea}"/>` : '')),
        }),
      },
    ],
  });

const pkg = (bytes: Uint8Array) => {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('deck');
  return r.pkg;
};

describe('the CJK fallback face', () => {
  test('a deck holds East Asian text when a slide shows kana, Hangul, or ideographs', () => {
    expect(hasEastAsianText(pkg(deck('日本語')))).toBe(true);
    expect(hasEastAsianText(pkg(deck('한국어')))).toBe(true);
    expect(hasEastAsianText(pkg(deck('Latin text')))).toBe(false);
  });

  test('the face loads only for East Asian text, and only when the package is installed', async () => {
    expect(await cjkFallbackFont(deck('Latin text'))).toBeUndefined();
    expect(await cjkFallbackFont(deck('日本語'), async () => null)).toBeUndefined();
    const face = await cjkFallbackFont(deck('日本語'));
    expect(face?.family).toBe(CJK_FAMILY);
    expect(face?.bytes.length).toBeGreaterThan(1_000_000);
    expect(String(await importCjkFace())).toEndWith('NotoSansCJKjp-Regular.otf');
  });

  test('East Asian text in a missing font draws with the face, with no missing glyphs', async () => {
    const result = await exportPdf(deck('日本語のテキスト', 'Yu Gothic'), {
      fidelityPolicy: 'best-effort',
    });
    const codes = result.diagnostics.map((d) => d.code);
    expect(codes).not.toContain('text.missing-glyph');
    expect(result.fontResolution.faces.some((f) => f.family === CJK_FAMILY)).toBe(true);
  });
});
