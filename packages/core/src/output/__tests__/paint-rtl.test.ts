import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide } from '../../layout/index.ts';
import { paintSlide } from '../index.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

function paint(text: string): SVGSVGElement {
  const shapes = shape({ id: 2, x: 0, y: 0, cx: 6_000_000, cy: 2_000_000, text });
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
  return paintSlide(doc, layout, { images: () => undefined, idPrefix: 't' });
}

const arabic = (t: string): string =>
  `<a:r><a:rPr lang="ar-SA" sz="1800"/><a:t xml:space="preserve">${t}</a:t></a:r>`;

describe('right-to-left paint', () => {
  test('each right-to-left word is one text element that reads right to left', () => {
    const root = paint(para(arabic('سلام دار'), '<a:pPr rtl="1" algn="r"/>'));
    const texts = [...root.querySelectorAll('text')];
    expect(texts.map((t) => t.textContent)).toEqual(['سلام', 'دار']);
    for (const t of texts) {
      expect(t.getAttribute('style')).toContain('direction:rtl');
      expect(t.getAttribute('style')).toContain('unicode-bidi:bidi-override');
      // One anchor, at the word's right edge, so the browser shapes the whole word.
      expect(t.getAttribute('x')?.split(' ').length).toBe(1);
    }
    // The first word in source order stands to the right of the second.
    const [first, second] = texts.map((t) => Number(t.getAttribute('x')));
    expect(first).toBeGreaterThan(second);
  });

  test('a right-to-left word fills the width layout measured, character spacing included', () => {
    // 24 pt text in the fixed measurer: 12 pt letters, 3 pt closer with spc="-300".
    const word = '\u05d1\u05e2\u05d1\u05e8\u05d9\u05ea';
    const root = paint(
      para(
        `<a:r><a:rPr lang="he-IL" sz="2400" spc="-300"/><a:t>${word}</a:t></a:r>`,
        '<a:pPr rtl="1"/>'
      )
    );
    const text = root.querySelector('text')!;
    expect(Number(text.getAttribute('textLength'))).toBeCloseTo(6 * (12 - 3), 6);
    expect(text.getAttribute('lengthAdjust')).toBe('spacing');
  });

  test('a right-to-left word kerns only when its run asks, as the measurer does', () => {
    const word = '\u05e9\u05dc\u05d5\u05dd';
    const style = (attrs: string): string =>
      paint(
        para(
          `<a:r><a:rPr lang="he-IL" sz="2400"${attrs}/><a:t>${word}</a:t></a:r>`,
          '<a:pPr rtl="1"/>'
        )
      )
        .querySelector('text')!
        .getAttribute('style')!;
    expect(style(' kern="1200"')).toContain('font-kerning:normal');
    // Text smaller than `kern` is not kerned.
    expect(style(' kern="4800"')).toContain('font-kerning:none');
    // Contextual alternates stay on, as in the measurer; ligatures stay off.
    expect(style('')).not.toContain('"calt" 0');
    expect(style('')).toContain('"liga" 0');
  });

  test('left-to-right text keeps one position for each character', () => {
    const root = paint(para('<a:r><a:rPr lang="en-US" sz="1800"/><a:t>abc</a:t></a:r>'));
    const text = root.querySelector('text')!;
    expect(text.getAttribute('x')?.split(' ').length).toBe(3);
    expect(text.getAttribute('style')).not.toContain('direction:rtl');
  });
});
