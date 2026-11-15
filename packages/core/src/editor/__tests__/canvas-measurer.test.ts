import { expect, test } from 'bun:test';
import { createCanvasMeasurer, cssFamily } from '../canvas-measurer.ts';

/** A canvas context whose characters are as wide as their code point modulo 7, plus 1. */
function fakeContext() {
  const fonts: string[] = [];
  let font = '';
  return {
    fonts,
    measured: [] as string[],
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
      fonts.push(value);
    },
    measureText(text: string) {
      this.measured.push(text);
      const cp = text.codePointAt(0) ?? 0;
      return { width: ((cp % 7) + 1) * 10, fontBoundingBoxAscent: 92, fontBoundingBoxDescent: 23 };
    },
  };
}

test('a family becomes one quoted CSS family name', () => {
  expect(cssFamily('Calibri')).toBe('"Calibri"');
  expect(cssFamily('Bad"; color: red')).toBe('"Bad\\"; color: red"');
  expect(cssFamily('a\\b')).toBe('"a\\\\b"');
  expect(cssFamily('x\u0000\n\u009fy')).toBe('"xy"');
  expect(cssFamily(' \t')).toBe('sans-serif');
  expect(cssFamily('A'.repeat(500))).toHaveLength(130);
});

test('the canvas measures each character in the stack that paint uses', () => {
  const ctx = fakeContext();
  const m = createCanvasMeasurer(ctx)!;
  const face = m.resolve({ family: 'Calibri', bold: true, italic: false });
  expect(face).toMatchObject({ id: '"Calibri"', synthBold: true, synthItalic: false });
  expect(face.metrics.ascent).toBeCloseTo(0.92);
  expect(face.metrics.descent).toBeCloseTo(0.23);
  // Paint draws `${face.id}, sans-serif` with the weight the face asks for.
  expect(ctx.fonts.at(-1)).toBe('bold 100px "Calibri", sans-serif');
  const text = 'ab😀a';
  const advances = m.advances(face, text);
  const width = (s: string) => (((s.codePointAt(0) ?? 0) % 7) + 1) / 10;
  expect([...advances]).toEqual([width('a'), width('b'), width('😀'), 0, width('a')]);
  // Each character is measured once per font.
  expect(ctx.measured.filter((t) => t === 'a')).toHaveLength(1);
  expect(m.resolve({ family: 'Calibri', bold: true, italic: false })).toBe(face);
});

test('a face box the canvas does not report falls back to the default metrics', () => {
  const ctx = { font: '', measureText: () => ({ width: 50 }) };
  const face = createCanvasMeasurer(ctx)!.resolve({ family: 'X', bold: false, italic: true });
  expect(face.metrics).toMatchObject({ ascent: 0.9, descent: 0.25 });
  expect(face.synthItalic).toBe(true);
});

test('no canvas, no measurer', () => {
  expect(createCanvasMeasurer(null)).toBeNull();
});
