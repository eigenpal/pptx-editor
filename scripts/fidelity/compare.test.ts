import { describe, expect, test } from 'bun:test';
import { compareLines, comparePixels, lineKey, slidePasses } from './compare.ts';
import { groupLines, type PdfLine, type TextItem } from './pdf.ts';

const line = (text: string, baseline = 100, xMin = 50): PdfLine => ({
  text,
  baseline,
  xMin,
  xMax: xMin + 10 * text.length,
});

describe('lineKey', () => {
  test('drops whitespace and a leading bullet, and skips lines it cannot check', () => {
    expect(lineKey('•  The quick  fox')).toBe('Thequickfox');
    expect(lineKey('1. First')).toBe('1.First');
    // Right-to-left text, unmapped glyphs, numbers, and single glyphs are checked by pixels.
    expect(lineKey('The word سلام')).toBe('');
    expect(lineKey('Arial \u0000')).toBe('');
    expect(lineKey('2.5')).toBe('');
    expect(lineKey('S S')).toBe('');
    expect(lineKey('日本')).toBe('日本');
  });
});

describe('compareLines', () => {
  test('matches the same text at the nearest place and measures the baseline', () => {
    const r = compareLines(
      [line('Total', 300), line('Total', 100.3), line('Extra text')],
      [line('Total', 100)]
    );
    expect(r).toMatchObject({ matched: 1, total: 1, missing: [] });
    expect(r.worstBaseline).toBeCloseTo(0.3, 6);
    expect(r.extra).toEqual(['Total', 'Extra text']);
  });

  test('another line break is a missing line', () => {
    const r = compareLines(
      [line('Anchored bottom', 70)],
      [line('Anchored', 70), line('bottom', 90)]
    );
    expect(r).toMatchObject({ matched: 0, total: 2, missing: ['Anchored', 'bottom'] });
  });

  test('words placed apart on one baseline join to hold a justified line', () => {
    const r = compareLines(
      [line('wizards', 100, 50), line('quietly', 100, 150), line('judge', 100, 250)],
      [line('wizards quietly judge', 100, 50)]
    );
    expect(r).toMatchObject({ matched: 1, total: 1, extra: [] });
  });

  test('a one-letter first word matches without it only when it is a list marker', () => {
    const ours = [{ text: 'Hollow square', xMin: 20, xMax: 120, baseline: 50 }];
    const word = { text: 'a Hollow square', xMin: 10, xMax: 120, baseline: 50 };
    expect(compareLines(ours, [word]).matched).toBe(0);
    expect(compareLines(ours, [{ ...word, marker: true }]).matched).toBe(1);
  });

  test('a matched line must start or end within 0.5 pt of the reference line', () => {
    const ref = [{ text: 'Centered title', xMin: 100, xMax: 200, baseline: 40 }];
    // Another width, the same left edge: the nearer edge counts.
    expect(compareLines([{ ...ref[0], xMax: 210 }], ref).worstShift).toBe(0);
    const moved = compareLines([{ ...ref[0], xMin: 101, xMax: 201 }], ref);
    expect(moved.worstShift).toBe(1);
    expect(slidePasses({ mismatch: 0 }, moved)).toBe(false);
  });

  test('our text that the reference lacks fails the slide', () => {
    const ref = [{ text: 'Title text', xMin: 10, xMax: 90, baseline: 40 }];
    const ours = [...ref, { text: 'Frame label', xMin: 10, xMax: 90, baseline: 80 }];
    const lines = compareLines(ours, ref);
    expect(lines.extra).toEqual(['Frame label']);
    expect(slidePasses({ mismatch: 0 }, lines)).toBe(false);
    expect(slidePasses({ mismatch: 0 }, compareLines(ref, ref))).toBe(true);
  });

  test('a bullet that one PDF writes as a letter matches a symbol bullet', () => {
    const r = compareLines(
      [line('❑ Hollow square')],
      [{ ...line('q Hollow square'), marker: true }]
    );
    expect(r.matched).toBe(1);
  });
});

describe('groupLines', () => {
  const item = (text: string, x: number, y = 100, width = 10 * text.length): TextItem => ({
    text,
    x,
    y,
    width,
    size: 18,
  });

  test('one row splits at a wide gap, and items off the page are left out', () => {
    const lines = groupLines(
      [
        item('The', 10),
        item('quick', 46),
        item('Column', 400),
        item('Off', -50),
        item('Lower', 10, 130),
      ],
      960,
      540
    );
    expect(lines.map((l) => l.text)).toEqual(['The quick', 'Column', 'Lower']);
    expect(lines[0]).toMatchObject({ xMin: 10, xMax: 96, baseline: 100 });
  });
});

describe('comparePixels', () => {
  const raster = (fill: number, mark?: number) => {
    const data = new Uint8Array(4 * 4 * 4).fill(fill);
    if (mark !== undefined) data.set([0, 0, 0, 255], mark * 4);
    for (let i = 3; i < data.length; i += 4) data[i] = 255;
    return { width: 4, height: 4, data };
  };

  test('counts the differing pixels and boxes them', () => {
    const r = comparePixels(raster(255), raster(255, 5), true);
    expect(r.mismatch).toBe(1 / 16);
    expect(r.diffBox).toEqual([1, 1, 2, 2]);
    expect(r.diff?.subarray(1, 4).toString()).toBe('PNG');
    expect(comparePixels(raster(255), raster(255), true)).toEqual({
      mismatch: 0,
      diff: undefined,
      diffBox: undefined,
    });
  });
});
