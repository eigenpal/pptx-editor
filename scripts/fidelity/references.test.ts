import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readdirSync, readFileSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pdfBaseline } from './compare.ts';
import { THRESHOLDS } from './thresholds.ts';
import { PNG } from 'pngjs';
import {
  decode,
  hasPoppler,
  pdfPageFonts,
  pruneRasterCache,
  rasterize,
  rasterizePages,
  readPpm,
  expectedRasterSizePx,
  rasterOptions,
  renderSizePx,
  rasterLimitMs,
  timedOut,
  withBaselines,
  type PdfLine,
} from './references.ts';

describe('decode', () => {
  test('decodes the five XML entities', () => {
    expect(decode('a &lt;b&gt; &quot;c&quot; &#39;d&#39; &amp; e')).toBe(`a <b> "c" 'd' & e`);
  });

  test('decodes once: an escaped entity stays an entity', () => {
    expect(decode('&amp;lt;')).toBe('&lt;');
    expect(decode('&amp;amp;')).toBe('&amp;');
  });
});

describe('withBaselines', () => {
  const line: PdfLine = { text: 'Body text', xMin: 100, xMax: 200, yMin: 40, yMax: 58 };

  test('takes the origin that starts the line', () => {
    const [[out]] = withBaselines(
      [[line]],
      [
        [
          { text: 'x', x: 150, y: 53.04, upsideDown: false },
          { text: 'x', x: 100, y: 54.48, upsideDown: false },
        ],
      ]
    );
    expect(out.baseline).toBe(54.48);
    expect(pdfBaseline(out)).toBe(54.48);
  });

  test('text turned 180° starts at the right edge', () => {
    const [[out]] = withBaselines(
      [[line]],
      [
        [
          { text: 'x', x: 100, y: 44, upsideDown: true },
          { text: 'x', x: 200, y: 43.2, upsideDown: true },
        ],
      ]
    );
    expect(out.baseline).toBe(43.2);
  });

  test('overlapping boxes: the origin near the estimated baseline wins', () => {
    // Tight lines: the second box (48.4–70.4) also holds the first line's origin (50.64).
    const second: PdfLine = { text: 'b', xMin: 43.2, xMax: 340, yMin: 48.38, yMax: 70.36 };
    const [[out]] = withBaselines(
      [[second]],
      [
        [
          { text: 'x', x: 43.2, y: 50.64, upsideDown: false },
          { text: 'x', x: 43.2, y: 65.52, upsideDown: false },
        ],
      ]
    );
    expect(out.baseline).toBe(65.52);
  });

  test('a line that starts with raised text takes the baseline of its other runs', () => {
    // Box 80–100: the estimate is 96. "x²…": the raised "2" starts at 92, the text at 96.12.
    const [[out]] = withBaselines(
      [[{ text: '2 and x', xMin: 100, xMax: 200, yMin: 80, yMax: 100 }]],
      [
        [
          { text: 'x', x: 100, y: 92, upsideDown: false },
          { text: 'x', x: 106, y: 96.12, upsideDown: false },
        ],
      ]
    );
    expect(out.baseline).toBe(96.12);
  });

  test('a lowered run does not move the line: the baseline of the line start wins', () => {
    // `font-effects`, "Subscript H2O, offset -40% C6": the lowered "2" and "6" make the box
    // taller, so the box estimate (70.44) is nearer the "2" (69.36) than the line start (66.48).
    const [[out]] = withBaselines(
      [
        [
          {
            text: 'Subscript H 2 O, offset -40% C 6',
            xMin: 36,
            xMax: 250,
            yMin: 49.34,
            yMax: 75.71,
          },
        ],
      ],
      [
        [
          { text: 'Subscript H', x: 36, y: 66.48, upsideDown: false },
          { text: '2', x: 118.95, y: 69.36, upsideDown: false },
          { text: 'O, offset', x: 125.08, y: 66.48, upsideDown: false },
          { text: '-', x: 191.28, y: 66.48, upsideDown: false },
          { text: '40% C', x: 196.78, y: 66.48, upsideDown: false },
          { text: '6', x: 241.66, y: 72.48, upsideDown: false },
        ],
      ]
    );
    expect(out.baseline).toBe(66.48);
  });

  test('the run of the line above is not the line’s own text', () => {
    // Tight lines: the box also holds the line above's long run; the line's own run wins.
    const [[out]] = withBaselines(
      [[{ text: 'over the lazy dog', xMin: 43.2, xMax: 340, yMin: 48.38, yMax: 70.36 }]],
      [
        [
          { text: 'The quick brown fox jumps', x: 43.2, y: 50.64, upsideDown: false },
          { text: 'over the lazy dog', x: 43.2, y: 65.52, upsideDown: false },
        ],
      ]
    );
    expect(out.baseline).toBe(65.52);
  });

  test('an origin outside the box is not used; the box estimate stays', () => {
    const [[out]] = withBaselines([[line]], [[{ text: 'x', x: 100, y: 80, upsideDown: false }]]);
    expect(out.baseline).toBeUndefined();
    expect(pdfBaseline(out)).toBeCloseTo(58 - 18 * 0.2, 6);
  });
});

// An A4 slide: 10693400 × 7562850 EMU, 842 × 595.5 pt. The reference PDF puts it on an
// 842 × 596 pt page.
const A4 = { widthEmu: 10_693_400, heightEmu: 7_562_850 };

describe('rasterOptions', () => {
  test('renders at the render scale and crops to the slide', () => {
    expect(renderSizePx(A4, 1920)).toEqual({ width: 1920, height: 1358 });
    expect(rasterOptions(A4, 1920)).toEqual([
      '-cropbox',
      '-r',
      '164.18052256532067',
      '-x',
      '0',
      '-y',
      '0',
      '-W',
      '1920',
      '-H',
      '1358',
    ]);
  });

  test('a 16:9 slide keeps whole pixels', () => {
    expect(rasterOptions({ widthEmu: 12_192_000, heightEmu: 6_858_000 }, 1920)).toEqual([
      '-cropbox',
      '-r',
      '144',
      '-x',
      '0',
      '-y',
      '0',
      '-W',
      '1920',
      '-H',
      '1080',
    ]);
  });

  test('at a raster scale, the resolution and the crop grow by the scale', () => {
    expect(rasterOptions({ widthEmu: 12_192_000, heightEmu: 6_858_000 }, 1920, 4)).toEqual([
      '-cropbox',
      '-r',
      '576',
      '-x',
      '0',
      '-y',
      '0',
      '-W',
      '7680',
      '-H',
      '4320',
    ]);
  });
});

/** A one-page PDF of `w` × `h` pt whose content stream is `content`, with a correct xref table. */
function onePagePdf(w: number, h: number, content: string): string {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Contents 4 0 R >>`,
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
  ];
  let out = '%PDF-1.4\n';
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const o of offsets) out += `${String(o).padStart(10, '0')} 00000 n \n`;
  return `${out}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
}

describe.skipIf(!hasPoppler())('rasterize', () => {
  test('an A4 reference page gives the slide area, without the 0.5 pt strip', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      // As the reference PDF draws it: the slide fills y 0.5 to 596 (PDF y points up); y 0 to 0.5
      // is not part of the slide.
      const file = join(dir, 'a4.pdf');
      writeFileSync(file, onePagePdf(842, 596, '0 0 0 rg 0 0.5 842 595.5 re f'));
      const [page] = rasterize(file, join(dir, 'out'), A4, 1920);
      const png = PNG.sync.read(readFileSync(page));
      expect([png.width, png.height]).toEqual([1920, 1358]);
      // The bottom row is slide (dark, allowing for anti-aliasing), not the white strip.
      const last = ((png.height - 1) * png.width + png.width / 2) * 4;
      expect(Math.max(...png.data.subarray(last, last + 3))).toBeLessThan(64);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe.skipIf(!hasPoppler())('pdfPageFonts', () => {
  test('lists the fonts of one page without the subset prefix', () => {
    const fonts = pdfPageFonts(join(import.meta.dir, '../../fixtures/text-runs.pdf'), 1);
    expect(fonts).toContain('Calibri');
    expect(fonts.every((f) => !/^[A-Z]{6}\+/.test(f))).toBe(true);
  });
});

// The `slide-size-rounding` fixture, 400.25 × 300.5 pt: the reference PDF makes it a 400 × 301 pt
// page.
const ROUNDING = { widthEmu: 5_083_175, heightEmu: 3_816_350 };
// 73.75 × 295.25 pt on a 74 × 295 pt page. 295 × dpi / 72 is 7680 in floating point, but
// 295 × (dpi / 72) is just above it: `pdftoppm` makes 7680 px, so the order of operations counts.
const FLOAT_EDGE = { widthEmu: 936_625, heightEmu: 3_749_675 };

describe('the expected raster size', () => {
  test('is the render size, cut where the reference PDF rounded a page side down', () => {
    expect(renderSizePx(ROUNDING, 1920)).toEqual({ width: 1920, height: 1441 });
    // The page ends 0.25 pt before the slide's right edge.
    expect(expectedRasterSizePx(ROUNDING, [400, 301], 1920)).toEqual({ width: 1919, height: 1441 });
    // A side rounded up (A4: 595.5 → 596) keeps the render size.
    expect(expectedRasterSizePx(A4, [842, 596], 1920)).toEqual({ width: 1920, height: 1358 });
    expect(expectedRasterSizePx(FLOAT_EDGE, [74, 295], 1920)).toEqual({
      width: 1920,
      height: 7680,
    });
  });

  test('the render height rounds as the harness page does', () => {
    // 1920 × 1011600 / 921600 = 2107.5 exactly, which rounds up.
    expect(renderSizePx({ widthEmu: 921_600, heightEmu: 1_011_600 }, 1920)).toEqual({
      width: 1920,
      height: 2108,
    });
  });
});

describe.skipIf(!hasPoppler())('the raster of a page that the reference PDF rounded down', () => {
  const rasterSize = (slide: typeof ROUNDING, page: [number, number]): number[] => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      const file = join(dir, 'down.pdf');
      writeFileSync(file, onePagePdf(page[0], page[1], `0 0 0 rg 0 0 ${page[0]} ${page[1]} re f`));
      const png = PNG.sync.read(readFileSync(rasterize(file, join(dir, 'out'), slide, 1920)[0]));
      return [png.width, png.height];
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  };

  test('has exactly the expected size, as the gate requires', () => {
    for (const [slide, page] of [
      [ROUNDING, [400, 301]],
      [FLOAT_EDGE, [74, 295]],
    ] as const) {
      const expected = expectedRasterSizePx(slide, page, 1920);
      expect(rasterSize(slide, [...page])).toEqual([expected.width, expected.height]);
    }
    // A 1920 × 7680 raster: more than 30 s while the parallel suite runs every other file beside it.
  }, 120_000);
});

/** The pixel at `x`, `y` as `rrggbb`. */
const pixelAt = (png: PNG, x: number, y: number): string => {
  const i = (y * png.width + x) * 4;
  return [0, 1, 2].map((c) => png.data[i + c].toString(16).padStart(2, '0')).join('');
};
/** Whether two `rrggbb` colors are within 12 per channel. */
const near = (rgb: string, want: string): boolean =>
  [0, 2, 4].every(
    (i) => Math.abs(parseInt(rgb.slice(i, i + 2), 16) - parseInt(want.slice(i, i + 2), 16)) < 12
  );

describe.skipIf(!hasPoppler())('the slide-size fixture', () => {
  test('its reference page puts the slide at scale 1 from the top-left corner', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      const file = join(import.meta.dir, '../../fixtures/slide-size.pdf');
      const [page] = rasterize(file, join(dir, 'out'), A4, 1920);
      const png = PNG.sync.read(readFileSync(page));
      expect([png.width, png.height]).toEqual([1920, 1358]);
      const at = (x: number, y: number): string => pixelAt(png, x, y);
      // The fixture: a green square at the top-left, a blue bar on the bottom edge, an orange bar
      // on the right edge. A shift moves these colors off the edges.
      expect(near(at(0, 0), '70ad47')).toBe(true);
      expect(near(at(960, png.height - 1), '1f4e79')).toBe(true);
      expect(near(at(png.width - 1, 600), 'ed7d31')).toBe(true);
      // The bar's top is at 1300.5 px. The page squeezed into 1358 px puts it at 1299.4 px, and a
      // slide stretched over the 596 pt page at 1301.6 px: either one changes row 1299 or 1301.
      expect(near(at(960, 1299), 'ffffff')).toBe(true);
      expect(near(at(960, 1301), '1f4e79')).toBe(true);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe.skipIf(!hasPoppler())('the slide-size-rounding fixture', () => {
  test('its reference page ends before the slide, at scale 1', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      const file = join(import.meta.dir, '../../fixtures/slide-size-rounding.pdf');
      const [page] = rasterize(file, join(dir, 'out'), ROUNDING, 1920);
      const png = PNG.sync.read(readFileSync(page));
      // 1 px short: the 400 pt page holds 1918.8 px of the 1920 px slide.
      expect([png.width, png.height]).toEqual([1919, 1441]);
      expect(near(pixelAt(png, 0, 0), '70ad47')).toBe(true);
      expect(near(pixelAt(png, 900, png.height - 1), '1f4e79')).toBe(true);
      // The orange bar starts at 375.05 pt, 1799.1 px. The 400.25 pt slide squeezed into 1919 px
      // would put it at 1798.2 px, so column 1798 is white only at scale 1.
      expect(near(pixelAt(png, 1798, 600), 'ffffff')).toBe(true);
      expect(near(pixelAt(png, 1800, 600), 'ed7d31')).toBe(true);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('the raster time limit', () => {
  test('grows with the page count', () => {
    expect(rasterLimitMs(1, 1)).toBe(125_000);
    expect(rasterLimitMs(60, 1)).toBe(420_000);
  });

  test('grows with the pixels of the raster scale', () => {
    expect(rasterLimitMs(1, 4)).toBe(200_000);
    expect(rasterLimitMs(1)).toBe(rasterLimitMs(1, THRESHOLDS.rasterScale));
  });

  test.skipIf(!hasPoppler())('a raster that runs out of time is a time-out', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      const pdf = join(import.meta.dir, '../../fixtures/slide-structure.pdf');
      let error: unknown;
      try {
        rasterize(pdf, join(dir, 'out'), { widthEmu: 12_192_000, heightEmu: 6_858_000 }, 1920, 1);
      } catch (e) {
        error = e;
      }
      expect(timedOut(error)).toBe(true);
      expect(timedOut(new Error('pdftoppm failed'))).toBe(false);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('readPpm', () => {
  test('gives opaque RGBA pixels, as a decoded RGB PNG does', () => {
    const header = new TextEncoder().encode('P6\n2 1\n255\n');
    const bytes = new Uint8Array([...header, 1, 2, 3, 250, 251, 252]);
    expect(readPpm(bytes)).toEqual({
      width: 2,
      height: 1,
      data: new Uint8Array([1, 2, 3, 255, 250, 251, 252, 255]),
    });
  });

  test('refuses a file that is not an 8-bit binary PPM, or is short', () => {
    const enc = (t: string): Uint8Array => new TextEncoder().encode(t);
    expect(() => readPpm(enc('P3\n1 1\n255\n'))).toThrow('not a binary PPM');
    expect(() => readPpm(enc('P6\n1 1\n65535\n'))).toThrow('is not 8 bits');
    expect(() => readPpm(enc('P6\n2 2\n255\nabc'))).toThrow('the PPM is short');
  });
});

describe.skipIf(!hasPoppler())('the reference raster cache', () => {
  test('a cached raster has the same pixels as a new one, and a full run prunes unused entries', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-raster-'));
    try {
      const file = join(import.meta.dir, '../../fixtures/slide-size.pdf');
      const cache = join(dir, 'cache');
      const fresh = rasterizePages(file, join(dir, 'a'), A4, 480, undefined, 4, cache);
      expect(readdirSync(cache)).toHaveLength(1);
      const cached = rasterizePages(file, join(dir, 'b'), A4, 480, undefined, 4, cache);
      expect(Buffer.from(cached[0].image.data).equals(Buffer.from(fresh[0].image.data))).toBe(true);
      expect(Buffer.from(cached[0].device().data).equals(Buffer.from(fresh[0].device().data))).toBe(
        true
      );
      // The same PNG file as without a cache.
      const plain = rasterize(file, join(dir, 'c'), A4, 480, undefined, 4);
      expect(readFileSync(cached[0].path).equals(readFileSync(plain[0]))).toBe(true);
      // An entry that no run used since a time is removed; a used one stays.
      const [entry] = readdirSync(cache);
      utimesSync(join(cache, entry, 'meta.json'), new Date(1000), new Date(1000));
      pruneRasterCache(cache, Date.now() - 60_000);
      expect(readdirSync(cache)).toHaveLength(0);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
