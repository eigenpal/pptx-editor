import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import {
  createHarfBuzzMeasurer,
  EAST_ASIAN_FALLBACK_FAMILIES,
  eastAsianFallbackFamilies,
  HOST_SUBSTITUTE_ORIGIN,
  initializeHarfBuzz,
  type TextMeasurer,
} from '../../index.ts';
import { AdvanceCache } from '../advance-cache.ts';
import { os2LineExtents } from '../face-line-metrics.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);

const liberation = () => ({
  family: 'Liberation Sans',
  bold: false,
  italic: false,
  bytes: new Uint8Array(readFileSync(join(fontsDir, 'LiberationSans-Regular.ttf'))),
  origin: 'substitute',
});

let measurer: TextMeasurer;
beforeAll(async () => {
  await initializeHarfBuzz();
  measurer = createHarfBuzzMeasurer([liberation()]);
});

describe('the HarfBuzz measurer (P7, P8)', () => {
  test('one reused buffer measures exactly like a fresh measurer', () => {
    const fresh = createHarfBuzzMeasurer([liberation()]);
    const face = measurer.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    const freshFace = fresh.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    // Right-to-left and East Asian text first: no buffer state may leak into the next text.
    measurer.advances(face, 'ﬁ العربية 日本語');
    for (const kerning of [false, true]) {
      const reused = Array.from(measurer.advances(face, 'AVATAR Wave', { kerning }));
      expect(reused).toEqual(Array.from(fresh.advances(freshFace, 'AVATAR Wave', { kerning })));
    }
    const plain = measurer.advances(face, 'AVATAR Wave');
    const kerned = measurer.advances(face, 'AVATAR Wave', { kerning: true });
    expect(kerned[0] + kerned[1]).toBeLessThan(plain[0] + plain[1]);
  });

  test('two measurers never share a face id, so two documents never collide', () => {
    const a = createHarfBuzzMeasurer([liberation(), liberation()]);
    const b = createHarfBuzzMeasurer([liberation()]);
    const ids = [...a.faces, ...b.faces].map((face) => face.id);
    expect(new Set(ids).size).toBe(3);
    for (const id of ids) expect(id).toMatch(/^pptx-f\d+-\d+$/);
  });

  test('text in use stays measured while more than the cache limit passes through', () => {
    const face = measurer.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    const hot = measurer.advances(face, 'The text being edited');
    const first = measurer.advances(face, 'other run 0');
    // 30 000 other strings, more than the 20 000-entry cache, with the hot text in use.
    for (let i = 0; i < 30_000; i += 1) {
      measurer.advances(face, `other run ${i}`);
      if (i % 1000 === 0) expect(measurer.advances(face, 'The text being edited')).toBe(hot);
    }
    expect(measurer.advances(face, 'The text being edited')).toBe(hot);
    // The least recently used entries went first: the first other run was measured again.
    expect(measurer.advances(face, 'other run 0')).not.toBe(first);
    expect(measurer.advances(face, 'other run 29999')).toBe(
      measurer.advances(face, 'other run 29999')
    );
  });
});

describe('the advance cache', () => {
  test('holds at most its limit and keeps entries in use', () => {
    const cache = new AdvanceCache<number>(100);
    cache.set('hot', 1);
    for (let i = 0; i < 1000; i += 1) {
      cache.set(`k${i}`, i);
      expect(cache.size).toBeLessThanOrEqual(100);
      if (i % 40 === 0) expect(cache.get('hot')).toBe(1);
    }
    expect(cache.get('hot')).toBe(1);
    expect(cache.get('k0')).toBeUndefined();
    expect(cache.get('k999')).toBe(999);
  });

  test('holds at most its weight limit, as typing in a long paragraph adds heavy entries', () => {
    const cache = new AdvanceCache<number>(20_000, 10_000);
    cache.set('hot', 0, 10);
    // Each keystroke measures the paragraph again: 3000 entries of about 2000 characters.
    for (let i = 0; i < 3000; i += 1) {
      cache.set(`p${i}`, i, 2000 + i);
      expect(cache.weight).toBeLessThanOrEqual(10_000);
      expect(cache.get('hot')).toBe(0);
    }
    expect(cache.get('p2999')).toBe(2999);
    expect(cache.get('p0')).toBeUndefined();
    // An entry heavier than half the limit is not kept.
    cache.set('huge', 1, 6000);
    expect(cache.get('huge')).toBeUndefined();
    expect(cache.get('p2999')).toBe(2999);
  });
});

describe('face matching', () => {
  const face = (file: string, bold: boolean, italic: boolean) => ({
    family: 'Deck Sans',
    bold,
    italic,
    bytes: new Uint8Array(readFileSync(join(fontsDir, file))),
    origin: 'embedded',
  });

  test('a regular run with no regular face gets an upright face, not the italic one', () => {
    const m = createHarfBuzzMeasurer([
      face('LiberationSans-Bold.ttf', true, false),
      face('LiberationSans-Italic.ttf', false, true),
      face('LiberationSans-BoldItalic.ttf', true, true),
    ]);
    const regular = m.resolve({ family: 'Deck Sans', bold: false, italic: false });
    expect(regular.italic).toBe(false);
    expect(regular.synthItalic).toBe(false);
    const italic = m.resolve({ family: 'Deck Sans', bold: false, italic: true });
    expect(italic.italic).toBe(true);
    expect(italic.bold).toBe(false);
  });
});

describe('line metrics', () => {
  test('a face with fsSelection bit 7 measures its lines with its typo values, not win', () => {
    // TeX Gyre Adventor: bit 7 set, typo 739 and 192, win 1212 and 505 (1000 units).
    const adventor = {
      family: 'TeX Gyre Adventor',
      bold: false,
      italic: false,
      bytes: new Uint8Array(readFileSync(join(fontsDir, 'TeXGyreAdventor-Regular.otf'))),
      origin: 'substitute',
    };
    const m = createHarfBuzzMeasurer([adventor]).resolve({
      family: 'TeX Gyre Adventor',
      bold: false,
      italic: false,
    }).metrics;
    expect(m.ascent).toBeCloseTo(0.739, 6);
    expect(m.descent).toBeCloseTo(0.192, 6);
  });

  test('a missing face keeps its own line metrics when the table knows them', () => {
    const yu = measurer.resolve({ family: 'Yu Gothic', bold: false, italic: false });
    expect(yu.substituted).toBe(true);
    expect(yu.family).toBe('Liberation Sans');
    expect(yu.metrics.ascent).toBeCloseTo(2017 / 2048, 6);
    expect(yu.metrics.descent).toBeCloseTo(619 / 2048, 6);
    // A family the table does not know measures with the face that draws it.
    const other = measurer.resolve({ family: 'Unknown Deck Face', bold: false, italic: false });
    expect(other.metrics.ascent).toBeCloseTo(1854 / 2048, 6);
    expect(other.metrics.descent).toBeCloseTo(434 / 2048, 6);
  });
});

describe('the OS/2 line extents', () => {
  /** An OS/2 table of 78 bytes with fsSelection, typo, and win values. */
  const os2 = (fsSelection: number, typo: [number, number], win: [number, number]) => {
    const bytes = new Uint8Array(78);
    const v = new DataView(bytes.buffer);
    v.setUint16(62, fsSelection);
    v.setInt16(68, typo[0]);
    v.setInt16(70, -typo[1]);
    v.setUint16(74, win[0]);
    v.setUint16(76, win[1]);
    return bytes;
  };

  test('fsSelection bit 7 selects the typo values, else the win values', () => {
    // Aptos: bit 7 set, typo 1923 and 577, win 2068 and 563.
    expect(os2LineExtents(os2(0x80 | 0x40, [1923, 577], [2068, 563]))).toEqual([1923, 577]);
    // Yu Gothic: bit 7 clear, typo 1802 and 246, win 2017 and 619.
    expect(os2LineExtents(os2(0x40, [1802, 246], [2017, 619]))).toEqual([2017, 619]);
  });

  test('a short or empty table gives nothing', () => {
    expect(os2LineExtents(undefined)).toBeUndefined();
    expect(os2LineExtents(new Uint8Array(70))).toBeUndefined();
    expect(os2LineExtents(os2(0, [0, 0], [0, 0]))).toBeUndefined();
  });
});

describe('East Asian fallback faces', () => {
  const cjk = () => ({ ...liberation(), family: 'Noto Sans CJK JP', origin: 'host' });

  test('East Asian text whose family is missing takes a host East Asian face', () => {
    const m = createHarfBuzzMeasurer([liberation(), cjk()]);
    const face = m.resolve({
      family: 'Yu Gothic',
      bold: false,
      italic: false,
      script: 'eastAsian',
    });
    expect(face.family).toBe('Noto Sans CJK JP');
    expect(face.substituted).toBe(true);
    // Latin text with the same missing family keeps the default fallback.
    expect(m.resolve({ family: 'Yu Gothic', bold: false, italic: false }).family).toBe(
      'Liberation Sans'
    );
  });

  test("the run's language orders the East Asian faces", () => {
    const region = (family: string) => ({ ...liberation(), family, origin: 'host' });
    const m = createHarfBuzzMeasurer([
      liberation(),
      region('Noto Sans CJK JP'),
      region('Noto Sans SC'),
      region('Noto Sans TC'),
      region('Noto Sans KR'),
    ]);
    const family = (lang?: string) =>
      m.resolve({
        family: 'Missing Face',
        bold: false,
        italic: false,
        script: 'eastAsian',
        ...(lang ? { lang } : {}),
      }).family;
    expect(family('zh-CN')).toBe('Noto Sans SC');
    expect(family('zh-TW')).toBe('Noto Sans TC');
    expect(family('ko-KR')).toBe('Noto Sans KR');
    expect(family('ja-JP')).toBe('Noto Sans CJK JP');
    expect(family()).toBe('Noto Sans CJK JP');
    expect(eastAsianFallbackFamilies('zh-HK')[0]).toBe('Noto Sans CJK TC');
    expect(eastAsianFallbackFamilies('en-US')).toBe(EAST_ASIAN_FALLBACK_FAMILIES);
  });

  test('without a host East Asian face, East Asian text takes the default fallback', () => {
    const face = measurer.resolve({
      family: 'Yu Gothic',
      bold: false,
      italic: false,
      script: 'eastAsian',
    });
    expect(face.family).toBe('Liberation Sans');
  });
});

describe('glyph fallback', () => {
  const face = (file: string, family: string) => ({
    family,
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, file))),
    origin: 'substitute',
  });

  test('a character the face lacks takes the first loaded face that has it', () => {
    const m = createHarfBuzzMeasurer([
      face('Selawik-Regular.ttf', 'Selawik'),
      liberation(),
      face('Carlito-Regular.ttf', 'Carlito'),
    ]);
    const selawik = m.resolve({ family: 'Selawik', bold: false, italic: false });
    // Selawik has no Cyrillic; Liberation Sans, loaded next, has it.
    expect(m.fallbackFace!(selawik, 0x0416)?.family).toBe('Liberation Sans');
    // Liberation Sans has no circled digits; Carlito has them.
    const sans = m.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    expect(m.fallbackFace!(sans, 0x2460)?.family).toBe('Carlito');
    // A character the face has, or that no face has, needs no fallback.
    expect(m.fallbackFace!(sans, 0x0416)).toBeUndefined();
    expect(m.fallbackFace!(sans, 0x10ffff)).toBeUndefined();
  });
});

describe('kerning', () => {
  const carlito = () => ({
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  });
  const sum = (a: Float64Array): number => a.reduce((s, v) => s + v, 0);

  test('a named face with no kern table draws without kerning', () => {
    // font-effects: "AVATAR WAVE kerned" (kern="1200", 28 pt Carlito) has the advances of the
    // same text with kern="0" in the reference PDF. Carlito keeps its pairs in GPOS only.
    const m = createHarfBuzzMeasurer([carlito()]);
    const face = m.resolve({ family: 'Carlito', bold: false, italic: false });
    expect(face.substituted).toBe(false);
    const kerned = m.advances(face, 'AVATAR WAVE', { kerning: true });
    expect(Array.from(kerned)).toEqual(Array.from(m.advances(face, 'AVATAR WAVE')));
  });

  test('a named face with a kern table kerns', () => {
    const face = measurer.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    expect(sum(measurer.advances(face, 'AVATAR WAVE', { kerning: true }))).toBeLessThan(
      sum(measurer.advances(face, 'AVATAR WAVE'))
    );
  });

  test('a host substitution keeps the kerning pairs of the face it draws', () => {
    // The host registers Carlito under the family that the deck names (`FontSourceSubstitution`):
    // it stands in for Calibri, as the measurer's own substitute does.
    const m = createHarfBuzzMeasurer([
      { ...carlito(), family: 'Calibri', origin: HOST_SUBSTITUTE_ORIGIN },
    ]);
    const face = m.resolve({ family: 'Calibri', bold: false, italic: false });
    expect(face.substituted).toBe(false);
    expect(sum(m.advances(face, 'AVATAR WAVE', { kerning: true }))).toBeLessThan(
      sum(m.advances(face, 'AVATAR WAVE'))
    );
  });

  test('a stand-in for a missing face keeps its own kerning pairs', () => {
    const m = createHarfBuzzMeasurer([carlito()]);
    const face = m.resolve({ family: 'Calibri', bold: false, italic: false });
    expect(face.family).toBe('Carlito');
    expect(face.substituted).toBe(true);
    expect(sum(m.advances(face, 'AVATAR WAVE', { kerning: true }))).toBeLessThan(
      sum(m.advances(face, 'AVATAR WAVE'))
    );
  });
});
