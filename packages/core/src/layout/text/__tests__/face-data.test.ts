// Face data that the HarfBuzz measurer reads from the font bytes: the OS/2 strikeout stroke,
// the cmap coverage of each character, and kerning that leaves full-width East Asian text alone.
import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { buildPptx } from '../../../../../../scripts/fixtures/builder.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import { readPptx } from '../../../store/index.ts';
import {
  createHarfBuzzMeasurer,
  initializeHarfBuzz,
  layoutSlide,
  strikeoutBars,
  type FaceMetrics,
  type FontSource,
  type TextLayout,
  type TextMeasurer,
} from '../../index.ts';
import { unkernedTextRanges } from '../run-kerning.ts';
import { cmapMapsCodePoint, unicodeCmapSubtable } from '../sfnt-cmap-coverage.ts';
import { sfntStrikeoutStrokeEm } from '../sfnt-strikeout-metrics.ts';

const assets = (pkg: string): string =>
  join(dirname(Bun.resolveSync(`${pkg}/package.json`, import.meta.dir)), 'assets');
const fontsDir = assets('@docx-editor.dev/fonts');

const source = (file: string, family: string, dir = fontsDir): FontSource => ({
  family,
  bold: false,
  italic: false,
  bytes: new Uint8Array(readFileSync(join(dir, file))),
  origin: 'substitute',
});

beforeAll(async () => {
  await initializeHarfBuzz();
});

/** Minimal sfnt with a `head` and, unless omitted, an `OS/2` table with a strikeout stroke. */
function strikeSfnt(options: {
  readonly unitsPerEm: number;
  readonly thickness: number;
  readonly position: number;
  readonly omitOs2?: boolean;
}): Uint8Array {
  const records = options.omitOs2 ? 1 : 2;
  const headOffset = 12 + records * 16;
  const os2Offset = headOffset + 54;
  const bytes = new Uint8Array(os2Offset + 96);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, records);
  view.setUint32(12, 0x68656164);
  view.setUint32(12 + 8, headOffset);
  view.setUint32(12 + 12, 54);
  view.setUint16(headOffset + 18, options.unitsPerEm);
  if (!options.omitOs2) {
    view.setUint32(28, 0x4f532f32);
    view.setUint32(28 + 8, os2Offset);
    view.setUint32(28 + 12, 96);
    view.setInt16(os2Offset + 26, options.thickness);
    view.setInt16(os2Offset + 28, options.position);
  }
  return bytes;
}

describe('the OS/2 strikeout stroke', () => {
  test('the reader returns em fractions and refuses faces it cannot trust', () => {
    expect(
      sfntStrikeoutStrokeEm(strikeSfnt({ unitsPerEm: 2048, thickness: 134, position: 512 }), 0)
    ).toEqual({ offsetEm: 512 / 2048, thicknessEm: 134 / 2048 });
    // File-supplied int16 over a file-supplied em: a stroke taller than the em is refused.
    const refused = [
      strikeSfnt({ unitsPerEm: 16, thickness: 32767, position: 0 }),
      strikeSfnt({ unitsPerEm: 2048, thickness: 0, position: 512 }),
      strikeSfnt({ unitsPerEm: 2048, thickness: 134, position: 30000 }),
      strikeSfnt({ unitsPerEm: 2048, thickness: 134, position: 512, omitOs2: true }),
      new Uint8Array(4),
    ];
    for (const bytes of refused) expect(sfntStrikeoutStrokeEm(bytes, 0)).toBeNull();
  });

  test('truncated bytes never read out of bounds', () => {
    const full = strikeSfnt({ unitsPerEm: 2048, thickness: 134, position: 512 });
    for (let length = 0; length < full.length; length += 7)
      expect(() => sfntStrikeoutStrokeEm(full.slice(0, length), 0)).not.toThrow();
  });

  test('the measurer gives each face its own stroke', () => {
    const m = createHarfBuzzMeasurer([source('Carlito-Regular.ttf', 'Carlito')]);
    const metrics = m.resolve({ family: 'Carlito', bold: false, italic: false }).metrics;
    expect(metrics.strikeoutPosition).toBeCloseTo(512 / 2048, 9);
    expect(metrics.strikeoutThickness).toBeCloseTo(134 / 2048, 9);
  });

  const carlito: FaceMetrics = {
    ascent: 0.75,
    descent: 0.25,
    lineGap: 0,
    underlinePosition: -0.05,
    underlineThickness: 0.0947,
    strikeoutPosition: 512 / 2048,
    strikeoutThickness: 134 / 2048,
  };

  test('a single stroke centers on the face position', () => {
    // font-effects and text-runs: 18 pt Carlito draws the stroke from 3.9 to 5.0 pt above
    // the baseline (whole points in the reference PDF), around 4.5 pt.
    const bars = strikeoutBars(carlito, 18, false);
    expect(bars.thickness).toBeCloseTo(18 * (134 / 2048), 9);
    expect(bars.tops).toHaveLength(1);
    expect(bars.tops[0]! + bars.thickness / 2).toBeCloseTo(-4.5, 9);
  });

  test('a double stroke is two bars one thickness apart, on the same center', () => {
    const { tops, thickness } = strikeoutBars(carlito, 18, true);
    expect(tops).toHaveLength(2);
    expect(tops[1]! - (tops[0]! + thickness)).toBeCloseTo(thickness, 9);
    expect((tops[0]! + tops[1]! + thickness) / 2).toBeCloseTo(-4.5, 9);
  });

  test('a face without a stroke keeps the fallback at 0.3 em', () => {
    const { strikeoutPosition: _p, strikeoutThickness: _t, ...plain } = carlito;
    const bars = strikeoutBars(plain, 20, false);
    expect(bars.thickness).toBeCloseTo(20 * 0.0947, 9);
    expect(bars.tops[0]! + bars.thickness / 2).toBeCloseTo(-6, 9);
    const thin = strikeoutBars({ ...plain, underlineThickness: 0.01 }, 20, false);
    expect(thin.thickness).toBeCloseTo(1, 9);
  });
});

/** A one-table sfnt whose cmap holds `subtable` under (platform, encoding). */
function cmapSfnt(subtable: number[], platform = 3, encoding = 1): Uint8Array {
  return cmapSfntOf([{ platform, encoding, subtable }]);
}

/** A one-table sfnt whose cmap lists each subtable under its (platform, encoding), in order. */
function cmapSfntOf(
  entries: readonly { platform: number; encoding: number; subtable: number[] }[]
): Uint8Array {
  const cmap = [0, 0, ...be16(entries.length)];
  let offset = 4 + entries.length * 8;
  for (const { platform, encoding, subtable } of entries) {
    cmap.push(...be16(platform), ...be16(encoding), ...be32(offset));
    offset += subtable.length;
  }
  for (const { subtable } of entries) cmap.push(...subtable);
  const header = [0, 1, 0, 0, 0, 1, 0, 16, 0, 0, 0, 0];
  const record = [0x63, 0x6d, 0x61, 0x70, 0, 0, 0, 0, ...be32(28), ...be32(cmap.length)];
  return new Uint8Array([...header, ...record, ...cmap]);
}
const be16 = (value: number) => [(value >> 8) & 0xff, value & 0xff];
const be32 = (value: number) => [
  (value >>> 24) & 0xff,
  (value >> 16) & 0xff,
  (value >> 8) & 0xff,
  value & 0xff,
];

/** Format 4: U+2010 and U+2013-U+2014 by delta, U+3042 through the glyph array, then 0xFFFF. */
function format4(): number[] {
  const segments = [
    { start: 0x2010, end: 0x2010, delta: 5 - 0x2010, array: false },
    { start: 0x2013, end: 0x2014, delta: 6 - 0x2013, array: false },
    { start: 0x3042, end: 0x3043, delta: 0, array: true },
    { start: 0xffff, end: 0xffff, delta: 1, array: false },
  ];
  const count = segments.length;
  const out = [...be16(4), 0, 0, 0, 0, ...be16(count * 2), 0, 0, 0, 0, 0, 0];
  for (const s of segments) out.push(...be16(s.end));
  out.push(0, 0);
  for (const s of segments) out.push(...be16(s.start));
  for (const s of segments) out.push(...be16(s.delta & 0xffff));
  // The array segment points past the remaining range offsets to two glyph ids: 9, then notdef.
  segments.forEach((s, index) => out.push(...be16(s.array ? (count - index) * 2 : 0)));
  out.push(...be16(9), ...be16(0));
  out[2] = out.length >> 8;
  out[3] = out.length & 0xff;
  return out;
}

describe('cmap coverage', () => {
  test('format 4 maps by delta and through the glyph array, and notdef counts as missing', () => {
    const bytes = cmapSfnt(format4());
    const subtable = unicodeCmapSubtable(bytes, 0)!;
    expect(subtable.format).toBe(4);
    const maps = (cp: number) => cmapMapsCodePoint(bytes, subtable, cp);
    expect([0x2010, 0x2011, 0x2013, 0x2014, 0x3042, 0x3043, 0x1f600].map(maps)).toEqual([
      true,
      false,
      true,
      true,
      true,
      false,
      false,
    ]);
  });

  test('format 12 is preferred and maps supplementary code points', () => {
    const groups = [
      [0x2010, 0x2010, 4],
      [0x1f600, 0x1f601, 10],
    ];
    const sub = [...be16(12), 0, 0, ...be32(16 + groups.length * 12), ...be32(0)];
    sub.push(...be32(groups.length));
    for (const [start, end, glyph] of groups)
      sub.push(...be32(start!), ...be32(end!), ...be32(glyph!));
    const bytes = cmapSfnt(sub, 3, 10);
    const subtable = unicodeCmapSubtable(bytes, 0)!;
    expect(subtable.format).toBe(12);
    expect(cmapMapsCodePoint(bytes, subtable, 0x2010)).toBe(true);
    expect(cmapMapsCodePoint(bytes, subtable, 0x2011)).toBe(false);
    expect(cmapMapsCodePoint(bytes, subtable, 0x1f601)).toBe(true);
  });

  test('a symbol cmap has no Unicode subtable to read', () => {
    expect(unicodeCmapSubtable(cmapSfnt(format4(), 3, 0), 0)).toBeNull();
  });

  /** Format 4 that maps U+2010 only. */
  const hyphenOnly = (): number[] => {
    const out = [...be16(4), 0, 0, 0, 0, ...be16(4), 0, 0, 0, 0, 0, 0];
    out.push(...be16(0x2010), ...be16(0xffff), 0, 0, ...be16(0x2010), ...be16(0xffff));
    out.push(...be16((5 - 0x2010) & 0xffff), ...be16(1), 0, 0, 0, 0);
    out[3] = out.length;
    return out;
  };

  test('(3,1) wins over an earlier (0,3), as HarfBuzz ranks them', () => {
    const bytes = cmapSfntOf([
      { platform: 0, encoding: 3, subtable: hyphenOnly() },
      { platform: 3, encoding: 1, subtable: format4() },
    ]);
    const subtable = unicodeCmapSubtable(bytes, 0)!;
    expect(cmapMapsCodePoint(bytes, subtable, 0x3042)).toBe(true);
  });

  test('a malformed subtable is passed over for the next one', () => {
    // The (3,1) table's segment count runs past the end of the font.
    const broken = format4();
    broken[6] = 0x7f;
    const bytes = cmapSfntOf([
      { platform: 3, encoding: 1, subtable: broken },
      { platform: 0, encoding: 3, subtable: hyphenOnly() },
    ]);
    const subtable = unicodeCmapSubtable(bytes, 0)!;
    expect(cmapMapsCodePoint(bytes, subtable, 0x2010)).toBe(true);
    expect(cmapMapsCodePoint(bytes, subtable, 0x3042)).toBe(false);
  });

  test('truncated or garbage bytes answer missing without reading out of bounds', () => {
    const full = cmapSfnt(format4());
    for (let length = 0; length < full.length; length += 1) {
      const bytes = full.slice(0, length);
      const subtable = unicodeCmapSubtable(bytes, 0);
      if (subtable) expect(typeof cmapMapsCodePoint(bytes, subtable, 0x3042)).toBe('boolean');
    }
    expect(unicodeCmapSubtable(new Uint8Array([0x74, 0x74, 0x63, 0x66, 0, 1, 0, 0]), 3)).toBeNull();
  });

  /** Format 4 with the most segments a table can hold: segment i maps U+(2i) only. */
  function largeFormat4(sorted: boolean): number[] {
    const count = 32767;
    const ends = Array.from({ length: count }, (_, i) =>
      i === count - 1 ? 0xffff : sorted ? i * 2 : 0x20
    );
    const out = [...be16(4), 0, 0, 0, 0, ...be16(count * 2), 0, 0, 0, 0, 0, 0];
    for (const end of ends) out.push(...be16(end));
    out.push(0, 0);
    for (const end of ends) out.push(...be16(end));
    for (let i = 0; i < count; i += 1) out.push(...be16(1));
    for (let i = 0; i < count; i += 1) out.push(0, 0);
    return out;
  }

  test('a format 4 table with 32,767 segments answers the whole BMP in bounded time', () => {
    const bytes = cmapSfnt(largeFormat4(true));
    const subtable = unicodeCmapSubtable(bytes, 0)!;
    expect(subtable.format).toBe(4);
    const started = performance.now();
    let mapped = 0;
    for (let cp = 0; cp < 0x10000; cp += 1) if (cmapMapsCodePoint(bytes, subtable, cp)) mapped += 1;
    // A scan of every segment took over a second here; a binary search takes milliseconds.
    expect(performance.now() - started).toBeLessThan(250);
    // The last segment, U+FFFF, maps to notdef.
    expect(mapped).toBe(32766);
  });

  test('a format 4 table whose segments are out of order is refused, so HarfBuzz answers', () => {
    expect(unicodeCmapSubtable(cmapSfnt(largeFormat4(false)), 0)).toBeNull();
  });

  // Liberation Sans has no U+2011, and HarfBuzz draws it with the face's U+2010.
  const sans = () => source('LiberationSans-Regular.ttf', 'Liberation Sans');
  const adventor = () => source('TeXGyreAdventor-Regular.otf', 'TeX Gyre Adventor');

  test('a character the cmap lacks takes a fallback face, also when shaping hides it', () => {
    const m = createHarfBuzzMeasurer([sans(), adventor()]);
    const face = m.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    expect(m.fallbackFace!(face, 0x2011)?.family).toBe('TeX Gyre Adventor');
    expect(m.fallbackFace!(face, 0x2010)).toBeUndefined();
  });

  test('a character the cmap lacks is measured as missing, not as the glyph shaping chose', () => {
    const m = createHarfBuzzMeasurer([sans()]);
    const face = m.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    // U+226E is not in the cmap; HarfBuzz draws it as '<' and a combining long solidus.
    const lessThan = m.advances(face, 'a<b');
    const notLess = m.advances(face, 'a\u226eb');
    expect(notLess[1]).toBe(0.5);
    expect(notLess[1]).not.toBe(lessThan[1]);
    expect([notLess[0], notLess[2]]).toEqual([lessThan[0], lessThan[2]]);
  });

  test("a missing hyphen or dash is measured at the face's own hyphen advance", () => {
    const m = createHarfBuzzMeasurer([sans()]);
    const face = m.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    const hyphen = m.advances(face, '\u2010')[0]!;
    expect(hyphen).toBeCloseTo(682 / 2048, 9);
    expect(m.advances(face, 'a\u2011b')[1]).toBe(hyphen);
  });

  test('spaces and default ignorable characters keep the advance shaping gives them', () => {
    const m = createHarfBuzzMeasurer([sans()]);
    const face = m.resolve({ family: 'Liberation Sans', bold: false, italic: false });
    // U+2060 WORD JOINER is not in the cmap and draws nothing; U+2007 FIGURE SPACE is in it.
    expect(m.advances(face, 'a⁠b')[1]).toBe(0);
    expect(m.advances(face, ' ')[0]).toBeGreaterThan(0.4);
  });
});

describe('kerning leaves full-width East Asian text alone', () => {
  const cjk = (): FontSource => ({
    ...source(
      'NotoSansCJKjp-Regular.otf',
      'Noto Sans CJK JP',
      assets('@docx-editor.dev/fonts-cjk')
    ),
    origin: 'host',
  });
  let m: TextMeasurer;
  beforeAll(() => {
    m = createHarfBuzzMeasurer([source('LiberationSans-Regular.ttf', 'Liberation Sans'), cjk()]);
  });

  test('the ranges cover kana, ideographs, and full-width forms only', () => {
    expect(unkernedTextRanges('AVアイ To 漢')).toEqual([
      [2, 4],
      [8, 9],
    ]);
    expect(unkernedTextRanges('ＡＶ｡ｱ')).toEqual([[0, 4]]);
    expect(unkernedTextRanges('AVATo 한국어')).toEqual([]);
  });

  test('a stand-in East Asian face keeps every kana advance, and its Latin text kerns', () => {
    // Noto Sans CJK JP stands in for a missing face, so its GPOS pairs apply: it kerns アイ.
    const face = m.resolve({
      family: 'Yu Gothic',
      bold: false,
      italic: false,
      script: 'eastAsian',
    });
    expect(face.substituted).toBe(true);
    expect(Array.from(m.advances(face, 'アイウエオ', { kerning: true }))).toEqual(
      Array.from(m.advances(face, 'アイウエオ'))
    );
    const mixed = m.advances(face, 'アイAVウエ', { kerning: true });
    const plain = m.advances(face, 'アイAVウエ');
    expect(Array.from(mixed.slice(0, 2))).toEqual(Array.from(plain.slice(0, 2)));
    expect(mixed[2]).toBeLessThan(plain[2]!);
    expect(Array.from(mixed.slice(4))).toEqual(Array.from(plain.slice(4)));
  });

  /** The first line of a 6 x 4 in shape whose run has `rPr` attributes, laid out with `m`. */
  function lay(rPr: string, text: string): TextLayout {
    const body =
      `<a:p><a:r><a:rPr lang="ja-JP" ${rPr}><a:latin typeface="Liberation Sans"/>` +
      `<a:ea typeface="Yu Gothic"/></a:rPr><a:t>${text}</a:t></a:r></a:p>`;
    const shapes =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      `<p:txBody><a:bodyPr wrap="none"/><a:lstStyle/>${body}</p:txBody></p:sp>`;
    const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
    if (!r.ok) throw new Error('open failed');
    const item = layoutSlide(resolveSlide(r.pkg, 0)!, m).items[0];
    if (item?.kind !== 'shape' || !item.text) throw new Error('no text');
    return item.text;
  }
  const xs = (t: TextLayout): number[] => t.lines[0]!.runs.flatMap((run) => Array.from(run.x));

  test('kern="1200" at 18 pt leaves East Asian text unkerned', () => {
    // ECMA-376 21.1.2.3.9: `kern` is the smallest size, in hundredths of a point, that kerns.
    const text = 'アイウエオ';
    expect(xs(lay('sz="1800" kern="1200"', text))).toEqual(xs(lay('sz="1800" kern="0"', text)));
    // The Latin text of the same run kerns at 18 pt, and not below the threshold.
    const latin = 'AVATAR';
    expect(xs(lay('sz="1800" kern="1200"', latin))).not.toEqual(
      xs(lay('sz="1800" kern="0"', latin))
    );
    expect(xs(lay('sz="1100" kern="1200"', latin))).toEqual(xs(lay('sz="1100" kern="0"', latin)));
  });
});
