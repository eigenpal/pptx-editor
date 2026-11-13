// The underline of a run (`a:rPr/@u`, `ST_TextUnderlineType`, ECMA-376 Part 1, 20.1.10.82): the
// shapes that the SVG paint, the warped-text outline, and `pptx-to-pdf` all fill, so the three
// agree. The rules are measured on the reference PDF of the `underline-styles` fixture.

import type { GlyphRun } from './text-layout.ts';

const EMU_PER_PT = 12_700;

/**
 * A closed polygon: x in EMU from the text box left, y in EMU down from the text box top.
 * @internal
 */
export type UnderlineRing = readonly (readonly [number, number])[];

/**
 * Dash patterns in multiples of the underline thickness, starting with a dash: dash, gap, dash,
 * gap. A heavy style uses the pattern of its plain style, with the thickness of the plain line.
 */
const DASHES: Readonly<Record<string, readonly number[]>> = {
  dotted: [1, 1],
  dash: [3, 2],
  dashLong: [7, 4],
  dotDash: [3, 2, 1, 2],
  dotDotDash: [3, 2, 1, 2, 1, 2],
};

/** At most this many dashes or wave segments for one run; a longer one draws solid. */
const MAX_PIECES = 20_000;

/** The thin wave's band, and the least band of the heavy wave, in EMU (1 and 2 pt). */
const WAVE_BAND = EMU_PER_PT;
const HEAVY_WAVE_MIN_BAND = 2 * EMU_PER_PT;
/** The least half period of a zigzag, in EMU (1 pt). */
const MIN_WAVE_HALF = EMU_PER_PT;

/** The plain style of a heavy style, and whether the style is heavy. */
function plainStyle(style: string): { readonly plain: string; readonly heavy: boolean } {
  if (style === 'heavy') return { plain: 'sng', heavy: true };
  if (style.endsWith('Heavy')) return { plain: style.slice(0, -'Heavy'.length), heavy: true };
  return { plain: style, heavy: false };
}

/**
 * Whether `style` is an underline type that `underlineRings` draws (every type but `none`).
 * @internal
 */
export function isDrawnUnderline(style: string): boolean {
  const { plain } = plainStyle(style);
  return ['sng', 'words', 'dbl', 'wavy', 'wavyDbl'].includes(plain) || plain in DASHES;
}

function rect(x0: number, x1: number, y0: number, y1: number): UnderlineRing {
  return [
    [x0, y0],
    [x1, y0],
    [x1, y1],
    [x0, y1],
  ];
}

/**
 * A zigzag from `x0` to at least `x1` that hangs from `top`, `height` high, with a vertical band
 * of `band` and 45° slopes: each half period is the height less the band, at least 1 pt, and a
 * band that leaves less makes the zigzag taller. The zigzag starts at the top at `origin` plus a
 * whole number of half periods, so runs of one line keep one phase, and it ends at the first
 * top at or past `x1`. Returns the zigzag and its bottom.
 */
function wave(
  origin: number,
  x0: number,
  x1: number,
  top: number,
  height: number,
  band: number
): { readonly rings: UnderlineRing[]; readonly bottom: number } {
  const half = Math.max(height - band, MIN_WAVE_HALF);
  const bottom = top + half + band;
  if ((x1 - x0) / half > MAX_PIECES) return { rings: [rect(x0, x1, top, bottom)], bottom };
  const k0 = Math.floor((x0 - origin) / half + 1e-9);
  // It ends on a top: at the first top at or past `x1`.
  const end = Math.max(k0 + 1, Math.ceil((x1 - origin) / half - 1e-9));
  const k1 = end + (end & 1);
  const upper: [number, number][] = [];
  const lower: [number, number][] = [];
  for (let k = k0; k <= k1; k += 1) {
    const x = origin + k * half;
    const peak = (k & 1) === 0;
    upper.push([x, peak ? top : bottom - band]);
    lower.push([x, peak ? top + band : bottom]);
  }
  return { rings: [[...upper, ...lower.reverse()]], bottom };
}

/** Dashes of `pattern` × `unit` from `origin`, cut to [x0, x1], between `y0` and `y1`. */
function dashes(
  pattern: readonly number[],
  unit: number,
  origin: number,
  x0: number,
  x1: number,
  y0: number,
  y1: number
): UnderlineRing[] {
  const period = pattern.reduce((a, b) => a + b, 0) * unit;
  if (!(period > 0) || ((x1 - x0) / period) * pattern.length > MAX_PIECES)
    return [rect(x0, x1, y0, y1)];
  const out: UnderlineRing[] = [];
  let pos = origin + Math.floor((x0 - origin) / period) * period;
  for (let i = 0; pos < x1; i = (i + 1) % pattern.length) {
    const end = pos + pattern[i]! * unit;
    if (i % 2 === 0) {
      const a = Math.max(pos, x0);
      const b = Math.min(end, x1);
      if (b > a) out.push(rect(a, b, y0, y1));
    }
    pos = end;
  }
  return out;
}

/**
 * The shapes of an underline of `style` from `x0` to `x1` on the baseline `y`, all in EMU, y
 * down. `top` is the distance from the baseline down to the top of a single underline, and
 * `thickness` is its thickness in whole points. `origin` is where the dash pattern and the
 * zigzag start.
 *
 * Measured on the reference PDF of `underline-styles`, with `t` the thickness and `p` the top:
 * - `sng` and `words` are one bar from `p` to `p + t`. `words` underlines the spaces too.
 * - `dbl` is two bars of `t / 2` from `p - t` to `p + t`, with `t` between them.
 * - `heavy` is one bar from `p - t` to `p + t / 2`: 1.5 `t` thick, at the top of the double bars.
 * - The dashed styles cut the single bar, or the heavy bar for their heavy style, into dashes
 *   that start with a dash at `origin`: `dotted` 1 on and 1 off, `dash` 3 on and 2 off,
 *   `dashLong` 7 on and 4 off, `dotDash` 3, 2, 1, 2, and `dotDotDash` 3, 2, 1, 2, 1, 2, all in
 *   multiples of `t`, also for the heavy styles. The last dash ends at `x1`.
 * - `wavy` and `wavyHeavy` are a zigzag with 45° slopes, 2 `t` high, that hangs from the
 *   baseline. The band is 1 pt high for `wavy`, and `t / 2`, at least 2 pt, for `wavyHeavy`.
 *   `wavyDbl` is two zigzags of the thin band, each `t` high, the first from `p - t`.
 * - A half period is at least 1 pt, and a band that leaves less makes the zigzag taller: where
 *   `t` is 1 pt, the heavy zigzag is 3 pt high and each double zigzag 2 pt. A zigzag ends at the
 *   first top at or past `x1`.
 * @internal
 */
export function underlineRings(
  style: string,
  origin: number,
  x0: number,
  x1: number,
  y: number,
  top: number,
  thickness: number
): UnderlineRing[] {
  if (!(x1 > x0) || !(thickness > 0)) return [];
  const t = thickness;
  const p = y + top;
  const { plain, heavy } = plainStyle(style);
  const pattern = DASHES[plain];
  const [y0, y1] = heavy ? [p - t, p + t / 2] : [p, p + t];
  if (pattern) return dashes(pattern, t, origin, x0, x1, y0, y1);
  if (plain === 'dbl') {
    return [rect(x0, x1, p - t, p - t / 2), rect(x0, x1, p + t / 2, p + t)];
  }
  if (plain === 'wavyDbl') {
    const upper = wave(origin, x0, x1, p - t, t, WAVE_BAND);
    return [...upper.rings, ...wave(origin, x0, x1, upper.bottom, t, WAVE_BAND).rings];
  }
  if (plain === 'wavy') {
    const band = heavy ? Math.max(HEAVY_WAVE_MIN_BAND, t / 2) : WAVE_BAND;
    return wave(origin, x0, x1, y, 2 * t, band).rings;
  }
  return [rect(x0, x1, y0, y1)];
}

/**
 * The underline shapes of each underlined run of a line whose baseline is `baseline` (EMU from
 * the text box top), as `underlineRings` draws them. Measured on the reference PDF of
 * `underline-styles`:
 * - Every underlined run of a line takes the position and thickness of the largest underlined
 *   run on it, so runs of two sizes share one bar.
 * - The thickness is the face's, rounded to whole points and at least 1 pt: the bars, dashes,
 *   and zigzags are whole multiples of it.
 * - Spaces at the end of the line are not underlined; spaces between runs are.
 *
 * Adjacent runs with the same style keep one dash and wave phase, as one run does.
 * @internal
 */
export function lineUnderlines(
  runs: readonly GlyphRun[],
  baseline: number
): Map<GlyphRun, UnderlineRing[]> {
  const out = new Map<GlyphRun, UnderlineRing[]>();
  const underlined = (r: GlyphRun): boolean =>
    r.kind === 'text' && r.props.underline !== 'none' && r.text.length > 0 && !r.picture;
  if (!runs.some(underlined)) return out;
  // The last run with a character other than a space ends the line's underline.
  let last = -1;
  runs.forEach((r, i) => {
    if (r.kind === 'text' && /[^ ]/.test(r.text)) last = i;
  });
  const spans: { run: GlyphRun; x0: number; x1: number }[] = [];
  runs.forEach((r, i) => {
    if (!underlined(r) || i > last) return;
    let end = r.text.length;
    if (i === last) while (end > 0 && r.text[end - 1] === ' ') end -= 1;
    // A right-to-left run ends at its left: its trailing spaces are its leftmost cells.
    const x0 = r.rtl ? r.x[r.text.length - end]! : r.x[0]!;
    const x1 = r.rtl ? r.x[r.x.length - 1]! : (r.x[end] ?? r.x[r.x.length - 1]!);
    if (x1 > x0) spans.push({ run: r, x0, x1 });
  });
  if (spans.length === 0) return out;
  let largest = spans[0]!.run;
  for (const s of spans) if (s.run.sizePt > largest.sizePt) largest = s.run;
  const size = largest.sizePt * EMU_PER_PT;
  const m = largest.face.metrics;
  const top = -m.underlinePosition * size;
  const face = Math.max(m.underlineThickness * size, size / 20);
  const thickness = Math.max(1, Math.round(face / EMU_PER_PT)) * EMU_PER_PT;
  let origin = spans[0]!.x0;
  spans.forEach((s, i) => {
    const prev = spans[i - 1];
    const joined =
      prev && prev.run.props.underline === s.run.props.underline && Math.abs(prev.x1 - s.x0) < 1;
    if (!joined) origin = s.x0;
    const y = baseline - s.run.rise;
    out.set(s.run, underlineRings(s.run.props.underline, origin, s.x0, s.x1, y, top, thickness));
  });
  return out;
}
