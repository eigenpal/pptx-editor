// Preset text warps (`a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19). Text lays out flat in its
// text rectangle first; this module then maps each glyph outline from the flat frame into the
// warp, as paths. The SVG paint and `pptx-to-pdf` fill the same paths, so both draw one warp.
//
// The spec gives each warp's paths (`presetTextWarpDefinitions.xml`), but not how text maps onto
// them. The rules below are measured on reference PDFs (the `text-warp` fixture):
// - An even number of paths makes pairs: the first path of a pair is the top of its text and
//   the second the bottom. The ink bounds of the pair's lines stretch to fill the band between
//   the two paths, by arc length across and linearly between them. The insets, the anchor, and
//   the font size do not change the result.
// - With more than one pair, each pair's ink bounds grow about their center to the widest and
//   the tallest ink bounds of all pairs (`textDeflateInflate`: "Lower" has no descender, so its
//   bounds grow to those of "Upper"). Every pair keeps one text size.
// - Not matched: on `textStop` the reference text along the top path is about 2 % narrower
//   than along the bottom path, measured from the left. Here both paths take the same width.
// - An odd number of paths (`textArchUp`, `textCircle`, `textButton`) makes each path a line of
//   text at its own size, aligned along the path as the paragraph aligns. The baseline of the
//   line nearest the shape's middle is 0.164 em below the path; the other lines stack away from
//   the middle, each as far as its face is tall (ascent and descent), not at the flat line
//   pitch. Text longer than the path scales down to fit.
// - Lines split between pairs (or paths) in order; the first ones take the extra line.
// - The warp's own path count decides pairs. A path of zero length (an adjust value at the end
//   of its range) is a point, and its band narrows to that point.
//
// Glyphs bend with the warp: every outline is cut into short pieces before it maps. The caret
// and the selection keep the flat layout. Each character draws its font's nominal glyph, with
// no shaping: ligatures, contextual forms, and complex-script clusters draw as separate
// glyphs. A picture bullet does not bend: it turns and stretches with the warp at its corners.
// The cutting spends one output budget, so a font with huge glyphs draws the text flat early.

import { runGlyphs } from './bidi-line.ts';
import type { PathCommand } from '@eigenpal/ooxml-core';
import { evaluatePresetTextWarp } from '@eigenpal/ooxml-core';
import type { ResolvedTextBody, TextWarp } from '../../cascade/index.ts';
import type { TextMeasurer } from './measurer.ts';
import { strikeoutBars } from './strikeout-bars.ts';
import type { GlyphRun, TextLayout } from './text-layout.ts';

const EMU_PER_PT = 12_700;
/** The baseline of the line on a path sits this far below the path, in em (measured). */
const PATH_BASELINE_EM = 0.164;
/** Characters a warped text body maps; a longer body draws flat. */
const MAX_WARP_CHARACTERS = 2_000;
/** Output commands for one text body. */
const MAX_WARP_COMMANDS = 400_000;
/** Pieces each curve of a warp path is cut into, to measure its length and direction. */
const PATH_PIECES = 64;

/** The warped paths of one glyph run: its glyphs, its bars, and its highlight. */
export interface WarpedRun {
  /** The index of the line in `TextLayout.lines`, and of the run in that line. */
  readonly line: number;
  readonly run: number;
  /** The glyph outlines and the strikethrough, in the text frame (EMU, y down). Fill nonzero. */
  readonly glyphs: readonly PathCommand[];
  /** The underline bars, filled with the underline fill. */
  readonly underline: readonly PathCommand[];
  /** The highlight behind the run. */
  readonly highlight: readonly PathCommand[];
  /**
   * The picture of a picture bullet, as a parallelogram in the text frame (EMU): the warped
   * place of its top-left corner, and the vectors to its top-right and bottom-left corners.
   * The picture does not bend; it turns and stretches with the warp at its corners.
   */
  readonly picture?: {
    readonly origin: readonly [number, number];
    readonly across: readonly [number, number];
    readonly down: readonly [number, number];
  };
}

type Point = readonly [number, number];
type Polyline = Point[];
/** Where a point of the flat frame on a line goes in the warp. */
type WarpMap = (p: Point, line: number) => Point;

/** A path cut into pieces, with the length along it to each point. */
interface Track {
  readonly pts: readonly Point[];
  readonly len: readonly number[];
  readonly total: number;
}

function trackOf(commands: readonly PathCommand[], dx: number, dy: number): Track | null {
  const pts: Point[] = [];
  let cur: Point = [0, 0];
  const add = (p: Point): void => {
    pts.push([p[0] - dx, p[1] - dy]);
    cur = p;
  };
  for (const c of commands) {
    if (c.op === 'M' || c.op === 'L') add([c.x, c.y]);
    else if (c.op === 'Q' || c.op === 'C') {
      const [x0, y0] = cur;
      for (let i = 1; i <= PATH_PIECES; i += 1) {
        const t = i / PATH_PIECES;
        const u = 1 - t;
        add(
          c.op === 'Q'
            ? [
                u * u * x0 + 2 * u * t * c.x1 + t * t * c.x,
                u * u * y0 + 2 * u * t * c.y1 + t * t * c.y,
              ]
            : [
                u * u * u * x0 + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
                u * u * u * y0 + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y,
              ]
        );
      }
    }
  }
  return trackThrough(pts);
}

/**
 * A track through points, or null when there are none. A path of zero length is a point track
 * (`total` 0): it keeps its place in the pairs, and a band between it and a path narrows to it.
 */
function trackThrough(pts: readonly Point[]): Track | null {
  if (pts.length === 0) return null;
  const len = [0];
  for (let i = 1; i < pts.length; i += 1)
    len.push(len[i - 1]! + Math.hypot(pts[i]![0] - pts[i - 1]![0], pts[i]![1] - pts[i - 1]![1]));
  const total = len[len.length - 1]!;
  return Number.isFinite(total) ? { pts, len, total } : null;
}

/**
 * The track `d` below a track (to the right of its direction, y down), so a line of text keeps
 * its own spacing on its own baseline.
 */
function offsetTrack(track: Track, d: number): Track {
  if (d === 0 || track.total <= 0) return track;
  const pts = track.pts.map((p, i): Point => {
    const t = along(track, track.len[i]!).t;
    return [p[0] - t[1] * d, p[1] + t[0] * d];
  });
  const moved = trackThrough(pts);
  return moved && moved.total > 0 ? moved : track;
}

/** The point at length `s` along a track and the unit direction there; extends past the ends. */
function along(track: Track, s: number): { p: Point; t: Point } {
  const { pts, len } = track;
  let lo = 0;
  let hi = pts.length - 1;
  const at = Math.max(0, Math.min(s, track.total));
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (len[mid]! <= at) lo = mid;
    else hi = mid;
  }
  // Skip pieces of zero length, so a direction always exists.
  while (lo > 0 && len[hi]! - len[lo]! <= 0) lo -= 1;
  while (hi < pts.length - 1 && len[hi]! - len[lo]! <= 0) hi += 1;
  const a = pts[lo]!;
  const b = pts[hi]!;
  const d = len[hi]! - len[lo]! || 1;
  const t: Point = [(b[0] - a[0]) / d, (b[1] - a[1]) / d];
  const k = s - len[lo]!;
  return { p: [a[0] + t[0] * k, a[1] + t[1] * k], t };
}

/** One glyph or bar, as closed polylines in the flat text frame. */
interface Shape {
  readonly line: number;
  readonly run: number;
  readonly kind: 'glyphs' | 'underline' | 'highlight' | 'picture';
  readonly rings: readonly Polyline[];
}

/**
 * The output commands still allowed for one text body. Cutting spends it point by point, so a
 * font with huge glyphs stops the work early instead of after every glyph is cut.
 */
interface Budget {
  left: number;
}

/** Cut a glyph outline (em, y up) at `step` (em) into polylines in the flat frame. */
function glyphRings(
  outline: readonly PathCommand[],
  ox: number,
  baseline: number,
  size: number,
  step: number,
  budget: Budget
): Polyline[] {
  const rings: Polyline[] = [];
  let ring: Polyline = [];
  let cur: Point = [0, 0];
  let start: Point = [0, 0];
  const put = (p: Point): void => {
    budget.left -= 1;
    ring.push([ox + p[0] * size, baseline - p[1] * size]);
  };
  const line = (p: Point): void => {
    const n = Math.min(64, Math.max(1, Math.ceil(Math.hypot(p[0] - cur[0], p[1] - cur[1]) / step)));
    for (let i = 1; i <= n; i += 1)
      put([cur[0] + ((p[0] - cur[0]) * i) / n, cur[1] + ((p[1] - cur[1]) * i) / n]);
    cur = p;
  };
  // Each ring also writes a close command.
  const close = (): void => {
    if (ring.length > 1) {
      rings.push(ring);
      budget.left -= 1;
    }
    ring = [];
  };
  for (const c of outline) {
    if (budget.left < 0) return rings;
    if (c.op === 'M') {
      close();
      cur = start = [c.x, c.y];
      put(cur);
    } else if (c.op === 'L') line([c.x, c.y]);
    else if (c.op === 'Q' || c.op === 'C') {
      const [x0, y0] = cur;
      const span =
        Math.hypot(c.x1 - x0, c.y1 - y0) +
        (c.op === 'C'
          ? Math.hypot(c.x2 - c.x1, c.y2 - c.y1) + Math.hypot(c.x - c.x2, c.y - c.y2)
          : Math.hypot(c.x - c.x1, c.y - c.y1));
      const n = Math.min(64, Math.max(2, Math.ceil(span / step)));
      for (let i = 1; i <= n; i += 1) {
        const t = i / n;
        const u = 1 - t;
        put(
          c.op === 'Q'
            ? [
                u * u * x0 + 2 * u * t * c.x1 + t * t * c.x,
                u * u * y0 + 2 * u * t * c.y1 + t * t * c.y,
              ]
            : [
                u * u * u * x0 + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
                u * u * u * y0 + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y,
              ]
        );
      }
      cur = [c.x, c.y];
    } else {
      line(start);
      close();
    }
  }
  close();
  return rings;
}

/** A rectangle in the flat frame, cut at `step` (EMU). */
function rectRing(
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  step: number,
  budget: Budget
): Polyline {
  const ring: Polyline = [];
  const edge = (a: Point, b: Point): void => {
    const n = Math.min(256, Math.max(1, Math.ceil(Math.hypot(b[0] - a[0], b[1] - a[1]) / step)));
    for (let i = 0; i < n; i += 1)
      ring.push([a[0] + ((b[0] - a[0]) * i) / n, a[1] + ((b[1] - a[1]) * i) / n]);
  };
  edge([x0, y0], [x1, y0]);
  edge([x1, y0], [x1, y1]);
  edge([x1, y1], [x0, y1]);
  edge([x0, y1], [x0, y0]);
  budget.left -= ring.length + 1;
  return ring;
}

/** The underline, strikethrough, and highlight bars of a run, as the flat paint draws them. */
function runBars(
  run: GlyphRun,
  top: number,
  height: number,
  baseline: number,
  step: number,
  budget: Budget
): [Polyline[], Polyline[], Polyline[]] {
  const size = run.sizePt * EMU_PER_PT;
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  // A right-to-left run ends at its left: its trailing spaces are its leftmost cells.
  const x0 = run.rtl ? run.x[run.text.length - end]! : run.x[0]!;
  const x1 = run.rtl ? run.x[run.x.length - 1]! : (run.x[end] ?? run.x[run.x.length - 1]!);
  const strike: Polyline[] = [];
  const under: Polyline[] = [];
  const highlight: Polyline[] = [];
  const rect = (a: number, b: number, c: number, d: number): Polyline =>
    rectRing(a, b, c, d, step, budget);
  if (run.props.highlight && run.kind === 'text')
    highlight.push(rect(x0, top, run.x[run.x.length - 1]!, top + height));
  if (run.kind === 'text' && x1 > x0) {
    const m = run.face.metrics;
    const thick = Math.max(m.underlineThickness * size, size / 20);
    const y = baseline - run.rise;
    if (run.props.underline !== 'none') {
      const ty = y - m.underlinePosition * size;
      under.push(rect(x0, ty, x1, ty + thick));
      if (run.props.underline === 'dbl') under.push(rect(x0, ty + thick * 2, x1, ty + thick * 3));
    }
    if (run.props.strike !== 'noStrike') {
      const bars = strikeoutBars(m, size, run.props.strike === 'dblStrike');
      for (const top of bars.tops) strike.push(rect(x0, y + top, x1, y + top + bars.thickness));
    }
  }
  return [strike, under, highlight];
}

/** Why a warp draws flat, or the warped runs. */
export type WarpResult = { readonly warped: readonly WarpedRun[] } | { readonly flat: string };

/**
 * Map the glyphs of a flat text layout into a preset text warp. `w` and `h` are the shape's
 * size: the warp spans the whole shape, not the text rectangle. `body` gives each paragraph's
 * alignment along a path.
 */
export function warpTextLayout(
  layout: TextLayout,
  body: ResolvedTextBody,
  warp: TextWarp,
  w: number,
  h: number,
  measurer: TextMeasurer
): WarpResult {
  if (layout.rotation !== 0 || layout.upright !== 'none')
    return { flat: 'rotated or vertical text' };
  if (!measurer.glyphOutline) return { flat: 'no glyph outlines' };
  const paths = evaluatePresetTextWarp(warp.preset, w, h, warp.adjust) ?? [];
  // One track for each path, in order: a path of zero length stays as a point, so the pairs
  // still form from the paths that the warp defines.
  const tracks: Track[] = [];
  for (const p of paths) {
    const track = trackOf(p.commands, layout.box.x, layout.box.y);
    if (!track) return { flat: 'no warp paths' };
    tracks.push(track);
  }
  if (!tracks.some((t) => t.total > 0)) return { flat: 'no warp paths' };
  let characters = 0;
  for (const line of layout.lines) for (const run of line.runs) characters += run.text.length;
  if (characters > MAX_WARP_CHARACTERS) return { flat: 'too much text' };

  // Every glyph and bar as polylines in the flat frame, within the output budget.
  const budget: Budget = { left: MAX_WARP_COMMANDS };
  const shapes: Shape[] = [];
  const lines = layout.lines;
  for (let li = 0; li < lines.length; li += 1) {
    const line = lines[li]!;
    for (let ri = 0; ri < line.runs.length; ri += 1) {
      const run = line.runs[ri]!;
      const pic = run.picture;
      if (pic) {
        const x0 = run.x[0]!;
        const x1 = run.x[run.x.length - 1]!;
        if (x1 > x0 && pic.height > 0) {
          const ring = rectRing(
            x0,
            line.baseline - pic.height,
            x1,
            line.baseline,
            Infinity,
            budget
          );
          shapes.push({ line: li, run: ri, kind: 'picture', rings: [ring] });
        }
        continue;
      }
      const size = run.sizePt * EMU_PER_PT;
      if (size <= 0) continue;
      const step = 0.03;
      const glyphs: Polyline[] = [];
      // A right-to-left run places each character in its own cell (`runGlyphs`).
      for (const glyph of runGlyphs(run)) {
        const cp = glyph.char.codePointAt(0) ?? 0;
        const x = glyph.x;
        const outline = measurer.glyphOutline(run.face, cp);
        if (!outline) continue;
        for (const ring of glyphRings(outline, x, line.baseline - run.rise, size, step, budget))
          glyphs.push(ring);
        if (budget.left < 0) return { flat: 'too much text' };
      }
      const bars = runBars(run, line.top, line.height, line.baseline, size * step, budget);
      if (budget.left < 0) return { flat: 'too much text' };
      const [strike, under, highlight] = bars;
      for (const ring of strike) glyphs.push(ring);
      shapes.push({ line: li, run: ri, kind: 'glyphs', rings: glyphs });
      if (under.length) shapes.push({ line: li, run: ri, kind: 'underline', rings: under });
      if (highlight.length) shapes.push({ line: li, run: ri, kind: 'highlight', rings: highlight });
    }
  }
  if (budget.left < 0) return { flat: 'too much text' };

  // Lines split between pairs, or between paths, in order. The warp's own path count decides.
  const pairs = tracks.length % 2 === 0;
  const groups = pairs ? tracks.length / 2 : tracks.length;
  const groupOf = (li: number): number => {
    for (let g = groups - 1; g >= 0; g -= 1)
      if (li >= Math.ceil((g * lines.length) / groups)) return g;
    return 0;
  };
  const owned = Array.from({ length: groups }, (_, g) =>
    shapes.filter((s) => groupOf(s.line) === g && s.kind !== 'highlight')
  );
  // Each pair stretches the ink bounds of its lines, grown about their center to the widest and
  // the tallest ink bounds of all pairs, so the pairs keep one text size.
  const inks = pairs ? owned.map(inkBounds) : [];
  let [widest, tallest] = [0, 0];
  for (const ink of inks)
    if (ink)
      [widest, tallest] = [Math.max(widest, ink[2] - ink[0]), Math.max(tallest, ink[3] - ink[1])];
  const maps = new Map<number, WarpMap>();
  for (let g = 0; g < groups; g += 1) {
    const own = owned[g]!;
    const ink = inks[g];
    const map = pairs
      ? ink && pairMap(grow(ink, widest, tallest), tracks[2 * g]!, tracks[2 * g + 1]!)
      : pathMap(
          layout,
          body,
          lines.map((_, li) => li).filter((li) => groupOf(li) === g),
          tracks[g]!,
          own,
          w / 2 - layout.box.x,
          h / 2 - layout.box.y
        );
    if (map) maps.set(g, map);
  }

  type Entry = {
    glyphs: PathCommand[];
    underline: PathCommand[];
    highlight: PathCommand[];
    picture?: WarpedRun['picture'];
  };
  const out = new Map<string, Entry>();
  for (const s of shapes) {
    const map = maps.get(groupOf(s.line));
    if (!map) continue;
    const key = `${s.line}|${s.run}`;
    let entry = out.get(key);
    if (!entry) out.set(key, (entry = { glyphs: [], underline: [], highlight: [] }));
    if (s.kind === 'picture') {
      // The ring starts at the top-left corner and runs to the top-right, then down.
      const ring = s.rings[0]!;
      const [x0, y0] = ring[0]!;
      const x1 = ring[1]![0];
      const y1 = ring[2]![1];
      const o = map([x0, y0], s.line);
      const a = map([x1, y0], s.line);
      const d = map([x0, y1], s.line);
      entry.picture = {
        origin: o,
        across: [a[0] - o[0], a[1] - o[1]],
        down: [d[0] - o[0], d[1] - o[1]],
      };
      continue;
    }
    const target = entry[s.kind];
    for (const ring of s.rings) {
      ring.forEach((pt, i) => {
        const [x, y] = map(pt, s.line);
        target.push({ op: i === 0 ? 'M' : 'L', x, y });
      });
      target.push({ op: 'Z' });
    }
  }
  const warped: WarpedRun[] = [];
  for (const [key, e] of out) {
    const [line, run] = key.split('|').map(Number) as [number, number];
    warped.push({ line, run, ...e });
  }
  return { warped };
}

/** The ink bounds of shapes in the flat frame, or null when they are empty. */
function inkBounds(shapes: readonly Shape[]): [number, number, number, number] | null {
  let [x0, y0, x1, y1] = [Infinity, Infinity, -Infinity, -Infinity];
  for (const s of shapes)
    for (const ring of s.rings)
      for (const [x, y] of ring) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
  return x1 > x0 && y1 > y0 ? [x0, y0, x1, y1] : null;
}

/** Bounds grown about their center to at least `w` wide and `h` tall. */
function grow(
  [x0, y0, x1, y1]: readonly [number, number, number, number],
  w: number,
  h: number
): [number, number, number, number] {
  const dx = Math.max(0, w - (x1 - x0)) / 2;
  const dy = Math.max(0, h - (y1 - y0)) / 2;
  return [x0 - dx, y0 - dy, x1 + dx, y1 + dy];
}

/** The map of a pair of paths: the bounds stretch between the top and the bottom path. */
function pairMap(
  [x0, y0, x1, y1]: readonly [number, number, number, number],
  top: Track,
  bottom: Track
): WarpMap {
  return ([x, y]) => {
    const u = (x - x0) / (x1 - x0);
    const v = (y - y0) / (y1 - y0);
    const a = along(top, u * top.total).p;
    const b = along(bottom, u * bottom.total).p;
    return [a[0] + (b[0] - a[0]) * v, a[1] + (b[1] - a[1]) * v];
  };
}

/** The map of one path: the lines follow it at their own size. */
function pathMap(
  layout: TextLayout,
  body: ResolvedTextBody,
  own: readonly number[],
  track: Track,
  shapes: readonly Shape[],
  cx: number,
  cy: number
): WarpMap | null {
  if (own.length === 0) return null;
  // Each line's extent along the path: its advances without trailing spaces.
  const extent = new Map<number, [number, number]>();
  let widest = 0;
  for (const li of own) {
    let [a, b] = [Infinity, -Infinity];
    for (const run of layout.lines[li]!.runs) {
      // A picture bullet has no text; its edges are its extent.
      let end = run.picture ? run.x.length - 1 : run.text.length;
      while (!run.picture && end > 0 && run.text[end - 1] === ' ') end -= 1;
      if (end === 0) continue;
      a = Math.min(a, run.x[0]!);
      b = Math.max(b, run.x[end]!);
    }
    if (b > a) {
      extent.set(li, [a, b]);
      widest = Math.max(widest, b - a);
    }
  }
  if (widest <= 0 || shapes.length === 0) return null;
  const scale = Math.min(1, track.total / widest);
  // The line nearest the middle sits on the path; the others stack away from the middle.
  const mid = along(track, track.total / 2);
  const down: Point = [-mid.t[1], mid.t[0]];
  const inward = down[0] * (cx - mid.p[0]) + down[1] * (cy - mid.p[1]) > 0;
  // Each line's baseline below the path: lines stack at the height of their tallest face.
  const tallest = (li: number): number => {
    let most = 0;
    for (const r of layout.lines[li]!.runs)
      most = Math.max(most, r.sizePt * (r.face.metrics.ascent + r.face.metrics.descent));
    return most * EMU_PER_PT;
  };
  const sizeOf = (li: number): number => {
    let most = 0;
    for (const r of layout.lines[li]!.runs) most = Math.max(most, r.sizePt);
    return most * EMU_PER_PT;
  };
  const below = new Map<number, number>();
  const at = inward ? own.length - 1 : 0;
  below.set(own[at]!, 0);
  for (let i = at + 1; i < own.length; i += 1)
    below.set(own[i]!, below.get(own[i - 1]!)! + tallest(own[i]!));
  for (let i = at - 1; i >= 0; i -= 1)
    below.set(own[i]!, below.get(own[i + 1]!)! - tallest(own[i + 1]!));
  // The line on the path keeps its spacing on the path; each other line keeps it on the path
  // moved by its stacking distance.
  const anchorSize = sizeOf(own[at]!);
  const baselines = new Map(
    own.map((li) => [li, offsetTrack(track, (below.get(li) ?? 0) * scale)])
  );
  return ([x, y], li) => {
    const base = baselines.get(li) ?? track;
    const [a, b] = extent.get(li) ?? [0, 0];
    const align = body.paragraphs[layout.lines[li]!.paragraph]?.align ?? 'l';
    const s =
      align === 'ctr'
        ? base.total / 2 + (x - (a + b) / 2) * scale
        : align === 'r'
          ? base.total - (b - x) * scale
          : (x - a) * scale;
    const d = (y - layout.lines[li]!.baseline + PATH_BASELINE_EM * anchorSize) * scale;
    const { p, t } = along(base, s);
    return [p[0] - t[1] * d, p[1] + t[0] * d];
  };
}
