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
// - An odd number of paths (`textArchUp`, `textCircle`, `textButton`) makes each path a line of
//   text at its own size, aligned along the path as the paragraph aligns. The baseline of the
//   line nearest the shape's middle is 0.164 em below the path; the other lines stack away from
//   the middle, each as far as its face is tall (ascent and descent), not at the flat line
//   pitch. Text longer than the path scales down to fit.
// - Lines split between pairs (or paths) in order; the first ones take the extra line.
//
// Glyphs bend with the warp: every outline is cut into short pieces before it maps. The caret
// and the selection keep the flat layout.

import type { PathCommand } from '@eigenpal/ooxml-core';
import { evaluatePresetTextWarp } from '@eigenpal/ooxml-core';
import type { ResolvedTextBody, TextWarp } from '../../cascade/index.ts';
import type { TextMeasurer } from './measurer.ts';
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

function trackThrough(pts: readonly Point[]): Track | null {
  const len = [0];
  for (let i = 1; i < pts.length; i += 1)
    len.push(len[i - 1]! + Math.hypot(pts[i]![0] - pts[i - 1]![0], pts[i]![1] - pts[i - 1]![1]));
  const total = len[len.length - 1] ?? 0;
  return total > 0 ? { pts, len, total } : null;
}

/**
 * The track `d` below a track (to the right of its direction, y down), so a line of text keeps
 * its own spacing on its own baseline.
 */
function offsetTrack(track: Track, d: number): Track {
  if (d === 0) return track;
  const pts = track.pts.map((p, i): Point => {
    const t = along(track, track.len[i]!).t;
    return [p[0] - t[1] * d, p[1] + t[0] * d];
  });
  return trackThrough(pts) ?? track;
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
  readonly kind: 'glyphs' | 'underline' | 'highlight';
  readonly rings: readonly Polyline[];
}

/** Cut a glyph outline (em, y up) at `step` (em) into polylines in the flat frame. */
function glyphRings(
  outline: readonly PathCommand[],
  ox: number,
  baseline: number,
  size: number,
  step: number
): Polyline[] {
  const rings: Polyline[] = [];
  let ring: Polyline = [];
  let cur: Point = [0, 0];
  let start: Point = [0, 0];
  const put = (p: Point): void => {
    ring.push([ox + p[0] * size, baseline - p[1] * size]);
  };
  const line = (p: Point): void => {
    const n = Math.min(64, Math.max(1, Math.ceil(Math.hypot(p[0] - cur[0], p[1] - cur[1]) / step)));
    for (let i = 1; i <= n; i += 1)
      put([cur[0] + ((p[0] - cur[0]) * i) / n, cur[1] + ((p[1] - cur[1]) * i) / n]);
    cur = p;
  };
  for (const c of outline) {
    if (c.op === 'M') {
      if (ring.length > 1) rings.push(ring);
      ring = [];
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
      if (ring.length > 1) rings.push(ring);
      ring = [];
    }
  }
  if (ring.length > 1) rings.push(ring);
  return rings;
}

/** A rectangle in the flat frame, cut at `step` (EMU). */
function rectRing(x0: number, y0: number, x1: number, y1: number, step: number): Polyline {
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
  return ring;
}

/** The underline, strikethrough, and highlight bars of a run, as the flat paint draws them. */
function runBars(
  run: GlyphRun,
  top: number,
  height: number,
  baseline: number,
  step: number
): Shape['rings'][] {
  const size = run.sizePt * EMU_PER_PT;
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  const x0 = run.x[0]!;
  const x1 = run.x[end] ?? run.x[run.x.length - 1]!;
  const strike: Polyline[] = [];
  const under: Polyline[] = [];
  const highlight: Polyline[] = [];
  if (run.props.highlight)
    highlight.push(rectRing(x0, top, run.x[run.x.length - 1]!, top + height, step));
  if (run.kind === 'text' && x1 > x0) {
    const m = run.face.metrics;
    const thick = Math.max(m.underlineThickness * size, size / 20);
    const y = baseline - run.rise;
    if (run.props.underline !== 'none') {
      const ty = y - m.underlinePosition * size;
      under.push(rectRing(x0, ty, x1, ty + thick, step));
      if (run.props.underline === 'dbl')
        under.push(rectRing(x0, ty + thick * 2, x1, ty + thick * 3, step));
    }
    if (run.props.strike !== 'noStrike') {
      const sy = y - size * 0.3;
      strike.push(rectRing(x0, sy - thick / 2, x1, sy + thick / 2, step));
      if (run.props.strike === 'dblStrike')
        strike.push(rectRing(x0, sy - thick * 2, x1, sy - thick, step));
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
  const tracks = paths
    .map((p) => trackOf(p.commands, layout.box.x, layout.box.y))
    .filter((t): t is Track => t !== null);
  if (tracks.length === 0) return { flat: 'no warp paths' };
  let characters = 0;
  for (const line of layout.lines) for (const run of line.runs) characters += run.text.length;
  if (characters > MAX_WARP_CHARACTERS) return { flat: 'too much text' };

  // Every glyph and bar as polylines in the flat frame.
  const shapes: Shape[] = [];
  const lines = layout.lines;
  lines.forEach((line, li) => {
    line.runs.forEach((run, ri) => {
      const size = run.sizePt * EMU_PER_PT;
      if (size <= 0) return;
      const step = 0.03;
      const glyphs: Polyline[] = [];
      for (let i = 0; i < run.text.length; i += 1) {
        const cp = run.text.codePointAt(i) ?? 0;
        const x = run.x[i]!;
        if (cp > 0xffff) i += 1;
        const outline = measurer.glyphOutline!(run.face, cp);
        if (!outline) continue;
        for (const ring of glyphRings(outline, x, line.baseline - run.rise, size, step))
          glyphs.push(ring);
      }
      const [strike, under, highlight] = runBars(
        run,
        line.top,
        line.height,
        line.baseline,
        size * step
      );
      shapes.push({ line: li, run: ri, kind: 'glyphs', rings: [...glyphs, ...strike] });
      if (under.length) shapes.push({ line: li, run: ri, kind: 'underline', rings: under });
      if (highlight.length) shapes.push({ line: li, run: ri, kind: 'highlight', rings: highlight });
    });
  });

  // Lines split between pairs, or between paths, in order.
  const pairs = tracks.length % 2 === 0;
  const groups = pairs ? tracks.length / 2 : tracks.length;
  const groupOf = (li: number): number => {
    for (let g = groups - 1; g >= 0; g -= 1)
      if (li >= Math.ceil((g * lines.length) / groups)) return g;
    return 0;
  };
  const maps = new Map<number, WarpMap>();
  for (let g = 0; g < groups; g += 1) {
    const own = shapes.filter((s) => groupOf(s.line) === g && s.kind !== 'highlight');
    const map = pairs
      ? pairMap(own, tracks[2 * g]!, tracks[2 * g + 1]!)
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

  const out = new Map<
    string,
    { glyphs: PathCommand[]; underline: PathCommand[]; highlight: PathCommand[] }
  >();
  let budget = MAX_WARP_COMMANDS;
  for (const s of shapes) {
    const map = maps.get(groupOf(s.line));
    if (!map) continue;
    const key = `${s.line}|${s.run}`;
    let entry = out.get(key);
    if (!entry) out.set(key, (entry = { glyphs: [], underline: [], highlight: [] }));
    const target = entry[s.kind];
    for (const ring of s.rings) {
      budget -= ring.length + 1;
      if (budget < 0) return { flat: 'too much text' };
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

/** The map of a pair of paths: the ink bounds stretch between the top and the bottom path. */
function pairMap(shapes: readonly Shape[], top: Track, bottom: Track): WarpMap | null {
  const ink = inkBounds(shapes);
  if (!ink) return null;
  const [x0, y0, x1, y1] = ink;
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
      let end = run.text.length;
      while (end > 0 && run.text[end - 1] === ' ') end -= 1;
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
