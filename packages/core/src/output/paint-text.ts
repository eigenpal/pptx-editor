// Text paint. Every glyph run becomes one `<text>` whose `x` list carries the layout's
// position for each character, so the browser draws exactly where HarfBuzz measured and never
// re-breaks or re-spaces a line. Underline, strikethrough, and highlight are drawn as shapes
// from the face metrics, not by CSS text decoration, so they match the measured geometry.

import { toCss, type ResolvedFill } from '../cascade/index.ts';
import {
  isIdeographic,
  lineUnderlines,
  strikeoutBars,
  type GlyphRun,
  type PathCommand,
  type TextLayout,
  type UnderlineRing,
  type WarpedRun,
} from '../layout/index.ts';
import { fillPaint, STORED_PIXELS_STYLE, type PaintContext } from './paint-fill.ts';
import { num, svg } from './svg.ts';

const EMU_PER_PT = 12_700;
/** The laid-out bounds of a text body, in points. */
interface TextBounds {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

function textBounds(layout: TextLayout): TextBounds {
  let x = Infinity;
  let r = -Infinity;
  let y = Infinity;
  let b = -Infinity;
  for (const line of layout.lines) {
    y = Math.min(y, line.top);
    b = Math.max(b, line.top + line.height);
    for (const run of line.runs)
      if (run.x.length) {
        x = Math.min(x, run.x[0]);
        r = Math.max(r, run.x[run.x.length - 1]);
      }
  }
  if (!Number.isFinite(x) || !Number.isFinite(y)) return { x: 0, y: 0, w: 1, h: 1 };
  const pt = (v: number): number => v / EMU_PER_PT;
  return { x: pt(x), y: pt(y), w: Math.max(pt(r - x), 1), h: Math.max(pt(b - y), 1) };
}

/**
 * A run's fill. Gradients and patterns span the text body's laid-out bounds, in points, the
 * space text paints in (design D3 of `fill-and-line-fidelity`).
 */
function textFill(
  fill: ResolvedFill,
  bounds: TextBounds,
  ctx: PaintContext
): { paint: string; opacity: number } {
  const paint = fillPaint(fill, bounds.w, bounds.h, ctx);
  if (fill.type === 'linear' || fill.type === 'path' || fill.type === 'pattern') {
    const server = ctx.defs.find(paint.paint);
    const attr = server?.tagName === 'pattern' ? 'patternTransform' : 'gradientTransform';
    // Pattern pixels are sized in EMU; text paints in points.
    const scale = fill.type === 'pattern' ? ` scale(${1 / EMU_PER_PT})` : '';
    server?.setAttribute(attr, `translate(${num(bounds.x)} ${num(bounds.y)})${scale}`);
  }
  return paint;
}

type TextPaint = { readonly paint: string; readonly opacity: number };

function decorationRects(
  doc: Document,
  run: GlyphRun,
  baselineEmu: number,
  text: TextPaint,
  underline: TextPaint,
  underlineRings: readonly UnderlineRing[] | undefined
): SVGElement[] {
  const out: SVGElement[] = [];
  // The underline's shapes come from its line (`lineUnderlines`). It takes its own fill
  // (`a:uFill`) when the run has one, and its outline (`a:uLn`) strokes it in a solid color.
  if (underlineRings?.length) {
    const line = run.props.underlineLine;
    const color = line?.fill.type === 'solid' ? line.fill.color : undefined;
    out.push(
      svg(doc, 'path', {
        d: ringPathData(underlineRings),
        fill: underline.paint,
        'fill-opacity': underline.opacity < 1 ? num(underline.opacity) : undefined,
        stroke: color ? toCss(color) : undefined,
        'stroke-width': color && line ? num(line.width / EMU_PER_PT) : undefined,
        'stroke-linejoin': color && line ? line.join : undefined,
      })
    );
  }
  if (run.props.strike === 'noStrike') return out;
  const size = run.sizePt;
  const baseline = baselineEmu / EMU_PER_PT;
  const xs = run.x.map((v) => v / EMU_PER_PT);
  // Trailing spaces are not struck through; the layout already excludes them from
  // alignment, and runs keep them, so trim here.
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  // A right-to-left run ends at its left: its trailing spaces are its leftmost cells.
  const x0 = run.rtl ? xs[run.text.length - end] : xs[0];
  const x1 = run.rtl ? xs[xs.length - 1] : (xs[end] ?? xs[xs.length - 1]);
  if (x1 <= x0) return out;
  const y = baseline - run.rise / EMU_PER_PT;
  // A strikethrough takes the run's fill and its alpha, as the glyphs do.
  const strike = strikeoutBars(run.face.metrics, size, run.props.strike === 'dblStrike');
  for (const top of strike.tops)
    out.push(
      svg(doc, 'rect', {
        x: num(x0),
        y: num(y + top),
        width: num(x1 - x0),
        height: num(strike.thickness),
        fill: text.paint,
        'fill-opacity': text.opacity < 1 ? num(text.opacity) : undefined,
      })
    );
  return out;
}

/** The path data of closed rings in EMU, in points. */
function ringPathData(rings: readonly UnderlineRing[]): string {
  const pt = (emu: number): string => num(emu / EMU_PER_PT);
  return rings
    .map((ring) => ring.map(([x, y], i) => `${i ? 'L' : 'M'}${pt(x)} ${pt(y)}`).join('') + 'Z')
    .join('');
}

/**
 * The style of a right-to-left word: its letters in source order, anchored at its right edge,
 * with the features the measurer shapes with (`harfbuzz-measurer.ts`): no standard,
 * contextual, discretionary, or historical ligatures, contextual alternates kept, and kerning
 * only when the run asks for it.
 */
function rtlStyle(kerning: boolean): string {
  return (
    'white-space:pre;direction:rtl;unicode-bidi:bidi-override;' +
    `font-kerning:${kerning ? 'normal' : 'none'};` +
    'font-variant-ligatures:no-common-ligatures no-discretionary-ligatures no-historical-ligatures;' +
    'font-feature-settings:"liga" 0,"clig" 0,"dlig" 0,"hlig" 0'
  );
}

/**
 * The words of a right-to-left run, each painted as one `<text>` that starts at the word's
 * right edge and reads right to left, so the browser shapes and joins its letters. Spaces
 * between the words take the room that layout gave them, justification included, and each
 * word fills the width layout measured for it (`textLength`), so character spacing, kerning
 * that the browser applies differently, and a fallback face for a missing glyph never move
 * the next word. Every
 * character of the run has an odd bidi level, so the override keeps the browser from
 * reordering the word again.
 */
function rtlWords(run: GlyphRun): { text: string; right: number; width: number }[] {
  const out: { text: string; right: number; width: number }[] = [];
  const n = run.text.length;
  // A tab draws nothing (`uprightPieces`): it splits words as a space does.
  const re = /[^ \t\u3000]+/gu;
  // Unit `i` draws in cell `n - 1 - i`, so a word that starts at unit `i` ends at `x[n - i]`.
  for (let m = re.exec(run.text); m; m = re.exec(run.text)) {
    const right = run.x[n - m.index];
    out.push({ text: m[0], right, width: right - run.x[n - m.index - m[0].length] });
  }
  return out;
}

/**
 * Split a run into stretches painted along the line and single characters painted upright.
 * Offsets are UTF-16; a surrogate pair stays together. Tabs have no glyph and are omitted.
 * Their advances remain in the layout, which also supplies decoration and caret positions.
 */
export function uprightPieces(
  text: string,
  mode: TextLayout['upright']
): { start: number; end: number; upright: boolean }[] {
  if (mode === 'none' && !text.includes('\t'))
    return [{ start: 0, end: text.length, upright: false }];
  const out: { start: number; end: number; upright: boolean }[] = [];
  let i = 0;
  while (i < text.length) {
    const cp = text.codePointAt(i) ?? 0;
    const len = cp > 0xffff ? 2 : 1;
    if (cp === 9) {
      i += len;
      continue;
    }
    const upright = mode === 'all' ? text[i] !== ' ' : mode !== 'none' && isIdeographic(cp);
    const last = out[out.length - 1];
    if (!upright && last && !last.upright && last.end === i) last.end = i + len;
    else out.push({ start: i, end: i + len, upright });
    i += len;
  }
  return out;
}

/**
 * The picture of a picture bullet, standing on the baseline, in points; null when the picture
 * has no image to show (the bullet then shows nothing).
 */
function bulletPicture(
  doc: Document,
  run: GlyphRun,
  baseline: number,
  ctx: PaintContext
): SVGElement | null {
  const pic = run.picture;
  if (!pic) return null;
  const url =
    ctx.images(pic.part, pic.svgRId ?? pic.rId) ??
    (pic.svgRId ? ctx.images(pic.part, pic.rId) : undefined);
  const x = run.x[0];
  const width = run.x[run.x.length - 1] - x;
  if (!url || !(width > 0) || !(pic.height > 0)) return null;
  return svg(doc, 'image', {
    href: url,
    x: num(x / EMU_PER_PT),
    y: num((baseline - pic.height) / EMU_PER_PT),
    width: num(width / EMU_PER_PT),
    height: num(pic.height / EMU_PER_PT),
    preserveAspectRatio: 'none',
    style: STORED_PIXELS_STYLE,
    'data-run': 'bullet',
  });
}

/** SVG path data for commands in EMU, written in points. */
function pointPathData(commands: readonly PathCommand[]): string {
  const p = (v: number): string => num(v / EMU_PER_PT);
  let d = '';
  for (const c of commands) {
    if (c.op === 'Z') d += 'Z';
    else if (c.op === 'M' || c.op === 'L') d += `${c.op}${p(c.x)} ${p(c.y)}`;
    else if (c.op === 'Q') d += `Q${p(c.x1)} ${p(c.y1)} ${p(c.x)} ${p(c.y)}`;
    else if (c.op === 'C') d += `C${p(c.x1)} ${p(c.y1)} ${p(c.x2)} ${p(c.y2)} ${p(c.x)} ${p(c.y)}`;
  }
  return d;
}

/**
 * The marks of a line's equations (fraction rules, radical signs, stretched delimiters, braces):
 * filled paths in the math run's fill, after the line's glyphs.
 */
function paintMarks(
  doc: Document,
  line: TextLayout['lines'][number],
  group: SVGElement,
  bounds: TextBounds,
  ctx: PaintContext
): void {
  for (const mark of line.marks ?? []) {
    const fill = textFill(mark.props.fill, bounds, ctx);
    if (fill.paint === 'none') continue;
    const commands = mark.commands.map((c): PathCommand => {
      if (c.op === 'Z') return c;
      if (c.op === 'Q') return { ...c, y1: line.baseline - c.y1, y: line.baseline - c.y };
      if (c.op === 'C')
        return { ...c, y1: line.baseline - c.y1, y2: line.baseline - c.y2, y: line.baseline - c.y };
      return { ...c, y: line.baseline - c.y };
    });
    group.appendChild(
      svg(doc, 'path', {
        d: pointPathData(commands),
        fill: fill.paint,
        'fill-opacity': fill.opacity < 1 ? num(fill.opacity) : undefined,
        'data-run': String(mark.runIndex),
      })
    );
  }
}

/**
 * Paint warped text (`a:prstTxWarp`): each run's glyphs as one filled path that layout mapped
 * into the warp, with its highlight behind and its underline over it, in the run's fills.
 */
function paintWarped(
  doc: Document,
  layout: TextLayout,
  warped: readonly WarpedRun[],
  group: SVGElement,
  bounds: TextBounds,
  ctx: PaintContext
): void {
  for (const w of warped) {
    const run = layout.lines[w.line]?.runs[w.run];
    if (!run) continue;
    const pic = run.picture;
    if (pic) {
      // The bullet's picture on the parallelogram the warp gives its corners.
      const url =
        ctx.images(pic.part, pic.svgRId ?? pic.rId) ??
        (pic.svgRId ? ctx.images(pic.part, pic.rId) : undefined);
      if (url && w.picture) {
        const { origin, across, down } = w.picture;
        const m = [across[0], across[1], down[0], down[1], origin[0], origin[1]];
        group.appendChild(
          svg(doc, 'image', {
            href: url,
            x: '0',
            y: '0',
            width: '1',
            height: '1',
            preserveAspectRatio: 'none',
            transform: `matrix(${m.map((v) => num(v / EMU_PER_PT)).join(' ')})`,
            style: STORED_PIXELS_STYLE,
            'data-run': 'bullet',
          })
        );
      }
      continue;
    }
    const fill = textFill(run.props.fill, bounds, ctx);
    if (w.highlight.length && run.props.highlight)
      group.appendChild(
        svg(doc, 'path', {
          d: pointPathData(w.highlight),
          fill: toCss({ ...run.props.highlight, a: 1 }),
        })
      );
    if (fill.paint === 'none' && !run.props.outline) continue;
    const outline = run.props.outline;
    group.appendChild(
      svg(doc, 'path', {
        d: pointPathData(w.glyphs),
        fill: fill.paint,
        'fill-opacity': fill.opacity < 1 ? num(fill.opacity) : undefined,
        stroke:
          outline && outline.fill.type === 'solid'
            ? toCss({ ...outline.fill.color, a: 1 })
            : undefined,
        'stroke-width': outline ? num(outline.width / EMU_PER_PT) : undefined,
        'data-run': run.kind === 'bullet' ? 'bullet' : String(run.runIndex),
      })
    );
    if (w.underline.length) {
      const under = run.props.underlineFill ? textFill(run.props.underlineFill, bounds, ctx) : fill;
      group.appendChild(
        svg(doc, 'path', {
          d: pointPathData(w.underline),
          fill: under.paint,
          'fill-opacity': under.opacity < 1 ? num(under.opacity) : undefined,
        })
      );
    }
  }
}

/**
 * The text style of a body: no kerning and no ligatures, since each character sits at its
 * measured place. Every property is inherited, so the runs take it from the body.
 */
const TEXT_STYLE =
  'font-kerning:none;font-variant-ligatures:none;font-feature-settings:"liga" 0,"clig" 0,"dlig" 0,"calt" 0';

/**
 * Whether a box can show: `corners` are its four corners in the text frame's space, in EMU.
 * Paint skips a line that cannot show (`PaintOptions.view`).
 */
export type Shows = (corners: readonly { x: number; y: number }[]) => boolean;

/**
 * The corners of `line` in the text frame's space, padded on every side by the line's height
 * and its largest font size, so a glyph that reaches past its line box (a tall accent, a
 * superscript, an italic overhang) still counts.
 */
function lineCorners(
  layout: TextLayout,
  line: TextLayout['lines'][number]
): { x: number; y: number }[] {
  let left = Infinity;
  let right = -Infinity;
  let size = 0;
  let top = line.top;
  let bottom = line.top + line.height;
  for (const run of line.runs) {
    if (!run.x.length) continue;
    left = Math.min(left, run.x[0], run.x[run.x.length - 1]);
    right = Math.max(right, run.x[0], run.x[run.x.length - 1]);
    size = Math.max(size, run.sizePt * EMU_PER_PT);
    // A picture bullet stands on the baseline and can be taller than the line.
    if (run.picture) top = Math.min(top, line.baseline - run.picture.height);
  }
  // An equation's marks (a radical sign, a stretched delimiter) can reach past its glyphs.
  for (const mark of line.marks ?? [])
    for (const c of mark.commands) {
      if (c.op === 'Z') continue;
      left = Math.min(left, c.x);
      right = Math.max(right, c.x);
      top = Math.min(top, line.baseline - c.y);
      bottom = Math.max(bottom, line.baseline - c.y);
    }
  const pad = Math.max(line.height, size) * 2;
  const box = layout.box;
  const x0 = box.x + left - pad;
  const x1 = box.x + right + pad;
  const y0 = box.y + top - pad;
  const y1 = box.y + bottom + pad;
  const corners = [
    { x: x0, y: y0 },
    { x: x1, y: y0 },
    { x: x1, y: y1 },
    { x: x0, y: y1 },
  ];
  if (layout.rotation === 0) return corners;
  // As the paint: rotated about the box center.
  const a = (layout.rotation * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  return corners.map(({ x, y }) => ({
    x: cx + (x - cx) * cos - (y - cy) * sin,
    y: cy + (x - cx) * sin + (y - cy) * cos,
  }));
}

/** Whether any line of `layout` can show. Warped text always can: the warp moves its lines. */
export function textShows(layout: TextLayout, shows: Shows): boolean {
  if (layout.warped) return true;
  return layout.lines.some((line) => shows(lineCorners(layout, line)));
}

/**
 * Paint a text layout into one group in the text frame's coordinates, or under `frame` (an
 * SVG transform from the frame to the parent's space). With `shows`, a line that cannot show
 * is left out: a text that overflows its slide by thousands of lines paints the lines in view
 * only.
 */
export function paintText(
  doc: Document,
  layout: TextLayout,
  ctx: PaintContext,
  shows?: Shows,
  frame?: string
): SVGElement {
  const box = layout.box;
  // One group places the body (the frame's transform, then the box, then its rotation) and
  // scales points to EMU: browsers clamp very large font sizes, so text paints in points. The
  // text style is inherited: one declaration for the body, not one for each run, so a slide of
  // thousands of runs does not parse it thousands of times.
  const place =
    layout.rotation !== 0
      ? `translate(${num(box.x)} ${num(box.y)}) rotate(${num(layout.rotation)} ${num(box.width / 2)} ${num(box.height / 2)})`
      : `translate(${num(box.x)} ${num(box.y)})`;
  const group = svg(doc, 'g', {
    transform: `${frame ? `${frame} ` : ''}${place} scale(${EMU_PER_PT})`,
    'data-pptx-text': '1',
  });
  group.setAttribute('style', TEXT_STYLE);
  const pt = (emu: number): string => num(emu / EMU_PER_PT);
  const bounds = textBounds(layout);
  // Warped glyphs lie where the warp maps them, not in their lines' boxes: every one paints.
  if (layout.warped) {
    paintWarped(doc, layout, layout.warped, group, bounds, ctx);
    return group;
  }
  for (const line of layout.lines) {
    if (shows && !shows(lineCorners(layout, line))) continue;
    const unders = lineUnderlines(line.runs, line.baseline);
    for (const run of line.runs) {
      if (run.picture) {
        const image = bulletPicture(doc, run, line.baseline, ctx);
        if (image) group.appendChild(image);
        continue;
      }
      if (run.text.length === 0) continue;
      const fill = textFill(run.props.fill, bounds, ctx);
      if (run.props.highlight && run.kind === 'text') {
        group.appendChild(
          svg(doc, 'rect', {
            x: pt(run.x[0]),
            y: pt(line.top),
            width: pt(run.x[run.x.length - 1] - run.x[0]),
            height: pt(line.height),
            fill: toCss({ ...run.props.highlight, a: 1 }),
          })
        );
      }
      if (fill.paint === 'none' && !run.props.outline) continue;
      const outline = run.props.outline;
      const baseline = line.baseline - run.rise;
      const makeText = (content: string, attrs: Record<string, string | undefined>): SVGElement => {
        const text = svg(doc, 'text', {
          ...attrs,
          // `sans-serif` only draws glyphs the face lacks; measurement used the face itself.
          'font-family': `${run.face.id}, sans-serif`,
          'font-size': num(run.sizePt),
          'font-weight': run.face.synthBold ? 'bold' : undefined,
          'font-style': run.face.synthItalic ? 'italic' : undefined,
          fill: fill.paint,
          'fill-opacity': fill.opacity < 1 ? num(fill.opacity) : undefined,
          stroke:
            outline && outline.fill.type === 'solid'
              ? toCss({ ...outline.fill.color, a: 1 })
              : undefined,
          'stroke-width': outline ? pt(outline.width) : undefined,
          'xml:space': 'preserve',
          'data-run': run.kind === 'bullet' ? 'bullet' : String(run.runIndex),
        });
        // Kept on each run: inherited from the body, the browser's PDF splits a spaced run
        // into pieces (measured on the `font-effects` fixture).
        text.setAttribute('style', 'white-space:pre');
        text.textContent = content;
        return text;
      };
      // Upright characters in a rotated line turn back around their advance box centers, on
      // the axis `cy` across the line. An East Asian character (`eaVert`) centers on the line
      // box and centers its em box in its advance: the ideographic em box ends 0.12 em below
      // the baseline (OpenType `BASE`, the default `ideo` baseline), so the baseline is
      // 0.38 em below the center. A stacked letter centers on the axis that the layout puts
      // in `baseline`, and its baseline sits at the face's ascent share of its square cell.
      // Reference PDF of `vertical-east-asian`: the characters are within 0.25 pt.
      const m = run.face.metrics;
      const em = run.sizePt * EMU_PER_PT;
      const stacked = layout.upright === 'all';
      const cy = stacked ? baseline : line.top + line.height / 2 - run.rise;
      const share = m.ascent / Math.max(0.01, m.ascent + m.descent);
      const drop = (i: number): number =>
        stacked ? (share - 0.5) * (run.x[i + 1] - run.x[i]) : 0.38 * em;
      // A stretch of upright characters paints as one `<text>` in a frame turned back by 90°:
      // there, the turn around each center (cx, cy) is a move to (-cy, cx + drop).
      // Each character has its own position, so `text-anchor` centers each one by itself.
      // The pixels are those of one turned `<text>` per character, at a fraction of the
      // elements. Only characters of one UTF-16 unit stack: a position list counts
      // characters, which a surrogate pair would shift.
      // The stack waiting to paint: characters `from` to `to`, none when `from` is -1.
      let from = -1;
      let to = -1;
      const flush = (): void => {
        if (from < 0) return;
        const xs: string[] = [];
        const ys: string[] = [];
        for (let i = from; i < to; i += 1) {
          xs.push(pt(-cy));
          ys.push(pt((run.x[i] + run.x[i + 1]) / 2 + drop(i)));
        }
        group.appendChild(
          svg(doc, 'g', { transform: 'rotate(-90)' }, [
            makeText(run.text.slice(from, to), {
              x: xs.join(' '),
              y: ys.join(' '),
              'text-anchor': 'middle',
            }),
          ])
        );
        from = -1;
      };
      const kerning = run.props.kern !== undefined && run.props.size >= run.props.kern;
      for (const word of run.rtl ? rtlWords(run) : []) {
        const text = makeText(word.text, {
          x: pt(word.right),
          y: pt(baseline),
          textLength: word.width > 0 ? pt(word.width) : undefined,
          lengthAdjust: word.width > 0 ? 'spacing' : undefined,
        });
        text.setAttribute('style', rtlStyle(kerning));
        group.appendChild(text);
      }
      for (const piece of run.rtl ? [] : uprightPieces(run.text, layout.upright)) {
        const content = run.text.slice(piece.start, piece.end);
        // A space draws nothing, so it can join a stack in any direction.
        const stacks =
          !/[\uD800-\uDFFF]/.test(content) &&
          (piece.upright || (from >= 0 && /^ +$/.test(content)));
        if (stacks) {
          if (from < 0) from = piece.start;
          to = piece.end;
          continue;
        }
        flush();
        if (!piece.upright) {
          const xs = run.x.slice(piece.start, piece.end).map(pt).join(' ');
          group.appendChild(makeText(content, { x: xs, y: pt(baseline) }));
          continue;
        }
        const cx = (run.x[piece.start] + run.x[piece.end]) / 2;
        group.appendChild(
          makeText(content, {
            x: pt(cx),
            y: pt(cy + drop(piece.start)),
            'text-anchor': 'middle',
            transform: `rotate(-90 ${pt(cx)} ${pt(cy)})`,
          })
        );
      }
      flush();
      if (run.kind === 'text') {
        const under = run.props.underlineFill
          ? textFill(run.props.underlineFill, bounds, ctx)
          : fill;
        const rings = unders.get(run);
        for (const d of decorationRects(doc, run, line.baseline, fill, under, rings))
          group.appendChild(d);
      }
    }
    if (line.marks) paintMarks(doc, line, group, bounds, ctx);
  }
  return group;
}
