// Text paint. Every glyph run becomes one `<text>` whose `x` list carries the layout's
// position for each character, so the browser draws exactly where HarfBuzz measured and never
// re-breaks or re-spaces a line. Underline, strikethrough, and highlight are drawn as shapes
// from the face metrics, not by CSS text decoration, so they match the measured geometry.

import { toCss, type ResolvedFill } from '../cascade/index.ts';
import { isIdeographic, type GlyphRun, type TextLayout } from '../layout/index.ts';
import { fillPaint, type PaintContext } from './paint-fill.ts';
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
  underline: TextPaint
): SVGElement[] {
  const out: SVGElement[] = [];
  const size = run.sizePt;
  const baseline = baselineEmu / EMU_PER_PT;
  const xs = run.x.map((v) => v / EMU_PER_PT);
  const x0 = xs[0];
  // Trailing spaces at a line end are not underlined; the layout already
  // excludes them from alignment, and runs keep them, so trim here.
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  const x1 = xs[end] ?? xs[xs.length - 1];
  if (x1 <= x0) return out;
  const m = run.face.metrics;
  const thickness = Math.max(m.underlineThickness * size, size / 20);
  const y = baseline - run.rise / EMU_PER_PT;
  // A strikethrough takes the run's fill and its alpha, as the glyphs do; an underline takes
  // its own fill (`a:uFill`) when the run has one.
  const bar = (top: number, { paint, opacity }: TextPaint): SVGElement =>
    svg(doc, 'rect', {
      x: num(x0),
      y: num(top),
      width: num(x1 - x0),
      height: num(thickness),
      fill: paint,
      'fill-opacity': opacity < 1 ? num(opacity) : undefined,
    });
  if (run.props.underline !== 'none') {
    const ty = y - m.underlinePosition * size;
    out.push(bar(ty, underline));
    if (run.props.underline === 'dbl') out.push(bar(ty + thickness * 2, underline));
  }
  if (run.props.strike !== 'noStrike') {
    const sy = y - size * 0.3;
    out.push(bar(sy - thickness / 2, text));
    if (run.props.strike === 'dblStrike') out.push(bar(sy - thickness * 2, text));
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
  for (const run of line.runs) {
    if (!run.x.length) continue;
    left = Math.min(left, run.x[0], run.x[run.x.length - 1]);
    right = Math.max(right, run.x[0], run.x[run.x.length - 1]);
    size = Math.max(size, run.sizePt * EMU_PER_PT);
  }
  const pad = Math.max(line.height, size) * 2;
  const box = layout.box;
  const x0 = box.x + left - pad;
  const x1 = box.x + right + pad;
  const y0 = box.y + line.top - pad;
  const y1 = box.y + line.top + line.height + pad;
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

/** Whether any line of `layout` can show. */
export function textShows(layout: TextLayout, shows: Shows): boolean {
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
  for (const line of layout.lines) {
    if (shows && !shows(lineCorners(layout, line))) continue;
    for (const run of line.runs) {
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
      // Upright characters in a rotated line turn back around their advance box centers.
      const m = run.face.metrics;
      const cy = baseline - ((m.ascent - m.descent) / 2) * run.sizePt * EMU_PER_PT;
      // A stretch of upright characters paints as one `<text>` in a frame turned back by 90°:
      // there, the turn around each center (cx, cy) is a move to (-cy, baseline + cx - cy).
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
          ys.push(pt(baseline + (run.x[i] + run.x[i + 1]) / 2 - cy));
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
      for (const piece of uprightPieces(run.text, layout.upright)) {
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
            y: pt(baseline),
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
        for (const d of decorationRects(doc, run, line.baseline, fill, under)) group.appendChild(d);
      }
    }
  }
  return group;
}
