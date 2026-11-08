// Text paint. Every glyph run becomes one `<text>` whose `x` list carries the layout's
// position for each character, so the browser draws exactly where HarfBuzz measured and never
// re-breaks or re-spaces a line. Underline, strikethrough, and highlight are drawn as shapes
// from the face metrics, not by CSS text decoration, so they match the measured geometry.

import { toCss, type ResolvedFill } from '../cascade/index.ts';
import {
  isIdeographic,
  strikeoutBars,
  type GlyphRun,
  type PathCommand,
  type TextLayout,
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
  underline: TextPaint
): SVGElement[] {
  const out: SVGElement[] = [];
  const size = run.sizePt;
  const baseline = baselineEmu / EMU_PER_PT;
  const xs = run.x.map((v) => v / EMU_PER_PT);
  // Trailing spaces at a line end are not underlined; the layout already
  // excludes them from alignment, and runs keep them, so trim here.
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  // A right-to-left run ends at its left: its trailing spaces are its leftmost cells.
  const x0 = run.rtl ? xs[run.text.length - end] : xs[0];
  const x1 = run.rtl ? xs[xs.length - 1] : (xs[end] ?? xs[xs.length - 1]);
  if (x1 <= x0) return out;
  const m = run.face.metrics;
  const thickness = Math.max(m.underlineThickness * size, size / 20);
  const y = baseline - run.rise / EMU_PER_PT;
  // A strikethrough takes the run's fill and its alpha, as the glyphs do; an underline takes
  // its own fill (`a:uFill`) when the run has one.
  const bar = (top: number, { paint, opacity }: TextPaint, height = thickness): SVGElement =>
    svg(doc, 'rect', {
      x: num(x0),
      y: num(top),
      width: num(x1 - x0),
      height: num(height),
      fill: paint,
      'fill-opacity': opacity < 1 ? num(opacity) : undefined,
    });
  if (run.props.underline !== 'none') {
    const ty = y - m.underlinePosition * size;
    out.push(bar(ty, underline));
    if (run.props.underline === 'dbl') out.push(bar(ty + thickness * 2, underline));
  }
  if (run.props.strike !== 'noStrike') {
    const strike = strikeoutBars(m, size, run.props.strike === 'dblStrike');
    for (const top of strike.tops) out.push(bar(y + top, text, strike.thickness));
  }
  return out;
}

/** A right-to-left word: its letters in source order, anchored at its right edge. */
const RTL_STYLE = 'direction:rtl;unicode-bidi:bidi-override';

/**
 * The words of a right-to-left run, each painted as one `<text>` that starts at the word's
 * right edge and reads right to left, so the browser shapes and joins its letters. Spaces
 * between the words take the room that layout gave them, justification included. Every
 * character of the run has an odd bidi level, so the override keeps the browser from
 * reordering the word again.
 */
function rtlWords(run: GlyphRun): { text: string; right: number }[] {
  const out: { text: string; right: number }[] = [];
  const n = run.text.length;
  const re = /[^ \u3000]+/gu;
  // Unit `i` draws in cell `n - 1 - i`, so a word that starts at unit `i` ends at `x[n - i]`.
  for (let m = re.exec(run.text); m; m = re.exec(run.text))
    out.push({ text: m[0], right: run.x[n - m.index] });
  return out;
}

/**
 * Split a run into stretches painted along the line and single characters painted upright.
 * Offsets are UTF-16; a surrogate pair stays together.
 */
export function uprightPieces(
  text: string,
  mode: TextLayout['upright']
): { start: number; end: number; upright: boolean }[] {
  if (mode === 'none') return [{ start: 0, end: text.length, upright: false }];
  const out: { start: number; end: number; upright: boolean }[] = [];
  let i = 0;
  while (i < text.length) {
    const cp = text.codePointAt(i) ?? 0;
    const len = cp > 0xffff ? 2 : 1;
    const upright = mode === 'all' ? text[i] !== ' ' : isIdeographic(cp);
    const last = out[out.length - 1];
    if (!upright && last && !last.upright) last.end = i + len;
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

/** Paint a text layout into a group in the text frame's coordinates. */
export function paintText(doc: Document, layout: TextLayout, ctx: PaintContext): SVGElement {
  const box = layout.box;
  const outer = svg(doc, 'g', {
    transform:
      layout.rotation !== 0
        ? `translate(${num(box.x)} ${num(box.y)}) rotate(${num(layout.rotation)} ${num(box.width / 2)} ${num(box.height / 2)})`
        : `translate(${num(box.x)} ${num(box.y)})`,
    'data-pptx-text': '1',
  });
  // Browsers clamp very large font sizes, so text paints in points under a scale transform.
  const group = svg(doc, 'g', { transform: `scale(${EMU_PER_PT})` });
  outer.appendChild(group);
  const pt = (emu: number): string => num(emu / EMU_PER_PT);
  const bounds = textBounds(layout);
  if (layout.warped) {
    paintWarped(doc, layout, layout.warped, group, bounds, ctx);
    return outer;
  }
  for (const line of layout.lines) {
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
        text.setAttribute(
          'style',
          'white-space:pre;font-kerning:none;font-variant-ligatures:none;font-feature-settings:"liga" 0,"clig" 0,"dlig" 0,"calt" 0'
        );
        text.textContent = content;
        return text;
      };
      for (const word of run.rtl ? rtlWords(run) : []) {
        const text = makeText(word.text, { x: pt(word.right), y: pt(baseline) });
        text.setAttribute('style', `${text.getAttribute('style')};${RTL_STYLE}`);
        group.appendChild(text);
      }
      for (const piece of run.rtl ? [] : uprightPieces(run.text, layout.upright)) {
        if (!piece.upright) {
          const xs = run.x.slice(piece.start, piece.end).map(pt).join(' ');
          group.appendChild(
            makeText(run.text.slice(piece.start, piece.end), { x: xs, y: pt(baseline) })
          );
          continue;
        }
        // An upright character in a rotated line: turn it back around its advance box center.
        const m = run.face.metrics;
        const cx = (run.x[piece.start] + run.x[piece.end]) / 2;
        const cy = baseline - ((m.ascent - m.descent) / 2) * run.sizePt * EMU_PER_PT;
        group.appendChild(
          makeText(run.text.slice(piece.start, piece.end), {
            x: pt(cx),
            y: pt(baseline),
            'text-anchor': 'middle',
            transform: `rotate(-90 ${pt(cx)} ${pt(cy)})`,
          })
        );
      }
      if (run.kind === 'text') {
        const under = run.props.underlineFill
          ? textFill(run.props.underlineFill, bounds, ctx)
          : fill;
        for (const d of decorationRects(doc, run, line.baseline, fill, under)) group.appendChild(d);
      }
    }
    if (line.marks) paintMarks(doc, line, group, bounds, ctx);
  }
  return outer;
}
