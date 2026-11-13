/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text as PDF text. Each character draws at the x position layout measured, so line breaks and
// positions are the engine's, and the text stays searchable.

import {
  appendBezierCurve,
  beginText,
  clip as clipPath,
  closePath,
  concatTransformationMatrix,
  endPath,
  endText,
  fill,
  fillAndStroke,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setFillingRgbColor,
  setFontAndSize,
  setLineJoin,
  setLineWidth,
  LineJoinStyle,
  setStrokingRgbColor,
  stroke,
  setTextMatrix,
  setTextRenderingMode,
  showText,
  PDFHexString,
  TextRenderingMode,
} from 'pdf-lib';
import type { ResolvedFill } from '@pptx-editor.dev/core/cascade';
import {
  isIdeographic,
  lineUnderlines,
  runGlyphs,
  strikeoutBars,
  type GlyphRun,
  type PathCommand,
  type TextLayout,
  type UnderlineRing,
  type WarpedRun,
} from '@pptx-editor.dev/core/layout';
import { EmbeddedFace } from './fonts.ts';
import { flatColor, gradientBox, ownFill, type FillScope } from './paint-fill.ts';
import { drawPicture } from './paint-picture.ts';
import { clip, clipNonZero, EMU_PER_PT, rgb, tracePath, type PageOps } from './page-ops.ts';

/** The embedded face for a layout face id, or null with a diagnostic. */
function fontFor(faceId: string, family: string, p: PageOps): EmbeddedFace | null {
  const cached = p.res.fonts.get(faceId);
  if (cached !== undefined) return cached;
  const face = p.res.faces.get(faceId);
  let font: EmbeddedFace | null = null;
  if (!face)
    p.diag.add(
      'font.missing',
      'unsupported',
      `The ${clip(family)} font has no font bytes. Its text is not drawn.`
    );
  else {
    try {
      // DOCX Editor's embedder: OS/2 fsType and collection checks, then a subset per face.
      font = new EmbeddedFace(
        p.res.doc,
        { identity: face.id, bytes: face.bytes, faceIndex: 0, request: { family: face.family } },
        p.res.fonts.size
      );
    } catch (error) {
      p.diag.add(
        'font.not-embedded',
        'unsupported',
        `The ${clip(face.family)} font cannot be embedded: ${(error as Error).message}. Its text is not drawn.`
      );
    }
  }
  p.res.fonts.set(faceId, font);
  return font;
}

/** Enter the text frame: its box offset, then its rotation about the box center. */
function enterFrame(layout: TextLayout, p: PageOps): void {
  const box = layout.box;
  p.push(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, box.x, box.y));
  if (layout.rotation !== 0) {
    const r = (layout.rotation * Math.PI) / 180;
    const [cx, cy] = [box.width / 2, box.height / 2];
    p.push(concatTransformationMatrix(1, 0, 0, 1, cx, cy));
    p.push(concatTransformationMatrix(Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0));
    p.push(concatTransformationMatrix(1, 0, 0, 1, -cx, -cy));
  }
}

/**
 * Whether a character of a vertical line draws upright, as `uprightPieces` of the SVG paint
 * decides: every character but a space in stacked text, and an ideograph in `eaVert` text.
 */
function drawsUpright(ch: string, mode: TextLayout['upright']): boolean {
  if (mode === 'all') return ch !== ' ' && ch !== '\t';
  return mode === 'eastAsian' && isIdeographic(ch.codePointAt(0) ?? 0);
}

/**
 * The text-space origin of an upright glyph in a turned line, for a glyph whose cell runs from
 * `x0` to `x1` and that is `advance` wide (EMU). As in the SVG paint, the glyph turns back
 * around its cell center on the axis across the line. An East Asian character centers on the
 * line box, and its baseline is 0.38 em below the center (the ideographic em box ends 0.12 em
 * below the baseline). A stacked letter centers on the line's baseline, and its baseline sits
 * at the face's ascent share of its cell. Turned back by 90°, the glyph's baseline runs across
 * the line, and its advance runs up the frame, so it starts `advance / 2` below the axis.
 */
function uprightOrigin(
  mode: TextLayout['upright'],
  line: Pick<TextLayout['lines'][number], 'top' | 'height' | 'baseline'>,
  run: Pick<GlyphRun, 'rise' | 'sizePt' | 'face'>,
  x0: number,
  x1: number,
  advance: number
): [number, number] {
  const stacked = mode === 'all';
  const m = run.face.metrics;
  const share = m.ascent / Math.max(0.01, m.ascent + m.descent);
  const cy = stacked ? line.baseline - run.rise : line.top + line.height / 2 - run.rise;
  const drop = stacked ? (share - 0.5) * (x1 - x0) : 0.38 * run.sizePt * EMU_PER_PT;
  return [(x0 + x1) / 2 + drop, cy + advance / 2];
}

/** Draw a text layout in its frame's space. `scope` resolves group fills. */
export async function drawText(layout: TextLayout, p: PageOps, scope: FillScope): Promise<void> {
  enterFrame(layout, p);
  if (layout.warped) {
    await drawWarped(layout, layout.warped, p, scope);
    p.push(popGraphicsState());
    return;
  }
  let bounds: ReturnType<typeof textBounds> | undefined;
  const markBounds: { box?: ReturnType<typeof textBounds> } = {};
  for (const line of layout.lines) {
    const unders = lineUnderlines(line.runs, line.baseline);
    const rtlOnly = line.runs.some((r) => r.rtl) && line.runs.every((r) => r.rtl || !r.text.trim());
    // Emit visual order across font/style boundaries too, so readers infer word spaces.
    const runs = line.runs.some((r) => r.rtl)
      ? [...line.runs].sort((a, b) => a.x[0] - b.x[0])
      : line.runs;
    for (const run of runs) {
      if (run.picture) {
        await drawBulletPicture(run, line.baseline, p);
        continue;
      }
      if (run.text.length === 0) continue;
      if (run.props.highlight && run.kind === 'text') {
        const x0 = run.x[0];
        const x1 = run.x[run.x.length - 1];
        p.push(setFillingRgbColor(...rgb(run.props.highlight)));
        p.push(moveTo(x0, line.top), lineTo(x1, line.top), lineTo(x1, line.top + line.height));
        p.push(lineTo(x0, line.top + line.height), closePath(), fill());
      }
      // A gradient clips to the glyphs and spans the text body's bounds, as in the SVG paint.
      const own = ownFill(run.props.fill, scope).fill;
      const gradient = own.type === 'linear' || own.type === 'path' ? own : null;
      // As in the SVG paint: a run with no fill and no outline draws no bars either.
      if (own.type === 'none' && !run.props.outline) continue;
      const runColor = gradient ? null : flatColor(run.props.fill, scope, p, 'text');
      // The bars draw over the glyphs, as in the SVG paint, and also on outline-only text
      // and on text whose font has no bytes.
      const decorate = (): void => {
        if (run.kind === 'text')
          drawDecorations(
            run,
            line.baseline,
            scope,
            p,
            () => (bounds ??= textBounds(layout)),
            unders.get(run)
          );
      };
      if (run.props.outline)
        p.diag.add('text.outline', 'unsupported', 'Text outlines are not drawn yet.');
      if (!gradient && !runColor) {
        decorate();
        continue;
      }
      if (run.face.synthBold || run.face.synthItalic)
        p.diag.add(
          'text.synthetic-style',
          'approximation',
          'Synthetic bold and italic are not drawn.'
        );
      const font = fontFor(run.face.id, run.face.family, p);
      if (!font) {
        decorate();
        continue;
      }
      if (run.rtl && !run.glyphs)
        p.diag.add(
          'text.shaping',
          'approximation',
          'Contextual glyphs are unavailable; text uses nominal glyphs.'
        );
      const key = p.fontKey(run.face.id, font);
      const size = run.sizePt * EMU_PER_PT;
      // Stacked text: the line's baseline is the axis through the letter centers (core's
      // `columnLines`). The turned letters center their ascent and descent on it, as the
      // upright letters of the SVG paint center on it.
      const m = run.face.metrics;
      const across = layout.upright === 'all' ? ((m.ascent - m.descent) / 2) * size : 0;
      const y = line.baseline - run.rise + across;
      // The font size goes in `Tf`, not in the text matrix: a reader starts a new text item where
      // the size changes, so a smaller subscript or superscript run reads as its own item, as
      // in the editor's printed PDF. With one size for all runs, a reader joins a subscript to
      // the text around it ("H2O" in place of "H", "2", and "O").
      p.push(pushGraphicsState(), beginText(), setFontAndSize(key, size));
      if (runColor) {
        p.alpha(runColor.a, 1);
        p.push(setFillingRgbColor(...rgb(runColor)));
      } else p.push(setTextRenderingMode(TextRenderingMode.Clip));
      let shown = 0;
      const extra: { id: number; x: number; y: number; scaleX: number }[] = [];
      const glyphs =
        run.glyphs?.map((g) => ({ ...g, x: run.x[0] + g.x, index: -1 })) ??
        [...runGlyphs(run)].map(({ index, char, x }) => ({
          id: font.font.glyphForCodePoint(char.codePointAt(0) ?? 0)?.id ?? 0,
          text: char,
          x,
          y: 0,
          scaleX: 1,
          index,
        }));
      for (const { id: glyph, text: ch, x, y: dy, scaleX, index } of glyphs) {
        if (rtlOnly && /^ +$/.test(ch)) continue;
        if (ch === '') {
          if (glyph > 0) extra.push({ id: glyph, x, y: y + dy, scaleX });
          continue;
        }
        if (glyph <= 0) {
          if (ch.trim() !== '')
            p.diag.add(
              'text.missing-glyph',
              'approximation',
              'Characters the face lacks are not drawn.'
            );
          continue;
        }
        const encoded = (): PDFHexString =>
          // Readers reorder visual RTL text, including the characters inside one cluster.
          PDFHexString.of(font.encode(glyph, run.rtl ? Array.from(ch).reverse().join('') : ch));
        if (!run.rtl && index >= 0 && drawsUpright(ch, layout.upright)) {
          const end = run.x[index + ch.length] ?? x;
          const advance = (font.font.getGlyph(glyph).advanceWidth / font.font.unitsPerEm) * size;
          const [ux, uy] = uprightOrigin(layout.upright, line, run, x, end, advance * scaleX);
          // The glyph turns back by 90° in the turned frame: its advance runs up the frame and
          // its ascent points against the line, as in the SVG paint.
          p.push(setTextMatrix(0, -scaleX, -1, 0, ux, uy), showText(encoded()));
          shown += 1;
          continue;
        }
        // Text space: y up again inside the flipped page, at the size layout measured.
        p.push(setTextMatrix(scaleX, 0, 0, -1, x, y + dy), showText(encoded()));
        shown += 1;
      }
      p.push(endText());
      // A clip of no glyphs clips nothing in some readers, so a run without glyphs paints none.
      if (gradient && shown > 0) {
        bounds ??= textBounds(layout);
        p.push(concatTransformationMatrix(1, 0, 0, 1, bounds.x, bounds.y));
        gradientBox(gradient, bounds.w, bounds.h, p);
      }
      p.push(popGraphicsState());
      // Marks without their own text use outlines, so extraction never duplicates a cluster.
      for (const g of extra) {
        p.push(pushGraphicsState());
        if (runColor) {
          p.alpha(runColor.a, 1);
          p.push(setFillingRgbColor(...rgb(runColor)));
        }
        traceGlyph(font, g.id, g.x, g.y, size / font.font.unitsPerEm, g.scaleX, p);
        if (gradient) {
          p.push(clipNonZero(), endPath());
          bounds ??= textBounds(layout);
          p.push(concatTransformationMatrix(1, 0, 0, 1, bounds.x, bounds.y));
          gradientBox(gradient, bounds.w, bounds.h, p);
        } else p.push(fill());
        p.push(popGraphicsState());
      }
      decorate();
    }
    if (line.marks) drawMarks(layout, line, p, scope, markBounds);
  }
  p.push(popGraphicsState());
}

/**
 * Fill a path of a text body in a text fill: a gradient spans the body's laid-out bounds, as in
 * the SVG paint. `bounds` keeps the bounds once measured.
 */
function fillTextPath(
  layout: TextLayout,
  fillValue: ResolvedFill,
  commands: readonly PathCommand[],
  p: PageOps,
  scope: FillScope,
  bounds: { box?: ReturnType<typeof textBounds> }
): void {
  if (commands.length === 0) return;
  const own = ownFill(fillValue, scope).fill;
  if (own.type === 'linear' || own.type === 'path') {
    if (own.stops.length === 0) return;
    const box = (bounds.box ??= textBounds(layout));
    p.push(pushGraphicsState());
    tracePath(commands, p);
    p.push(clipNonZero(), endPath(), concatTransformationMatrix(1, 0, 0, 1, box.x, box.y));
    gradientBox(own, box.w, box.h, p);
    p.push(popGraphicsState());
    return;
  }
  const color = flatColor(fillValue, scope, p, 'text');
  if (!color) return;
  p.push(pushGraphicsState());
  p.alpha(color.a, 1);
  p.push(setFillingRgbColor(...rgb(color)));
  tracePath(commands, p);
  p.push(fill(), popGraphicsState());
}

/**
 * The marks of a line's equations (fraction rules, radical signs, stretched delimiters,
 * braces), filled in the math run's fill, as in the SVG paint.
 */
function drawMarks(
  layout: TextLayout,
  line: TextLayout['lines'][number],
  p: PageOps,
  scope: FillScope,
  bounds: { box?: ReturnType<typeof textBounds> }
): void {
  for (const mark of line.marks ?? []) {
    const commands = mark.commands.map((c): PathCommand => {
      if (c.op === 'Z') return c;
      if (c.op === 'Q') return { ...c, y1: line.baseline - c.y1, y: line.baseline - c.y };
      if (c.op === 'C')
        return { ...c, y1: line.baseline - c.y1, y2: line.baseline - c.y2, y: line.baseline - c.y };
      return { ...c, y: line.baseline - c.y };
    });
    fillTextPath(layout, mark.props.fill, commands, p, scope, bounds);
  }
}

/**
 * Warped text (`a:prstTxWarp`) as the paths layout mapped into the warp, as in the SVG paint
 * (`core/src/output/paint-text.ts`, `paintWarped`): the highlight, then the glyphs in the run's
 * fill (nonzero, as glyph outlines fill), their solid outline, then the underline. A picture
 * bullet draws on the parallelogram that the warp gives it. Paths, not text, so the warped
 * words are not in the PDF text layer.
 */
async function drawWarped(
  layout: TextLayout,
  warped: readonly WarpedRun[],
  p: PageOps,
  scope: FillScope
): Promise<void> {
  const bounds: { box?: ReturnType<typeof textBounds> } = {};
  const paint = (fillValue: ResolvedFill, commands: readonly PathCommand[]): void =>
    fillTextPath(layout, fillValue, commands, p, scope, bounds);
  for (const w of warped) {
    const run = layout.lines[w.line]?.runs[w.run];
    if (!run) continue;
    if (run.picture) {
      if (w.picture) await drawWarpedPicture(run, w.picture, p);
      continue;
    }
    if (run.props.highlight && w.highlight.length) {
      p.push(pushGraphicsState(), setFillingRgbColor(...rgb(run.props.highlight)));
      tracePath(w.highlight, p);
      p.push(fill(), popGraphicsState());
    }
    paint(run.props.fill, w.glyphs);
    // The outline strokes the warped glyphs over their fill, as the SVG paint does: a solid
    // outline in its color, any other outline not at all.
    const outline = run.props.outline;
    if (outline && outline.fill.type === 'solid' && w.glyphs.length) {
      p.push(pushGraphicsState(), setStrokingRgbColor(...rgb(outline.fill.color)));
      p.push(setLineWidth(outline.width));
      tracePath(w.glyphs, p);
      p.push(stroke(), popGraphicsState());
    }
    paint(run.props.underlineFill ?? run.props.fill, w.underline);
  }
}

/** A picture bullet of warped text, stretched over the parallelogram the warp gives it. */
async function drawWarpedPicture(
  run: GlyphRun,
  box: NonNullable<WarpedRun['picture']>,
  p: PageOps
): Promise<void> {
  const pic = run.picture;
  const w = run.x[run.x.length - 1] - run.x[0];
  if (!pic || !(w > 0) || !(pic.height > 0)) return;
  const { origin, across, down } = box;
  p.push(
    pushGraphicsState(),
    concatTransformationMatrix(
      across[0] / w,
      across[1] / w,
      down[0] / pic.height,
      down[1] / pic.height,
      origin[0],
      origin[1]
    )
  );
  await drawPicture(pic.part, bulletFill(pic), w, pic.height, p);
  p.push(popGraphicsState());
}

/**
 * Underline and strikethrough as filled shapes, with the geometry of the SVG paint
 * (`core/src/output/paint-text.ts`, `decorationRects`): the underline shapes of the line
 * (`lineUnderlines`) in their own fill (`a:uFill`) when the run has one and stroked with a solid
 * underline outline (`a:uLn`), and the strikethrough bars in the run's fill, trailing spaces
 * trimmed. A gradient fill spans the text body's bounds, as the glyphs' does.
 */
function drawDecorations(
  run: GlyphRun,
  baseline: number,
  scope: FillScope,
  p: PageOps,
  bounds: () => { x: number; y: number; w: number; h: number },
  underline: readonly UnderlineRing[] | undefined
): void {
  const { strike } = run.props;
  const shapes = (fillValue: ResolvedFill, rings: readonly UnderlineRing[]): void => {
    const own = ownFill(fillValue, scope).fill;
    const trace = (): void => {
      for (const ring of rings) {
        ring.forEach(([x, y], i) => p.push(i ? lineTo(x, y) : moveTo(x, y)));
        p.push(closePath());
      }
    };
    if (own.type === 'solid') {
      p.push(pushGraphicsState());
      p.alpha(own.color.a, 1);
      p.push(setFillingRgbColor(...rgb(own.color)));
      trace();
      p.push(fill(), popGraphicsState());
    } else if ((own.type === 'linear' || own.type === 'path') && own.stops.length > 0) {
      const box = bounds();
      p.push(pushGraphicsState());
      trace();
      p.push(clipPath(), endPath(), concatTransformationMatrix(1, 0, 0, 1, box.x, box.y));
      gradientBox(own, box.w, box.h, p);
      p.push(popGraphicsState());
    } else if (own.type !== 'none') {
      // A picture or pattern fill is not drawn on text yet: this reports it once.
      flatColor(fillValue, scope, p, 'text');
    }
  };
  if (underline?.length) {
    shapes(run.props.underlineFill ?? run.props.fill, underline);
    const line = run.props.underlineLine;
    if (line && line.fill.type === 'solid' && line.width > 0) {
      p.push(pushGraphicsState());
      p.alpha(1, line.fill.color.a);
      p.push(setStrokingRgbColor(...rgb(line.fill.color)), setLineWidth(line.width));
      p.push(setLineJoin(JOINS[line.join]));
      for (const ring of underline) {
        ring.forEach(([x, y], i) => p.push(i ? lineTo(x, y) : moveTo(x, y)));
        p.push(closePath());
      }
      p.push(stroke(), popGraphicsState());
    }
  }
  if (strike === 'noStrike') return;
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  const x0 = run.rtl ? run.x[run.text.length - end] : run.x[0];
  const x1 = run.rtl ? run.x[run.x.length - 1] : (run.x[end] ?? run.x[run.x.length - 1]);
  if (!(x1 > x0)) return;
  const y = baseline - run.rise;
  const bar = strikeoutBars(run.face.metrics, run.sizePt * EMU_PER_PT, strike === 'dblStrike');
  shapes(
    run.props.fill,
    bar.tops.map((top): UnderlineRing => {
      const [a, b] = [y + top, y + top + bar.thickness];
      return [
        [x0, a],
        [x1, a],
        [x1, b],
        [x0, b],
      ];
    })
  );
}

const JOINS = {
  round: LineJoinStyle.Round,
  bevel: LineJoinStyle.Bevel,
  miter: LineJoinStyle.Miter,
} as const;

/**
 * The picture of a picture bullet, standing on the baseline, stretched over its box as the SVG
 * paint draws it. A picture that cannot be read draws nothing.
 */
async function drawBulletPicture(run: GlyphRun, baseline: number, p: PageOps): Promise<void> {
  const pic = run.picture;
  if (!pic) return;
  const x = run.x[0];
  const w = run.x[run.x.length - 1] - x;
  if (!(w > 0) || !(pic.height > 0)) return;
  p.push(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, x, baseline - pic.height));
  await drawPicture(pic.part, bulletFill(pic), w, pic.height, p);
  p.push(popGraphicsState());
}

/** The picture of a picture bullet as a stretched picture fill. */
function bulletFill(
  pic: NonNullable<GlyphRun['picture']>
): Extract<ResolvedFill, { type: 'picture' }> {
  const none = { l: 0, t: 0, r: 0, b: 0 };
  return {
    type: 'picture',
    rId: pic.rId,
    ...(pic.svgRId ? { svgRId: pic.svgRId } : {}),
    srcRect: none,
    mode: 'stretch',
    fillRect: none,
    effects: [],
    rotWithShape: true,
  };
}

/** The laid-out bounds of a text body in its frame, as the SVG paint's gradient text spans. */
function textBounds(layout: TextLayout): { x: number; y: number; w: number; h: number } {
  let [x, y, r, b] = [Infinity, Infinity, -Infinity, -Infinity];
  for (const line of layout.lines) {
    y = Math.min(y, line.top);
    b = Math.max(b, line.top + line.height);
    for (const run of line.runs)
      if (run.x.length) {
        x = Math.min(x, run.x[0]);
        r = Math.max(r, run.x[run.x.length - 1]);
      }
  }
  if (!Number.isFinite(x) || !Number.isFinite(y))
    return { x: 0, y: 0, w: EMU_PER_PT, h: EMU_PER_PT };
  return { x, y, w: Math.max(r - x, EMU_PER_PT), h: Math.max(b - y, EMU_PER_PT) };
}

/**
 * Fill the glyph outlines of a text layout in the current color, for effect silhouettes, and
 * with `strokeToo` stroke them with the inherited line width to grow them. Outlines, not text,
 * so a shadow never adds a second copy of the words to the PDF text layer. A run that paints
 * nothing (no fill) casts nothing, as in the SVG paint.
 */
export function fillTextOutlines(
  layout: TextLayout,
  p: PageOps,
  scope: FillScope,
  strokeToo = false
): void {
  enterFrame(layout, p);
  // Warped text casts the warped glyphs.
  for (const w of layout.warped ?? []) {
    const run = layout.lines[w.line]?.runs[w.run];
    if (!run || w.glyphs.length === 0) continue;
    if (ownFill(run.props.fill, scope).fill.type === 'none') continue;
    tracePath(w.glyphs, p);
    p.push(strokeToo ? fillAndStroke() : fill());
  }
  for (const line of layout.warped ? [] : layout.lines) {
    for (const run of line.runs) {
      if (run.text.length === 0) continue;
      if (ownFill(run.props.fill, scope).fill.type === 'none') continue;
      const face = p.res.fonts.get(run.face.id) ?? fontFor(run.face.id, run.face.family, p);
      if (!face) continue;
      const k = (run.sizePt * EMU_PER_PT) / face.font.unitsPerEm;
      const y = line.baseline - run.rise;
      const glyphs =
        run.glyphs?.map((g) => ({ ...g, x: run.x[0] + g.x })) ??
        [...runGlyphs(run)].map(({ char, x }) => ({
          id: face.font.glyphForCodePoint(char.codePointAt(0) ?? 0)?.id ?? 0,
          x,
          y: 0,
          scaleX: 1,
        }));
      for (const g of glyphs) {
        if (g.id <= 0) continue;
        traceGlyph(face, g.id, g.x, y + g.y, k, g.scaleX, p);
        p.push(strokeToo ? fillAndStroke() : fill());
      }
    }
  }
  p.push(popGraphicsState());
}

/** Trace an admitted glyph in page units, including contextual mark offsets. */
function traceGlyph(
  face: EmbeddedFace,
  id: number,
  x: number,
  y: number,
  k: number,
  scaleX: number,
  p: PageOps
): void {
  const glyph = face.font.getGlyph(id);
  // Font units, y up → EMU, y down, at the glyph's origin.
  const px = (gx: number): number => x + gx * k * scaleX;
  const py = (gy: number): number => y - gy * k;
  let [cx, cy] = [0, 0];
  for (const c of glyph.path.commands) {
    const a = c.args;
    if (c.command === 'moveTo') p.push(moveTo(px(a[0]), py(a[1])));
    else if (c.command === 'lineTo') p.push(lineTo(px(a[0]), py(a[1])));
    else if (c.command === 'bezierCurveTo')
      p.push(appendBezierCurve(px(a[0]), py(a[1]), px(a[2]), py(a[3]), px(a[4]), py(a[5])));
    else if (c.command === 'quadraticCurveTo')
      p.push(
        appendBezierCurve(
          px(cx + ((a[0] - cx) * 2) / 3),
          py(cy + ((a[1] - cy) * 2) / 3),
          px(a[2] + ((a[0] - a[2]) * 2) / 3),
          py(a[3] + ((a[1] - a[3]) * 2) / 3),
          px(a[2]),
          py(a[3])
        )
      );
    else p.push(closePath());
    if (c.command !== 'closePath') [cx, cy] = [a[a.length - 2], a[a.length - 1]];
  }
}
