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
  setTextMatrix,
  setTextRenderingMode,
  showText,
  PDFHexString,
  TextRenderingMode,
} from 'pdf-lib';
import type { ResolvedFill } from '@pptx-editor.dev/core/cascade';
import { hasCharacters, type GlyphRun, type TextLayout } from '@pptx-editor.dev/core/layout';
import { EmbeddedFace } from './fonts.ts';
import { flatColor, gradientBox, ownFill, type FillScope } from './paint-fill.ts';
import { clip, EMU_PER_PT, rgb, type PageOps } from './page-ops.ts';

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

/** Draw a text layout in its frame's space. `scope` resolves group fills. */
export async function drawText(layout: TextLayout, p: PageOps, scope: FillScope): Promise<void> {
  enterFrame(layout, p);
  let bounds: ReturnType<typeof textBounds> | undefined;
  if (layout.upright !== 'none' && hasCharacters(layout))
    p.diag.add('text.upright', 'approximation', 'Upright vertical characters are drawn rotated.');
  for (const line of layout.lines) {
    for (const run of line.runs) {
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
          drawDecorations(run, line.baseline, scope, p, () => (bounds ??= textBounds(layout)));
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
      const key = p.fontKey(run.face.id, font);
      const size = run.sizePt * EMU_PER_PT;
      const y = line.baseline - run.rise;
      p.push(pushGraphicsState(), beginText(), setFontAndSize(key, 1));
      if (runColor) {
        p.alpha(runColor.a, 1);
        p.push(setFillingRgbColor(...rgb(runColor)));
      } else p.push(setTextRenderingMode(TextRenderingMode.Clip));
      let shown = 0;
      for (let i = 0; i < run.text.length; i += 1) {
        const code = run.text.charCodeAt(i);
        const pair = code >= 0xd800 && code <= 0xdbff && i + 1 < run.text.length;
        const ch = run.text.slice(i, pair ? i + 2 : i + 1);
        const x = run.x[i];
        if (pair) i += 1;
        // The glyph the face's cmap gives, as the editor's SVG text draws it (no ligatures).
        const glyph = font.font.glyphForCodePoint(ch.codePointAt(0) ?? 0)?.id ?? 0;
        if (glyph <= 0) {
          if (ch.trim() !== '')
            p.diag.add(
              'text.missing-glyph',
              'approximation',
              'Characters the face lacks are not drawn.'
            );
          continue;
        }
        // Text space: y up again inside the flipped page, at the size layout measured.
        p.push(
          setTextMatrix(size, 0, 0, -size, x, y),
          showText(PDFHexString.of(font.encode(glyph, ch)))
        );
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
      decorate();
    }
  }
  p.push(popGraphicsState());
}

/** Underline styles the bars draw exactly; the SVG paint draws every other style the same way. */
const EXACT_UNDERLINES = new Set(['none', 'sng', 'dbl']);

/**
 * Underline and strikethrough as filled bars, with the geometry of the SVG paint
 * (`core/src/output/paint-text.ts`, `decorationRects`): face metrics, trailing spaces trimmed,
 * the underline in its own fill (`a:uFill`) when the run has one, the strikethrough in the
 * run's fill. A gradient fill spans the text body's bounds, as the glyphs' does.
 */
function drawDecorations(
  run: GlyphRun,
  baseline: number,
  scope: FillScope,
  p: PageOps,
  bounds: () => { x: number; y: number; w: number; h: number }
): void {
  const { underline, strike } = run.props;
  if (underline === 'none' && strike === 'noStrike') return;
  if (!EXACT_UNDERLINES.has(underline))
    p.diag.add(
      'text.underline-style',
      'approximation',
      'Dotted, dashed, wavy, heavy, and word-only underlines are drawn as single lines.'
    );
  let end = run.text.length;
  while (end > 0 && run.text[end - 1] === ' ') end -= 1;
  const x0 = run.x[0];
  const x1 = run.x[end] ?? run.x[run.x.length - 1];
  if (!(x1 > x0)) return;
  const size = run.sizePt * EMU_PER_PT;
  const m = run.face.metrics;
  const thickness = Math.max(m.underlineThickness * size, size / 20);
  const y = baseline - run.rise;
  const bars = (fillValue: ResolvedFill, tops: readonly number[]): void => {
    const own = ownFill(fillValue, scope).fill;
    const trace = (): void => {
      for (const top of tops)
        p.push(
          moveTo(x0, top),
          lineTo(x1, top),
          lineTo(x1, top + thickness),
          lineTo(x0, top + thickness),
          closePath()
        );
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
  if (underline !== 'none') {
    const top = y - m.underlinePosition * size;
    bars(
      run.props.underlineFill ?? run.props.fill,
      underline === 'dbl' ? [top, top + thickness * 2] : [top]
    );
  }
  if (strike !== 'noStrike') {
    const mid = y - size * 0.3;
    bars(
      run.props.fill,
      strike === 'dblStrike' ? [mid - thickness / 2, mid - thickness * 2] : [mid - thickness / 2]
    );
  }
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
  for (const line of layout.lines) {
    for (const run of line.runs) {
      if (run.text.length === 0) continue;
      if (ownFill(run.props.fill, scope).fill.type === 'none') continue;
      const face = p.res.fonts.get(run.face.id) ?? fontFor(run.face.id, run.face.family, p);
      if (!face) continue;
      const k = (run.sizePt * EMU_PER_PT) / face.font.unitsPerEm;
      const y = line.baseline - run.rise;
      for (let i = 0; i < run.text.length; i += 1) {
        const code = run.text.charCodeAt(i);
        const pair = code >= 0xd800 && code <= 0xdbff && i + 1 < run.text.length;
        const cp = run.text.codePointAt(i) ?? 0;
        const x = run.x[i];
        if (pair) i += 1;
        const glyph = face.font.glyphForCodePoint(cp);
        if (!glyph || glyph.id <= 0) continue;
        // Font units, y up → EMU, y down, at the glyph's origin.
        const px = (gx: number): number => x + gx * k;
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
        // A path object ends at its paint operator; glyph outlines fill nonzero.
        if (glyph.path.commands.length > 0) p.push(strokeToo ? fillAndStroke() : fill());
      }
    }
  }
  p.push(popGraphicsState());
}
