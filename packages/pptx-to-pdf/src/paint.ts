/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Paint one laid-out slide into a PDF page. Like DOCX Editor's `docx-to-pdf`
// (`packages/docx-to-pdf/src/paint.ts`), this reads the engine's layout, the same
// `layoutSlide` output the editor's SVG paint reads; it is not a second layout.
//
// The page content works in slide EMU with y down: one page matrix maps EMU to PDF points and
// flips y, so geometry paths and layout matrices apply unchanged. Text draws one character at a
// time at the x position layout measured, so line breaks and positions are the engine's.

import {
  closePath,
  concatTransformationMatrix,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setLineWidth,
  stroke,
} from 'pdf-lib';
import {
  hasCharacters,
  type LayoutItem,
  type PathCommand,
  type SlideLayout,
} from '@pptx-editor.dev/core/layout';
import { drawEffectsBehind, drawInnerShadow, effectDiagnostics } from './paint-effects.ts';
import {
  canDrawPicture,
  drawPicture,
  ownFill,
  paintFill,
  strokeStyle,
  type FillScope,
} from './paint-fill.ts';
import { drawTable } from './paint-table.ts';
import { drawText } from './paint-text.ts';
import {
  clip,
  clipEvenOdd,
  clipNonZero,
  matrixScale,
  EMU_PER_PT,
  endPath,
  PageOps,
  tracePath,
  traceRect,
  type PaintResources,
  type SlideDiagnostics,
} from './page-ops.ts';

/** One pixel at 96 DPI (9525 EMU, 0.75 pt): less than this past the box is rounding in the presets. */
const PAST_BOX_EMU = 9525;

/** The point at `s` (0–1) along a curve that starts at (x0, y0). */
function curvePoint(
  c: Extract<PathCommand, { op: 'C' | 'Q' }>,
  x0: number,
  y0: number,
  s: number
): [number, number] {
  const u = 1 - s;
  if (c.op === 'Q')
    return [
      u * u * x0 + 2 * u * s * c.x1 + s * s * c.x,
      u * u * y0 + 2 * u * s * c.y1 + s * s * c.y,
    ];
  return [
    u * u * u * x0 + 3 * u * u * s * c.x1 + 3 * u * s * s * c.x2 + s * s * s * c.x,
    u * u * u * y0 + 3 * u * u * s * c.y1 + 3 * u * s * s * c.y2 + s * s * s * c.y,
  ];
}

/** Whether a path reaches past the `w` × `h` box by more than one point, curves included. */
function pastBox(commands: readonly PathCommand[], w: number, h: number): boolean {
  const t = PAST_BOX_EMU;
  const out = (x: number, y: number): boolean => x < -t || y < -t || x > w + t || y > h + t;
  let [x0, y0] = [0, 0];
  for (const c of commands) {
    if (c.op === 'Z') continue;
    if (out(c.x, c.y)) return true;
    // A curve can bulge past the box between its ends (`heart`): test points along it.
    if (c.op === 'C' || c.op === 'Q')
      for (let i = 1; i < 8; i += 1) if (out(...curvePoint(c, x0, y0, i / 8))) return true;
    [x0, y0] = [c.x, c.y];
  }
  return false;
}

/** A path with no `close`: the only kind that gets line ends (as the SVG paint's `isOpenPath`). */
function isOpenPath(commands: readonly PathCommand[]): boolean {
  return commands.length >= 2 && !commands.some((c) => c.op === 'Z');
}

type Drawable = Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>;

/**
 * Clip to the outside of a picture's path and double the line width, so the stroke lies
 * outside the picture, as the
 * editor's paint does (`output/paint.ts`, `outsideOutline`).
 */
function outsideOf(
  commands: readonly PathCommand[],
  width: number,
  cx: number,
  cy: number,
  p: PageOps
): void {
  const pad = 2 * width + Math.max(cx, cy);
  p.push(
    moveTo(-pad, -pad),
    lineTo(cx + pad, -pad),
    lineTo(cx + pad, cy + pad),
    lineTo(-pad, cy + pad),
    closePath()
  );
  tracePath(commands, p);
  p.push(clipEvenOdd(), endPath(), setLineWidth(2 * width));
}

/** A shape's or picture's geometry: fills, the picture clipped to the geometry, outlines. */
async function drawGeometry(
  item: Drawable,
  p: PageOps,
  scope: FillScope,
  groupScale: number
): Promise<void> {
  const el = item.element;
  const { cx: w, cy: h } = el.xfrm;
  const paths = item.geometry.paths;
  if (item.kind === 'picture') {
    const own = ownFill(el.fill, scope);
    if (own.fill.type === 'picture') {
      // The picture shows inside its geometry (nonzero, as the SVG clip path).
      p.push(pushGraphicsState());
      if (paths.length) for (const path of paths) tracePath(path.commands, p);
      else traceRect(0, 0, w, h, p);
      p.push(clipNonZero(), endPath());
      await drawPicture(own.sourcePart, own.fill, w, h, p);
      p.push(popGraphicsState());
    }
  }
  // Picture fills and rectangular path gradients paint the shape's box only.
  const boxed = ownFill(el.fill, scope).fill;
  // Only a shape's fill: the editor draws a `p:pic` the same way, its image clipped to the
  // geometry, so a picture's geometry past its box differs from nothing.
  const boxOnly =
    item.kind === 'shape' &&
    (boxed.type === 'picture' || (boxed.type === 'path' && boxed.path !== 'circle'));
  // Only an outline that draws: a picture with `a:ln/a:noFill` has none to split.
  if (item.kind === 'picture' && el.line && el.line.fill.type !== 'none' && paths.length > 1)
    p.diag.add(
      'picture.outline',
      'approximation',
      'The border of a picture cropped to a shape with several paths is drawn path by path.'
    );
  // Path by path, fill then outline, in the SVG paint's order.
  for (const path of paths) {
    if (boxOnly && path.fill !== 'none' && pastBox(path.commands, w, h))
      p.diag.add(
        'fill.outside-box',
        'approximation',
        "Picture fills and rectangular path gradients paint only the shape's box."
      );
    if (item.kind === 'shape' && path.fill !== 'none')
      await paintFill(
        el.fill,
        w,
        h,
        () => tracePath(path.commands, p),
        p,
        scope,
        'shapes',
        path.fill
      );
    // A picture's outline follows all of its geometry, as in the SVG paint.
    if (!el.line || (!path.stroke && item.kind !== 'picture')) continue;
    p.push(pushGraphicsState());
    const color = strokeStyle(el.line, p, scope, groupScale);
    if (color) {
      // Line ends show on open paths only, as in the SVG paint.
      if (item.kind === 'shape' && (el.line.head || el.line.tail) && isOpenPath(path.commands))
        p.diag.add('line.ends', 'unsupported', 'Arrowheads are not drawn yet.');
      p.alpha(1, color.a);
      // A picture's border lies outside the picture.
      if (item.kind === 'picture')
        outsideOf(
          path.commands,
          (el.line.width > 0 ? el.line.width : 9525) / (groupScale || 1),
          w,
          h,
          p
        );
      tracePath(path.commands, p);
      p.push(stroke());
    }
    p.push(popGraphicsState());
  }
}

async function drawItem(
  item: LayoutItem,
  p: PageOps,
  scope: FillScope,
  groupScale: number
): Promise<void> {
  if (item.element.hidden) return;
  await p.step();
  const local: FillScope = { ...scope, sourcePart: item.element.partName };
  if (item.kind === 'group') {
    const inner: FillScope = {
      ...local,
      groupFill: item.element.fill,
      groupSourcePart: item.element.partName,
    };
    for (const child of item.children) await drawItem(child, p, inner, matrixScale(item.matrix));
    return;
  }
  if (item.kind === 'frame') {
    // An OLE object shows its picture, as in the SVG paint.
    if (item.poster) return drawItem(item.poster, p, local, groupScale);
    if (!item.table) {
      const kind = item.element.frameKind;
      p.diag.add(`frame.${kind}`, 'unsupported', `${kind} frames are not drawn yet.`);
      return;
    }
    p.push(pushGraphicsState());
    p.matrix(item.matrix);
    await drawTable(item.table, item.element.table?.background ?? { type: 'none' }, p, local);
    p.push(popGraphicsState());
    return;
  }
  // A picture fill whose image is not drawn (a `p:pic`, or a shape's picture fill) casts no
  // shadow and has no inside to shade.
  const own = ownFill(item.element.fill, local);
  const shown = own.fill.type !== 'picture' || (await canDrawPicture(own.sourcePart, own.fill, p));
  if (shown) effectDiagnostics(item, p, local);
  // Shadows and glow paint under the content, as in the SVG paint.
  if (shown) await drawEffectsBehind(item, p, local);
  p.push(pushGraphicsState());
  p.matrix(item.matrix);
  await drawGeometry(item, p, local, groupScale);
  p.push(popGraphicsState());
  // Over the geometry and under the text, in slide space.
  if (shown) await drawInnerShadow(item, p, local);
  // An empty body's caret line draws nothing: no operators for it.
  if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
    p.push(pushGraphicsState());
    p.matrix(item.textMatrix);
    await drawText(item.text, p, local);
    p.push(popGraphicsState());
  }
}

/** Map layout's fidelity notes to PDF diagnostics, with fixed codes. */
function layoutDiagnostics(layout: SlideLayout, diag: SlideDiagnostics): void {
  const after = (note: string, prefix: string): string => clip(note.slice(prefix.length));
  for (const note of layout.diagnostics) {
    if (note.startsWith('approximation: '))
      diag.add(
        'layout.approximation',
        'information',
        `Engine approximation: ${after(note, 'approximation: ')}.`
      );
    else if (note.startsWith('font substituted: '))
      diag.add(
        'font.substituted',
        'information',
        `Font substituted: ${after(note, 'font substituted: ')}.`
      );
    else if (note.startsWith('unknown preset geometry: '))
      diag.add(
        'geometry.unknown-preset',
        'approximation',
        `Unknown preset geometry drawn as a rectangle: ${after(note, 'unknown preset geometry: ')}.`
      );
    else if (note.startsWith('effect not rendered: '))
      diag.add(
        'effect.not-rendered',
        'unsupported',
        `Effect not rendered: ${after(note, 'effect not rendered: ')}.`
      );
  }
}

/** Paint one slide layout onto a new page. */
export async function paintSlide(
  layout: SlideLayout,
  res: PaintResources,
  diag: SlideDiagnostics
): Promise<void> {
  const { width, height } = layout.slide;
  const page = res.doc.addPage([width / EMU_PER_PT, height / EMU_PER_PT]);
  const p = new PageOps(page, res, diag);
  layoutDiagnostics(layout, diag);
  // EMU, y down → points, y up.
  p.push(
    pushGraphicsState(),
    concatTransformationMatrix(1 / EMU_PER_PT, 0, 0, -1 / EMU_PER_PT, 0, height / EMU_PER_PT)
  );
  const scope: FillScope = { sourcePart: layout.slide.backgroundPart };
  // A page is white, as the SVG paint's background without a fill.
  await paintFill(
    layout.slide.background,
    width,
    height,
    () => traceRect(0, 0, width, height, p),
    p,
    scope,
    'the background'
  );
  for (const item of layout.items) await drawItem(item, p, scope, 1);
  p.push(popGraphicsState());
  p.flush();
}
