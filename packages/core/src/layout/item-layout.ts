// One element → its positioned record: the group child placement, the geometry at the
// element's size, the text in the geometry's text rectangle, and the group's children.
// `relayoutItem` lays out one item again at a new box, for the editor's drag preview; it is
// not on the `layout` subpath. A resized shape takes the autofit the release writes
// (`fitShape`), and an object's picture takes the frame's new box.

import type {
  ResolvedElement,
  ResolvedGraphicFrame,
  ResolvedShape,
  Xfrm,
} from '../cascade/index.ts';
import { fitShape } from './autofit.ts';
import { drawingRenderers } from './drawing-renderers.ts';
import { evaluateGeometry } from './geometry/geometry.ts';
import { boxMatrix, multiply, type Matrix } from './matrix.ts';
import type { ItemPlacement, LayoutItem, SourceMap } from './slide-layout.ts';
import { layoutTable } from './table-layout.ts';
import { layoutTextOn, type TextLayout } from './text/text-layout.ts';
import type { LineBasis } from './text/line-metrics.ts';
import type { TextMeasurer } from './text/measurer.ts';

/** Maps a group's child coordinates (`chOff`/`chExt`) into the group's local box. */
type ChildMap = SourceMap;

/**
 * Place a group child in the group's local box. The group scales a child's BOX with the
 * group, never its fonts or line widths, so the child is laid out at its scaled size and the
 * group contributes only translation, rotation, and flips through the matrix.
 */
function placeInGroup<T extends ResolvedElement>(el: T, map: ChildMap | null): T {
  if (!map) return el;
  const place = (x: Xfrm): Xfrm => ({
    ...x,
    x: (x.x - map.ox) * map.sx,
    y: (x.y - map.oy) * map.sy,
    cx: x.cx * map.sx,
    cy: x.cy * map.sy,
  });
  const textXfrm = el.kind === 'shape' ? el.textXfrm : undefined;
  return { ...el, xfrm: place(el.xfrm), ...(textXfrm ? { textXfrm: place(textXfrm) } : {}) };
}

/** Whether a laid-out diagram item falls outside what the shape renderer draws exactly. */
function unsupportedInDiagram(item: LayoutItem): boolean {
  if (item.kind === 'group') return item.children.some(unsupportedInDiagram);
  return item.kind === 'shape' && item.geometry.unknownPreset !== undefined;
}

/**
 * Lay out a diagram's saved drawing in the frame's box. The drawing's shapes are in the frame's
 * file space, so a frame scaled by a group scales them the same way. Null when one shape uses
 * a preset geometry that the renderer does not know: a diagram renders whole or not at all.
 */
function layoutDiagram(
  source: ResolvedGraphicFrame,
  el: ResolvedGraphicFrame,
  matrix: Matrix,
  measurer: TextMeasurer,
  diagnostics: Set<string>
): LayoutItem[] | null {
  if (!source.drawing) return null;
  const map: ChildMap = {
    ox: 0,
    oy: 0,
    sx: source.xfrm.cx ? el.xfrm.cx / source.xfrm.cx : 1,
    sy: source.xfrm.cy ? el.xfrm.cy / source.xfrm.cy : 1,
  };
  const own = new Set<string>();
  const items: LayoutItem[] = [];
  for (const c of source.drawing) {
    const item = layoutElement(c, matrix, map, measurer, own);
    if (!item) continue;
    if (unsupportedInDiagram(item)) return null;
    items.push(item);
  }
  for (const d of own) diagnostics.add(d);
  return items;
}

/**
 * The text frame → slide matrix for a shape whose box maps to the slide by `shape`. Text stays
 * readable: the text rectangle moves with the flipped geometry, but the glyphs are
 * never mirrored. A matrix with a negative determinant (an odd number of flips, from the shape
 * and every enclosing group) is corrected by a horizontal mirror about the text frame's center.
 * The result: a net horizontal flip leaves text upright, and a net vertical flip turns it 180°.
 */
function uprightTextMatrix(
  shape: Matrix,
  frame: { readonly x: number; readonly width: number }
): Matrix {
  if (shape[0] * shape[3] - shape[1] * shape[2] >= 0) return shape;
  const cx = frame.x + frame.width / 2;
  return multiply(shape, [-1, 0, 0, 1, 2 * cx, 0]);
}

export function layoutElement(
  source: ResolvedElement,
  parent: Matrix,
  map: ChildMap | null,
  measurer: TextMeasurer,
  diagnostics: Set<string>,
  /** The single line that text in a shape's own geometry is spaced on (`LineBasis`). */
  basis: LineBasis = 'size'
): LayoutItem | null {
  if (source.hidden) return null;
  const el = placeInGroup(source, map);
  return layoutPlaced(el, source, parent, map, measurer, diagnostics, basis);
}

/**
 * Lay out one item again with a new box in its parent space (`xfrm`), with nothing written:
 * the editor's drag preview. Its children and text follow the new box. A shape whose size
 * changed takes its autofit at the new size, as the release re-fits it: `spAutoFit` sizes the
 * box to the text, and `normAutofit` scales the text to the box.
 */
export function relayoutItem(
  item: LayoutItem,
  xfrm: Xfrm,
  measurer: TextMeasurer
): LayoutItem | null {
  const old = item.element.xfrm;
  let el = { ...item.element, xfrm } as ResolvedElement;
  const resized = Math.abs(old.cx - xfrm.cx) > 1 || Math.abs(old.cy - xfrm.cy) > 1;
  if (el.kind === 'shape' && resized) el = fitted(el, measurer);
  // The poster has the frame's file box (`layoutElement` places it in the group again).
  if (el.kind === 'graphicFrame' && el.poster)
    el = { ...el, poster: { ...el.poster, xfrm: fileBox(xfrm, item.toSource) } };
  // A diagram's drawing is in the frame's old file box: it scales onto the new box.
  const source = { ...item.element, xfrm: fileBox(old, item.toSource) } as ResolvedElement;
  return layoutPlaced(el, source, item.parentMatrix, item.toSource, measurer, new Set());
}

/** A shape with the autofit its box gives (`fitShape`). */
function fitted(el: ResolvedShape, measurer: TextMeasurer): ResolvedShape {
  const fit = fitShape(el, measurer);
  if (fit?.kind === 'size') return { ...el, xfrm: { ...el.xfrm, cx: fit.cx, cy: fit.cy } };
  if (!fit || !el.text) return el;
  const autofit = {
    kind: 'normal' as const,
    fontScale: fit.scale,
    lineSpaceReduction: fit.reduction,
  };
  return { ...el, text: { ...el.text, body: { ...el.text.body, autofit } } };
}

/** A box in a group's local space as a box in its child space (the inverse of `placeInGroup`). */
function fileBox(x: Xfrm, map: ChildMap | null): Xfrm {
  if (!map) return x;
  const sx = map.sx || 1;
  const sy = map.sy || 1;
  return { ...x, x: x.x / sx + map.ox, y: x.y / sy + map.oy, cx: x.cx / sx, cy: x.cy / sy };
}

/**
 * Lay out an element whose box is in its parent space already (`placeInGroup`); `source` is the
 * element as the file places it.
 */
function layoutPlaced(
  el: ResolvedElement,
  source: ResolvedElement,
  parent: Matrix,
  map: ChildMap | null,
  measurer: TextMeasurer,
  diagnostics: Set<string>,
  basis: LineBasis = 'size'
): LayoutItem | null {
  const matrix = multiply(parent, boxMatrix(el.xfrm));
  const placement: ItemPlacement = { parentMatrix: parent, toSource: map };
  if (el.kind === 'shape' || el.kind === 'connector' || el.kind === 'picture')
    for (const name of el.effects?.unsupported ?? [])
      diagnostics.add(`effect not rendered: ${name}`);
  switch (el.kind) {
    case 'shape':
    case 'connector': {
      const geometry = evaluateGeometry(el.geometry, el.xfrm.cx, el.xfrm.cy);
      if (geometry.unknownPreset)
        diagnostics.add(`unknown preset geometry: ${geometry.unknownPreset}`);
      let text: TextLayout | undefined;
      let textMatrix: Matrix | undefined;
      const tx = el.text && el.text.paragraphs.length > 0 ? el.textXfrm : undefined;
      if (el.text && tx) {
        // A diagram shape's own text rectangle, in the parent's space. Its turn adds to the
        // shape's turn, and it never flips. A flipped frame or group keeps the glyphs upright,
        // as for any other shape. Diagram text spaces its lines on the face's line
        // (`LineBasis` in `line-metrics.ts`).
        const box = { x: 0, y: 0, width: tx.cx, height: tx.cy };
        text = layoutTextOn(el.text, box, measurer, 'face');
        for (const s of text.substitutions) diagnostics.add(`font substituted: ${s}`);
        textMatrix = uprightTextMatrix(
          multiply(parent, boxMatrix({ ...tx, rot: tx.rot + el.xfrm.rot }, false)),
          text.box
        );
      } else if (el.text && el.text.paragraphs.length > 0) {
        const r = geometry.textRect;
        text = layoutTextOn(
          el.text,
          { x: r.l, y: r.t, width: r.r - r.l, height: r.b - r.t },
          measurer,
          basis
        );
        for (const s of text.substitutions) diagnostics.add(`font substituted: ${s}`);
        textMatrix = uprightTextMatrix(matrix, text.box);
      }
      return {
        ...placement,
        kind: 'shape',
        element: el as ResolvedShape,
        matrix,
        geometry,
        ...(text ? { text, textMatrix } : {}),
      };
    }
    case 'picture':
      return {
        ...placement,
        kind: 'picture',
        element: el,
        matrix,
        geometry: evaluateGeometry(el.geometry, el.xfrm.cx, el.xfrm.cy),
      };
    case 'group': {
      const chExt = el.childExtent;
      const childMap: ChildMap = {
        ox: el.childOffset.x,
        oy: el.childOffset.y,
        sx: chExt.cx ? el.xfrm.cx / chExt.cx : 1,
        sy: chExt.cy ? el.xfrm.cy / chExt.cy : 1,
      };
      const children: LayoutItem[] = [];
      for (const c of el.children) {
        const item = layoutElement(c, matrix, childMap, measurer, diagnostics);
        if (item) children.push(item);
      }
      return { ...placement, kind: 'group', element: el, matrix, children };
    }
    case 'graphicFrame': {
      if (el.table) {
        const substituted = new Set<string>();
        const table = layoutTable(el.table, measurer, substituted);
        for (const s of substituted) diagnostics.add(`font substituted: ${s}`);
        return { ...placement, kind: 'frame', element: el, matrix, table };
      }
      // The poster has the frame's file box; it is placed in a group as the frame is.
      const poster = el.poster && layoutElement(el.poster, parent, map, measurer, diagnostics);
      if (poster?.kind === 'picture')
        return { ...placement, kind: 'frame', element: el, matrix, poster };
      const drawing = layoutDiagram(
        source as ResolvedGraphicFrame,
        el,
        matrix,
        measurer,
        diagnostics
      );
      if (drawing) return { ...placement, kind: 'frame', element: el, matrix, drawing };
      const chart = el.chart && layoutChart(el, matrix, measurer, diagnostics);
      if (chart) return { ...placement, kind: 'frame', element: el, matrix, chart };
      diagnostics.add(`graphic frame rendered as placeholder: ${el.frameKind}`);
      // A drawing that resolved but holds a shape layout cannot draw (an unknown preset).
      const element: ResolvedGraphicFrame = el.drawing
        ? { ...el, drawing: undefined, drawingFallback: 'unsupported' }
        : el;
      return { ...placement, kind: 'frame', element, matrix };
    }
  }
}

/**
 * Lay out a chart frame through the installed chart renderer, or null when none is installed or
 * it does not draw this chart. The renderer's elements sit in the frame's local box.
 */
function layoutChart(
  frame: ResolvedGraphicFrame,
  matrix: Matrix,
  measurer: TextMeasurer,
  diagnostics: Set<string>
): LayoutItem[] | null {
  const chart = frame.chart;
  if (!chart) return null;
  const render = drawingRenderers().chart;
  const elements = render?.({ frame, chart, measurer });
  if (!elements) return null;
  const items: LayoutItem[] = [];
  // Chart labels space their lines on the face's line, as diagram text does: each label line is
  // as tall as its face's ascent plus descent, with the baseline at the ascent (the chart
  // labels of the `charts` and `slide-structure` reference PDFs).
  for (const el of elements) {
    const item = layoutElement(el, matrix, null, measurer, diagnostics, 'face');
    if (item) items.push(item);
  }
  return items;
}
