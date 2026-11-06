// Slide layout: resolved elements → positioned records. Geometry is evaluated at each
// element's size, transforms are composed through groups, and text bodies are laid out in
// their geometry's text rectangle. Paint reads these records and never re-derives geometry.

import type {
  ResolvedElement,
  ResolvedGraphicFrame,
  ResolvedGroup,
  ResolvedPicture,
  ResolvedShape,
  ResolvedSlide,
  Xfrm,
} from '../cascade/index.ts';
import { slideApproximations } from './approximations.ts';
import { evaluateGeometry, type EvaluatedGeometry } from './geometry/geometry.ts';
import { layoutTable, type TableLayout } from './table-layout.ts';
import { layoutText, type TextLayout } from './text/text-layout.ts';
import type { TextMeasurer } from './text/measurer.ts';

/** An affine matrix `[a, b, c, d, e, f]` (SVG convention). EMU units. */
export type Matrix = readonly [number, number, number, number, number, number];

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `m × n`: apply `n` first, then `m`. */
export function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

/** Apply a matrix to a point. */
export function apply(m: Matrix, x: number, y: number): { x: number; y: number } {
  return { x: m[0] * x + m[2] * y + m[4], y: m[1] * x + m[3] * y + m[5] };
}

/** Invert an affine matrix. */
export function invert(m: Matrix): Matrix {
  const det = m[0] * m[3] - m[1] * m[2] || 1e-12;
  return [
    m[3] / det,
    -m[1] / det,
    -m[2] / det,
    m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det,
    (m[1] * m[4] - m[0] * m[5]) / det,
  ];
}

const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const scale = (sx: number, sy: number): Matrix => [sx, 0, 0, sy, 0, 0];
function rotate(deg: number): Matrix {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [c, s, -s, c, 0, 0];
}

/** Box-local → parent: rotate and flip around the box center, then place at the offset. */
export function boxMatrix(x: Xfrm, withFlip = true, extraRotation = 0): Matrix {
  const cx = x.cx / 2;
  const cy = x.cy / 2;
  let m = translate(x.x + cx, x.y + cy);
  m = multiply(m, rotate(x.rot + extraRotation));
  if (withFlip) m = multiply(m, scale(x.flipH ? -1 : 1, x.flipV ? -1 : 1));
  return multiply(m, translate(-cx, -cy));
}

/**
 * Maps a group child's laid-out box back to the file: `x_file = x / sx + ox`,
 * `cx_file = cx / sx`. Null for top-level elements, whose boxes are the file's.
 */
export interface SourceMap {
  readonly ox: number;
  readonly oy: number;
  readonly sx: number;
  readonly sy: number;
}

/** Where an item sits: its parent space → slide, and its box → the file's coordinates. */
export interface ItemPlacement {
  readonly parentMatrix: Matrix;
  readonly toSource: SourceMap | null;
}

/** One positioned element. `matrix` maps the element's local box to slide EMU. */
export type LayoutItem = ItemPlacement &
  (
    | {
        readonly kind: 'shape';
        readonly element: ResolvedShape;
        readonly matrix: Matrix;
        readonly geometry: EvaluatedGeometry;
        readonly text?: TextLayout;
        /** Local text frame → slide. Text is never mirrored by a flip. */
        readonly textMatrix?: Matrix;
      }
    | {
        readonly kind: 'picture';
        readonly element: ResolvedPicture;
        readonly matrix: Matrix;
        readonly geometry: EvaluatedGeometry;
      }
    | {
        readonly kind: 'group';
        readonly element: ResolvedGroup;
        readonly matrix: Matrix;
        readonly children: readonly LayoutItem[];
      }
    | {
        readonly kind: 'frame';
        readonly element: ResolvedGraphicFrame;
        readonly matrix: Matrix;
        /** The laid-out table, for a table frame. */
        readonly table?: TableLayout;
        /**
         * The laid-out OLE picture (`element.poster`), in the frame's box. A frame with no
         * table and no poster draws a labeled placeholder.
         */
        readonly poster?: Extract<LayoutItem, { kind: 'picture' }>;
        /**
         * The laid-out saved drawing of a diagram (`element.drawing`): read-only shapes and
         * groups placed in the frame's box. A frame with none draws a labeled placeholder.
         */
        readonly drawing?: readonly LayoutItem[];
      }
  );

/** A laid-out slide. */
export interface SlideLayout {
  readonly slide: ResolvedSlide;
  readonly items: readonly LayoutItem[];
  /**
   * For the fidelity report: unknown presets, substituted fonts, unrendered effects and frames,
   * and `approximation: <id>` for each known approximation the slide uses.
   */
  readonly diagnostics: readonly string[];
}

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

function layoutElement(
  source: ResolvedElement,
  parent: Matrix,
  map: ChildMap | null,
  measurer: TextMeasurer,
  diagnostics: Set<string>
): LayoutItem | null {
  if (source.hidden) return null;
  const el = placeInGroup(source, map);
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
        // as for any other shape.
        text = layoutText(el.text, { x: 0, y: 0, width: tx.cx, height: tx.cy }, measurer);
        for (const s of text.substitutions) diagnostics.add(`font substituted: ${s}`);
        textMatrix = uprightTextMatrix(
          multiply(parent, boxMatrix({ ...tx, rot: tx.rot + el.xfrm.rot }, false)),
          text.box
        );
      } else if (el.text && el.text.paragraphs.length > 0) {
        const r = geometry.textRect;
        text = layoutText(
          el.text,
          { x: r.l, y: r.t, width: r.r - r.l, height: r.b - r.t },
          measurer
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
      diagnostics.add(`graphic frame rendered as placeholder: ${el.frameKind}`);
      return { ...placement, kind: 'frame', element: el, matrix };
    }
  }
}

/** Lay out a resolved slide. */
export function layoutSlide(slide: ResolvedSlide, measurer: TextMeasurer): SlideLayout {
  const diagnostics = new Set<string>();
  const items: LayoutItem[] = [];
  for (const el of slide.elements) {
    const item = layoutElement(el, IDENTITY, null, measurer, diagnostics);
    if (item) items.push(item);
  }
  for (const id of slideApproximations(items, slide.background))
    diagnostics.add(`approximation: ${id}`);
  return { slide, items, diagnostics: [...diagnostics] };
}
