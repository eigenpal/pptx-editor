// Slide layout: resolved elements → positioned records. Geometry is evaluated at each
// element's size, transforms are composed through groups, and text bodies are laid out in
// their geometry's text rectangle. Paint reads these records and never re-derives geometry.

import type {
  ResolvedElement,
  ResolvedGraphicFrame,
  ResolvedGroup,
  ResolvedPicture,
  ResolvedLine,
  ResolvedShape,
  ResolvedSlide,
} from '../cascade/index.ts';
import { slideApproximations } from './approximations.ts';
import type { EvaluatedGeometry } from './geometry/geometry.ts';
import { IDENTITY, type Matrix } from './matrix.ts';
import { layoutElement } from './item-layout.ts';
import type { TableLayout } from './table-layout.ts';
import type { TextLayout } from './text/text-layout.ts';
import type { TextMeasurer } from './text/measurer.ts';

export * from './matrix.ts';

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
        /** Outline used for paint when native dash expansion exceeds the resource limit. */
        readonly paintLine?: ResolvedLine;
        readonly text?: TextLayout;
        /** Local text frame → slide. Text is never mirrored by a flip. */
        readonly textMatrix?: Matrix;
      }
    | {
        readonly kind: 'picture';
        readonly element: ResolvedPicture;
        readonly matrix: Matrix;
        readonly geometry: EvaluatedGeometry;
        /** Outline used for paint when native dash expansion exceeds the resource limit. */
        readonly paintLine?: ResolvedLine;
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
        /**
         * The laid-out chart, for a chart frame that an installed chart renderer draws: shapes
         * and labels in the frame's box, in paint order. They carry ids under the frame's id
         * and are never hit targets: the frame is the object.
         */
        readonly chart?: readonly LayoutItem[];
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

/**
 * Laid-out elements by their resolved element, with the measurer they were measured with. The
 * cascade keeps a resolved element's object while nothing it reads changed
 * (`resolveSlide`), and a top-level element lays out from it and the measurer only: the same
 * pair gives the same item, so an unchanged shape keeps its item and its paint.
 */
const laidByElement = new WeakMap<
  ResolvedElement,
  {
    readonly measurer: TextMeasurer;
    readonly item: LayoutItem | null;
    readonly diagnostics: readonly string[];
  }
>();

/** Lay out a resolved slide. */
export function layoutSlide(slide: ResolvedSlide, measurer: TextMeasurer): SlideLayout {
  const diagnostics = new Set<string>();
  const items: LayoutItem[] = [];
  for (const el of slide.elements) {
    let hit = laidByElement.get(el);
    if (!hit || hit.measurer !== measurer) {
      const own = new Set<string>();
      const item = layoutElement(el, IDENTITY, null, measurer, own);
      hit = { measurer, item, diagnostics: [...own] };
      laidByElement.set(el, hit);
    }
    for (const d of hit.diagnostics) diagnostics.add(d);
    if (hit.item) items.push(hit.item);
  }
  for (const id of slideApproximations(items, slide.background))
    diagnostics.add(`approximation: ${id}`);
  return { slide, items, diagnostics: [...diagnostics] };
}
