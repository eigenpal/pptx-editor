// Drawing renderers that a capability package installs, after DOCX Editor's
// `packages/core/src/store/package/drawing-renderers.ts`.
//
// The engine paints shapes, pictures, fills, lines, text, and outer shadows on its own. A shape
// or picture paints without the effects that no renderer paints: glow, inner shadow,
// reflection, and soft edge come from the `effects` renderer (`@pptx-editor.dev/pro/drawing`).
// A chart keeps its labeled placeholder until a `chart` renderer is installed. Other capability
// packages add their own keys to `DrawingRenderers`.
//
// The boundary is DOM-free: an effects renderer returns SVG as plain data (`EffectNode`), and
// the SVG paint (`output/paint.ts`) builds the elements. A chart renderer returns resolved
// shapes in the frame's box; `layoutSlide` lays them out with the same geometry and text
// pipeline as every other shape, and the SVG paint and the PDF writer draw those records. Installation is process-wide, because every
// editor and export in the process paints with the same engine; install renderers before any
// document opens. Looking up a renderer is one property read.

import type {
  ResolvedChartPart,
  ResolvedEffects,
  ResolvedElement,
  ResolvedGraphicFrame,
} from '../cascade/index.ts';
import type { Matrix } from './slide-layout.ts';
import type { TextMeasurer } from './text/measurer.ts';

/**
 * One SVG element as plain data: the SVG paint creates it with `createElementNS` and
 * `setAttribute`, so no markup is ever parsed.
 *
 * @internal
 */
export interface EffectNode {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly children?: readonly EffectNode[];
}

/** Axis-aligned bounds in slide EMU. @internal */
export interface EffectBounds {
  readonly x: number;
  readonly y: number;
  readonly r: number;
  readonly b: number;
}

/** Unique ids within one painted slide. @internal */
export interface EffectIds {
  id(kind: string): string;
}

/** A filter definition and the id of the `<filter>` that the paint applies. @internal */
export interface EffectFilterPaint {
  /** Definitions for `<defs>`; one of them is the `<filter>` with id `filter`. */
  readonly defs: readonly EffectNode[];
  readonly filter: string;
}

/** What a copy of a shape's content needs: the painted group and its geometry. @internal */
export interface EffectCopySource {
  /** The id of the painted content group, for `<use href="#id">`. */
  readonly contentId: string;
  /** The shape's local box to the slide. */
  readonly matrix: Matrix;
  readonly width: number;
  readonly height: number;
  /** Slide-space corners of everything the content paints (box and text). */
  readonly corners: readonly { x: number; y: number }[];
}

/** Definitions and one node that the paint puts under the content. @internal */
export interface EffectCopyPaint {
  readonly defs: readonly EffectNode[];
  readonly node: EffectNode;
}

/**
 * Glow, inner shadow, reflection, and soft edge (ECMA-376 Part 1, 20.1.8). The outer shadow is
 * the engine's.
 *
 * @internal
 */
export interface DrawingEffectsRenderer {
  /** Inner shadow and soft edge, as a filter over the geometry. `bounds` hold the geometry. */
  readonly body: (
    effects: ResolvedEffects,
    bounds: EffectBounds,
    ids: EffectIds
  ) => EffectFilterPaint | null;
  /** Glow, as a filter over the content. `bounds` hold the content. */
  readonly glow: (
    effects: ResolvedEffects,
    bounds: EffectBounds,
    ids: EffectIds
  ) => EffectFilterPaint | null;
  /** The reflection, painted under the content. */
  readonly reflection: (
    effects: ResolvedEffects,
    source: EffectCopySource,
    ids: EffectIds
  ) => EffectCopyPaint | null;
}

/** What a chart renderer reads for one chart frame. @internal */
export interface ChartRenderInput {
  /** The graphic frame, at its laid-out size (`xfrm.cx`, `xfrm.cy`). */
  readonly frame: ResolvedGraphicFrame;
  /** The chart part, with the colors and theme its values resolve against. */
  readonly chart: ResolvedChartPart;
  /** The measurer that lays out every label, so sizes match the drawn text. */
  readonly measurer: TextMeasurer;
}

/**
 * A chart renderer: the shapes and labels of a chart as resolved elements in the frame's local
 * box (EMU, origin at the frame's top-left corner), in paint order. Null when the chart is
 * outside what the renderer draws; the frame then keeps its placeholder.
 *
 * @internal
 */
export type ChartRenderer = (input: ChartRenderInput) => readonly ResolvedElement[] | null;

/**
 * Renderers a capability package installs.
 *
 * @internal
 */
export interface DrawingRenderers {
  /** Glow, inner shadow, reflection, and soft edge. */
  readonly effects?: DrawingEffectsRenderer;
  /** Charts drawn from the values cached in their chart parts (ECMA-376 Part 1, 21.2). */
  readonly chart?: ChartRenderer;
}

let installed: DrawingRenderers = Object.freeze({});
/** The renderers as their package passed them, so installing the same ones again is silent. */
const passed = new WeakSet<object>();
let consulted = false;
let warnedLate = false;

/** `render`, returning `fallback` instead of throwing, so one bad shape paints without effects. */
function guard<A extends unknown[], R>(render: (...args: A) => R, fallback: R): (...args: A) => R {
  return (...args) => {
    try {
      return Reflect.apply(render, undefined, args);
    } catch {
      return fallback;
    }
  };
}

function guardedRenderers(renderers: Partial<DrawingRenderers>): Partial<DrawingRenderers> {
  const { effects, chart } = renderers;
  // Spread first, so a key named with `undefined` clears the installed renderer.
  return {
    ...renderers,
    ...(effects
      ? {
          effects: Object.freeze({
            body: guard(effects.body, null),
            glow: guard(effects.glow, null),
            reflection: guard(effects.reflection, null),
          }),
        }
      : {}),
    // A chart whose renderer fails keeps its placeholder.
    ...(chart ? { chart: guard(chart, null) } : {}),
  };
}

/**
 * Install drawing renderers for every editor and export in this process. A later install
 * replaces the renderers it names and keeps the others; a key set to `undefined` removes its
 * renderer. Install before any document opens: an install after the first paint logs a
 * console warning once, and slides already painted keep their paint until they paint again.
 *
 * @internal
 */
export function installDrawingRenderers(renderers: Partial<DrawingRenderers>): void {
  const added: object[] = Object.values(renderers).filter(
    (renderer: unknown): renderer is object =>
      (typeof renderer === 'object' && renderer !== null) || typeof renderer === 'function'
  );
  const fresh = added.filter((renderer) => !passed.has(renderer));
  for (const renderer of fresh) passed.add(renderer);
  if (fresh.length > 0 && consulted && !warnedLate) {
    warnedLate = true;
    console.warn(
      '[pptx-editor] Drawing renderers were installed after a slide painted. Slides painted ' +
        'before the install keep their paint until they paint again. Install them before the ' +
        'first document opens.'
    );
  }
  installed = Object.freeze({ ...installed, ...guardedRenderers(renderers) });
  version += 1;
  for (const listener of [...listeners]) {
    try {
      listener();
    } catch {
      // One failing listener does not stop the others or the install.
    }
  }
}

let version = 0;
const listeners = new Set<() => void>();

/**
 * A number that each `installDrawingRenderers` call increases. An adapter reads it with
 * {@link subscribeDrawingRenderers} (for example through `useSyncExternalStore`), so the
 * choices it shows refresh after an install. @public
 */
export function drawingRenderersVersion(): number {
  return version;
}

/**
 * Call `listener` after each `installDrawingRenderers` call. Returns the function that stops the
 * calls. Like {@link drawingCapabilities}, a subscription does not count as a use, so it logs no
 * warning. @public
 */
export function subscribeDrawingRenderers(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** The live {@link subscribeDrawingRenderers} subscriptions. Tests check cleanup with it. @internal */
export function drawingRenderersListenerCount(): number {
  return listeners.size;
}

/**
 * What the installed drawing renderers paint in this process.
 *
 * @public
 */
export interface DrawingCapabilities {
  /** Glow, inner shadow, reflection, and soft edge paint (`@pptx-editor.dev/pro/drawing`). */
  readonly effects: boolean;
}

const WITHOUT_EFFECTS: DrawingCapabilities = Object.freeze({ effects: false });
const WITH_EFFECTS: DrawingCapabilities = Object.freeze({ effects: true });

/**
 * What the installed drawing renderers paint, read at the time of the call. The chrome reads it
 * when it renders, to offer only the controls whose result paints. Unlike a paint, this query
 * does not count as a use, so an install after it logs no warning.
 *
 * @public
 */
export function drawingCapabilities(): DrawingCapabilities {
  return installed.effects ? WITH_EFFECTS : WITHOUT_EFFECTS;
}

/** The installed renderers. @internal */
export function drawingRenderers(): DrawingRenderers {
  consulted = true;
  return installed;
}
