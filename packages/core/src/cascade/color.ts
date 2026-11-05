// DrawingML color resolution over the canonical tree. The color models and the transform math
// are format-neutral and live in `@eigenpal/ooxml-core` (`drawingml-color.ts`). This module
// keeps the PresentationML glue: the theme color scheme read through the effective color map
// (`p:clrMap` with any `p:clrMapOvr`, ECMA-376 Part 1, 19.3.1.6 and 19.3.1.7) and the style
// reference color that `phClr` stands for.

import {
  applyColorTransforms,
  isDrawingColorModel,
  resolveDrawingColor,
  type DrawingColorTransform,
  type Rgba,
} from '@eigenpal/ooxml-core';
import { attr, elements, NS, type Element } from '../store/index.ts';

export {
  BLACK,
  linearToSrgb,
  parseHex,
  srgbToLinear,
  toCss,
  toHex,
  type Rgba,
} from '@eigenpal/ooxml-core';

/** What a color reference resolves against. */
export interface ColorContext {
  /** The theme color scheme: dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink. */
  readonly scheme: ReadonlyMap<string, Rgba>;
  /** The effective color map: bg1 → lt1, tx1 → dk1, … (master clrMap with any override). */
  readonly colorMap: ReadonlyMap<string, string>;
  /** The style-reference color substituted for `phClr`, when inside a style reference. */
  readonly placeholderColor?: Rgba;
}

/** Whether an element is one of the six DrawingML color models. */
export function isColorElement(el: Element): boolean {
  return el.ns === NS.a && isDrawingColorModel(el.local);
}

/** The first color-model child of a fill or other color container. */
export function colorChild(el: Element | undefined): Element | undefined {
  if (!el) return undefined;
  for (const c of elements(el)) if (isColorElement(c)) return c;
  return undefined;
}

/** Resolve a scheme color name through the color map to the theme scheme. */
export function schemeColor(name: string, ctx: ColorContext): Rgba | undefined {
  if (name === 'phClr') return ctx.placeholderColor;
  const mapped = ctx.colorMap.get(name) ?? name;
  return ctx.scheme.get(mapped) ?? ctx.scheme.get(name);
}

/** The transform children of a color element in the DrawingML main namespace, in order. */
function transformsOf(children: readonly Element[]): DrawingColorTransform[] {
  const out: DrawingColorTransform[] = [];
  for (const t of children) if (t.ns === NS.a) out.push({ name: t.local, val: attr(t, 'val') });
  return out;
}

/**
 * Apply the transforms listed as children of a color element, in document order, as the spec
 * requires (each transform sees the result of the previous one).
 */
export function applyTransforms(base: Rgba, transforms: readonly Element[]): Rgba {
  return applyColorTransforms(base, transformsOf(transforms));
}

/** Resolve a color element (`a:srgbClr` …) to RGBA, or undefined if it cannot resolve. */
export function resolveColor(el: Element | undefined, ctx: ColorContext): Rgba | undefined {
  if (!el || !isColorElement(el)) return undefined;
  return resolveDrawingColor(
    {
      model: el.local,
      readAttribute: (name) => attr(el, name),
      transforms: transformsOf(elements(el)),
    },
    (name) => schemeColor(name, ctx)
  );
}
