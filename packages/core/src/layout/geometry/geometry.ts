// DrawingML geometry over the canonical tree. The guide formulas, the path commands, the text
// rectangle, the preset table, and the `a:custGeom` reader are format-neutral and live in
// `@eigenpal/ooxml-core` (`drawingml-geometry.ts`). This module keeps the glue: the cascade's
// `GeometrySpec`, the canonical tree read through `DrawingGeometryElement`, and the cache of
// custom definitions for each `a:custGeom` element.

import {
  customGeometryDefinition,
  evaluateGeometry as evaluateGeometryDefinition,
  evaluatePresetGeometry,
  geometryConnectionSites,
  presetConnectionSites,
  type ConnectionSite,
  type DrawingGeometryElement,
  type EvaluatedGeometry,
  type PresetShapeDefinition,
} from '@eigenpal/ooxml-core';
import { attr, elements, type Element } from '../../store/index.ts';
import type { GeometrySpec } from '../../cascade/index.ts';

export {
  adjustForHandle,
  evaluateFormula,
  presetAdjustments,
  presetHandles,
  presetNames,
  type ConnectionSite,
  type EvaluatedGeometry,
  type PresetAdjustment,
  type PresetHandle,
  type GeometryPath,
  type PathCommand,
} from '@eigenpal/ooxml-core';

/** A canonical tree element as the neutral geometry reader sees it. */
function geometryElement(el: Element): DrawingGeometryElement {
  return {
    namespaceUri: el.ns,
    localName: el.local,
    attribute: (name) => attr(el, name),
    childElements: () => elements(el).map(geometryElement),
  };
}

const customCache = new WeakMap<Element, PresetShapeDefinition>();

function customDefinition(element: Element): PresetShapeDefinition {
  let def = customCache.get(element);
  if (!def) {
    def = customGeometryDefinition(geometryElement(element));
    customCache.set(element, def);
  }
  return def;
}

/** Evaluate a geometry spec at a size (EMU). */
export function evaluateGeometry(spec: GeometrySpec, w: number, h: number): EvaluatedGeometry {
  if (spec.kind === 'paths') return { paths: spec.paths, textRect: { l: 0, t: 0, r: w, b: h } };
  if (spec.kind === 'custom')
    return evaluateGeometryDefinition(customDefinition(spec.element), w, h);
  return evaluatePresetGeometry(spec.preset, w, h, spec.adjust);
}

/**
 * The connection sites of a geometry spec at a size (EMU), in `a:cxnLst` order. Evaluated
 * paths have none.
 */
export function connectionSites(spec: GeometrySpec, w: number, h: number): ConnectionSite[] {
  if (spec.kind === 'paths') return [];
  if (spec.kind === 'custom') return geometryConnectionSites(customDefinition(spec.element), w, h);
  return presetConnectionSites(spec.preset, w, h, spec.adjust);
}
