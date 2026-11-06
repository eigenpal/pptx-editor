/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The tree and color reads that the chart modules ported from DOCX Editor use, over the PPTX
// Editor canonical tree and the cascade's color resolution. The names and results follow DOCX
// Editor's `packages/core/src/store/package/` readers (commit ba742d7d1), so the ported chart
// modules keep their source shape.

import { resolveColor, type ColorContext, type Rgba } from '@pptx-editor.dev/core/cascade';
import { NS, type Attr, type Element, type Node } from '@pptx-editor.dev/core/store';

export type { ColorContext } from '@pptx-editor.dev/core/cascade';

/** A chart part element. */
export type OoxmlElement = Element;

export const CHART_NAMESPACE_URI = NS.c;
export const DRAWINGML_MAIN_NAMESPACE_URI = NS.a;

/** A color as charts paint it: six hex digits and an opacity from 0 to 1. */
export interface ChartColor {
  readonly hex: string;
  readonly alpha: number;
}

/** The theme's Latin heading and body faces, which chart text takes through tokens. */
export interface ChartThemeFaces {
  readonly major: string | null;
  readonly minor: string | null;
}

/** Whether a node is an element. */
export function isElement(node: Node | undefined | null): node is Element {
  return node?.type === 'element';
}

/** The first child element with this expanded name, or null. */
export function findDirectChild(
  nodes: readonly Node[],
  name: { readonly namespaceUri: string; readonly localName: string }
): Element | null {
  for (const node of nodes)
    if (node.type === 'element' && node.ns === name.namespaceUri && node.local === name.localName)
      return node;
  return null;
}

/** The value of an unprefixed attribute, or undefined. */
export function schemaAttributeValue(
  attributes: readonly Attr[],
  localName: string
): string | undefined {
  for (const attribute of attributes)
    if (attribute.local === localName && attribute.ns === '') return attribute.value;
  return undefined;
}

const MAX_EMU = 20_116_800;

/** A non-negative EMU length, clamped; null when absent or not an integer. */
export function parseEmu(value: string | undefined): number | null {
  if (value === undefined || !/^-?\d{1,15}$/.test(value)) return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return null;
  return Math.min(Math.max(parsed, 0), MAX_EMU);
}

function hex(rgba: Rgba): ChartColor {
  const part = (v: number): string =>
    Math.round(Math.min(255, Math.max(0, v)))
      .toString(16)
      .padStart(2, '0');
  return { hex: `${part(rgba.r)}${part(rgba.g)}${part(rgba.b)}`.toUpperCase(), alpha: rgba.a };
}

/** Resolve a DrawingML color element, or null when it is not one or does not resolve. */
export function resolveColorNode(node: Node | undefined, colors?: ColorContext): ChartColor | null {
  if (!isElement(node) || !colors) return null;
  const rgba = resolveColor(node, colors);
  return rgba ? hex(rgba) : null;
}

const FILL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/**
 * The fill element of a properties element: `present` when it names one, and the color of a
 * solid fill.
 */
export function directFill(
  parent: Element,
  colors?: ColorContext
): { readonly present: boolean; readonly color: ChartColor | null } {
  const fill = parent.children.find(
    (c): c is Element => c.type === 'element' && c.ns === NS.a && FILL_NAMES.has(c.local)
  );
  if (!fill) return { present: false, color: null };
  if (fill.local !== 'solidFill') return { present: true, color: null };
  const color = fill.children.find(isElement);
  return { present: true, color: resolveColorNode(color, colors) };
}

let generatedIds = 0;

/** A generated DrawingML element (never from a file), for theme colors with transforms. */
export function drawingElement(
  local: string,
  attrs: Readonly<Record<string, string>>,
  children: readonly Element[] = []
): Element {
  generatedIds += 1;
  return {
    type: 'element',
    id: `chart-generated-${generatedIds}`,
    kind: 'generic',
    ns: NS.a,
    local,
    nsDecls: [],
    attrs: Object.entries(attrs).map(([name, value]) => ({ ns: '', local: name, value })),
    children,
  };
}
