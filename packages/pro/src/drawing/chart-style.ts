/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Ported from DOCX Editor `packages/pro/src/drawing/chart-style.ts` at commit ba742d7d1.
// Fills, outlines and default colours of chart elements.
//
// A chart element's `c:spPr` decides its fill and outline when it names them. Without one the
// chart style decides: the colourful styles cycle the theme accents over the series, varying
// the brightness of the accent on every further cycle; the monochrome styles shade one accent;
// the greyscale style uses the text colour.

import {
  DRAWINGML_MAIN_NAMESPACE_URI,
  directFill,
  drawingElement,
  findDirectChild,
  isElement,
  parseEmu,
  resolveColorNode,
  schemaAttributeValue,
  type OoxmlElement,
  type ColorContext,
} from './chart-xml.ts';

const A = DRAWINGML_MAIN_NAMESPACE_URI;
const EMU_PER_POINT = 12_700;

export interface ChartPaint {
  readonly hex: string;
  readonly alpha: number;
}

export interface ChartStroke extends ChartPaint {
  readonly widthPt: number;
  readonly cap: 'butt' | 'round' | 'square';
  /** Dash and gap lengths as multiples of the width, or null for a solid line. */
  readonly dash: readonly number[] | null;
}

/**
 * What a `c:spPr` says about one element. `undefined` leaves the decision to the style; `null`
 * is an explicit `a:noFill`.
 */
export interface ChartShapeProperties {
  readonly fill: ChartPaint | null | undefined;
  readonly line: ChartStroke | null | undefined;
}

const PRESET_DASHES: ReadonlyMap<string, readonly number[]> = new Map([
  ['dash', [4, 3]],
  ['dashDot', [4, 3, 1, 3]],
  ['dot', [1, 3]],
  ['lgDash', [8, 3]],
  ['lgDashDot', [8, 3, 1, 3]],
  ['lgDashDotDot', [8, 3, 1, 3, 1, 3]],
  ['sysDash', [3, 1]],
  ['sysDashDot', [3, 1, 1, 1]],
  ['sysDashDotDot', [3, 1, 1, 1, 1, 1]],
  ['sysDot', [1, 1]],
]);

/** The mean colour of a gradient, for elements that paint one flat colour. */
function gradientMean(gradFill: OoxmlElement, colors?: ColorContext): ChartPaint | null {
  const list = findDirectChild(gradFill.children, { namespaceUri: A, localName: 'gsLst' });
  let red = 0;
  let green = 0;
  let blue = 0;
  let alpha = 0;
  let count = 0;
  for (const stop of list?.children ?? []) {
    if (!isElement(stop) || stop.local !== 'gs' || count >= 16) continue;
    const node = stop.children.find(isElement);
    const color = node ? resolveColorNode(node, colors) : null;
    if (!color) continue;
    red += Number.parseInt(color.hex.slice(0, 2), 16);
    green += Number.parseInt(color.hex.slice(2, 4), 16);
    blue += Number.parseInt(color.hex.slice(4, 6), 16);
    alpha += color.alpha;
    count += 1;
  }
  if (count === 0) return null;
  const hex = [red, green, blue]
    .map((channel) =>
      Math.round(channel / count)
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();
  return { hex, alpha: alpha / count };
}

function readFill(parent: OoxmlElement, colors?: ColorContext): ChartPaint | null | undefined {
  const fill = directFill(parent, colors);
  if (!fill.present) return undefined;
  if (fill.color) return fill.color;
  const gradient = findDirectChild(parent.children, { namespaceUri: A, localName: 'gradFill' });
  if (gradient) return gradientMean(gradient, colors);
  const pattern = findDirectChild(parent.children, { namespaceUri: A, localName: 'pattFill' });
  if (pattern) {
    const foreground = findDirectChild(pattern.children, { namespaceUri: A, localName: 'fgClr' });
    const node = foreground?.children.find(isElement);
    return node ? resolveColorNode(node, colors) : null;
  }
  return null;
}

/** The fill and outline a `c:spPr` names. */
export function readChartShapeProperties(
  spPr: OoxmlElement | null,
  colors?: ColorContext
): ChartShapeProperties {
  if (!spPr) return { fill: undefined, line: undefined };
  const fill = readFill(spPr, colors);
  const ln = findDirectChild(spPr.children, { namespaceUri: A, localName: 'ln' });
  if (!ln) return { fill, line: undefined };
  const lineFill = readFill(ln, colors);
  const widthEmu = parseEmu(schemaAttributeValue(ln.attrs, 'w'));
  const capRaw = schemaAttributeValue(ln.attrs, 'cap');
  const dashElement = findDirectChild(ln.children, { namespaceUri: A, localName: 'prstDash' });
  const dashName = dashElement ? schemaAttributeValue(dashElement.attrs, 'val') : undefined;
  if (lineFill === null) return { fill, line: null };
  // A line without its own fill keeps the style's colour but its own width and dash.
  return {
    fill,
    line:
      lineFill === undefined && widthEmu === null && dashName === undefined
        ? undefined
        : {
            hex: lineFill?.hex ?? '',
            alpha: lineFill?.alpha ?? 1,
            widthPt: widthEmu !== null ? widthEmu / EMU_PER_POINT : 0.75,
            cap: capRaw === 'rnd' ? 'round' : capRaw === 'sq' ? 'square' : 'butt',
            dash: dashName ? (PRESET_DASHES.get(dashName) ?? null) : null,
          },
  };
}

/** A line whose colour the style supplies keeps the authored width, cap and dash. */
export function mergeStroke(
  authored: ChartStroke | null | undefined,
  fallback: ChartStroke | null
): ChartStroke | null {
  if (authored === null) return null;
  if (authored === undefined) return fallback;
  if (authored.hex !== '') return authored;
  return fallback ? { ...authored, hex: fallback.hex, alpha: fallback.alpha } : null;
}

type Transform = readonly (readonly [string, number])[];

/** Brightness variations of the accent cycle, one per pass over the six accents. */
const CYCLE_VARIATIONS: readonly Transform[] = [
  [],
  [
    ['lumMod', 60_000],
    ['lumOff', 40_000],
  ],
  [['lumMod', 60_000]],
  [
    ['lumMod', 80_000],
    ['lumOff', 20_000],
  ],
  [['lumMod', 80_000]],
  [['lumMod', 50_000]],
  [
    ['lumMod', 70_000],
    ['lumOff', 30_000],
  ],
  [['lumMod', 70_000]],
  [
    ['lumMod', 50_000],
    ['lumOff', 50_000],
  ],
];

/** A theme colour with a generated transform, resolved against the theme. */
export function themeColor(
  scheme: string,
  transform: Transform,
  colors?: ColorContext
): ChartPaint | null {
  const element = drawingElement(
    'schemeClr',
    { val: scheme },
    transform.map(([name, val]) => drawingElement(name, { val: String(val) }))
  );
  return resolveColorNode(element, colors);
}

/**
 * The colour the chart style gives data item `index` of `count` (a series, or a point of a
 * plot that varies colours by point).
 */
export function styleSeriesColor(
  style: number,
  index: number,
  count: number,
  colors?: ColorContext
): ChartPaint {
  const scheme = ((style - 1) % 8) | 0;
  let color: ChartPaint | null;
  if (scheme === 1) {
    const accent = (index % 6) + 1;
    const variation = CYCLE_VARIATIONS[Math.floor(index / 6) % CYCLE_VARIATIONS.length]!;
    color = themeColor(`accent${accent}`, variation, colors);
  } else {
    // Monochrome and greyscale styles run from a dark shade to a light tint of one colour.
    const base = scheme === 0 ? 'tx1' : `accent${scheme - 1}`;
    const step = count > 1 ? index / (count - 1) : 0.5;
    const percent = Math.round((count > 1 ? 0.6 + step * 0.8 : 1) * 100_000);
    const transform: Transform =
      scheme === 0
        ? [
            ['lumMod', Math.round((0.2 + step * 0.6) * 100_000)],
            ['lumOff', Math.round((0.35 + step * 0.4) * 100_000)],
          ]
        : percent <= 100_000
          ? [['shade', percent]]
          : [['tint', 200_000 - percent]];
    color = themeColor(base, transform, colors);
  }
  return color ?? { hex: '4472C4', alpha: 1 };
}

/** Text in the gray the chart styles use for labels, as a fallback with no theme. */
export const BLACK: ChartPaint = { hex: '000000', alpha: 1 };
