/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Ported from DOCX Editor `packages/pro/src/drawing/chart-text.ts` at commit ba742d7d1.
// Chart text: run styles resolved through the chart's text property layers, widths measured
// with the layout's text measurer, and the label boxes the chart lowers to text shapes.
//
// Chart text inherits in layers: the chart space `c:txPr`, then the element's own `c:txPr`
// (axis, legend, title, data labels), then rich text run properties. Automatic layout sizes the
// plot area from the widths the layout's measurer gives, so the labels fit the space it keeps
// for them. Each label is set aligned inside a frame a little wider than its text, so a width
// difference moves no text off its anchor.

import { resolveThemeFont, type Theme } from '@pptx-editor.dev/core/cascade';
import type { TextMeasurer } from '@pptx-editor.dev/core/layout';
import { isValidXmlText } from '@pptx-editor.dev/core/store';
import {
  DRAWINGML_MAIN_NAMESPACE_URI,
  findDirectChild,
  isElement,
  resolveColorNode,
  schemaAttributeValue,
  type OoxmlElement,
  type ColorContext,
} from './chart-xml.ts';

const A = DRAWINGML_MAIN_NAMESPACE_URI;
export const MAX_LABEL_UNITS = 512;
const MAX_FACE_UNITS = 128;
const MAX_LINES = 16;

/**
 * The ascent and descent of the default chart face (Calibri: 1950 and 550 of 2048 units), as
 * multiples of the size, for chart text laid out without a measurer.
 */
const DEFAULT_ASCENT = 0.952;
const DEFAULT_DESCENT = 0.268;

export interface ChartTextStyle {
  readonly sizePt: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly colorHex: string;
  /** A theme token (`+mn-lt`) or a face name. */
  readonly latin: string;
  readonly eastAsian: string | null;
  readonly complex: string | null;
}

function flag(element: OoxmlElement, name: string): boolean | null {
  const raw = schemaAttributeValue(element.attrs, name);
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return null;
}

function face(element: OoxmlElement, name: string): string | null {
  const font = findDirectChild(element.children, { namespaceUri: A, localName: name });
  const typeface = font ? schemaAttributeValue(font.attrs, 'typeface') : undefined;
  return typeface && typeface.length <= MAX_FACE_UNITS && isValidXmlText(typeface)
    ? typeface
    : null;
}

/** Apply one run property layer (`a:defRPr` or `a:rPr`) over `base`. */
export function applyRunLayer(
  base: ChartTextStyle,
  layer: OoxmlElement | null,
  colors?: ColorContext
): ChartTextStyle {
  if (!layer) return base;
  const sizeRaw = schemaAttributeValue(layer.attrs, 'sz');
  const size = sizeRaw && /^\d{3,6}$/.test(sizeRaw) ? Number(sizeRaw) / 100 : null;
  const fill = findDirectChild(layer.children, { namespaceUri: A, localName: 'solidFill' });
  const colorNode = fill?.children.find(isElement);
  const color = colorNode ? resolveColorNode(colorNode, colors) : null;
  const underline = schemaAttributeValue(layer.attrs, 'u');
  return {
    sizePt: size !== null && size >= 1 && size <= 400 ? size : base.sizePt,
    bold: flag(layer, 'b') ?? base.bold,
    italic: flag(layer, 'i') ?? base.italic,
    underline: underline === undefined ? base.underline : underline !== 'none',
    colorHex: color?.hex ?? base.colorHex,
    latin: face(layer, 'latin') ?? base.latin,
    eastAsian: face(layer, 'ea') ?? base.eastAsian,
    complex: face(layer, 'cs') ?? base.complex,
  };
}

/** The first paragraph's `a:defRPr` of a `c:txPr` or `c:rich` body. */
export function bodyDefaultRun(body: OoxmlElement | null): OoxmlElement | null {
  if (!body) return null;
  const paragraph = findDirectChild(body.children, { namespaceUri: A, localName: 'p' });
  const pPr = paragraph
    ? findDirectChild(paragraph.children, { namespaceUri: A, localName: 'pPr' })
    : null;
  return pPr ? findDirectChild(pPr.children, { namespaceUri: A, localName: 'defRPr' }) : null;
}

/**
 * The clockwise turn a `c:txPr`/`c:rich` body asks for, in degrees, or null for automatic
 * (no `rot`, or the `-60000000` sentinel).
 */
export function bodyRotation(body: OoxmlElement | null): number | null {
  const bodyPr = body
    ? findDirectChild(body.children, { namespaceUri: A, localName: 'bodyPr' })
    : null;
  if (!bodyPr) return null;
  const vert = schemaAttributeValue(bodyPr.attrs, 'vert');
  if (vert === 'vert') return 90;
  if (vert === 'vert270') return -90;
  const raw = schemaAttributeValue(bodyPr.attrs, 'rot');
  if (raw === undefined || !/^-?\d{1,9}$/.test(raw)) return null;
  const value = Number(raw);
  if (value === -60_000_000 || Math.abs(value) > 5_400_000) return null;
  return value / 60_000;
}

/** What chart text measures with while one chart lays out. */
interface ChartTextContext {
  readonly measurer: TextMeasurer;
  readonly theme: Theme;
}

let context: ChartTextContext | null = null;

/**
 * Run `layout` with the measurer and theme that chart text measures with. Chart layout is
 * synchronous, so one context serves the whole call.
 */
export function withChartText<T>(next: ChartTextContext, layout: () => T): T {
  const previous = context;
  context = next;
  try {
    return layout();
  } finally {
    context = previous;
  }
}

/** The typeface a face slot names: a theme token (`+mn-lt`) resolves through the theme. */
export function chartTypeface(face: string, theme: Theme): string {
  return resolveThemeFont(face, theme) ?? theme.minorFont.latin ?? 'Calibri';
}

/** The face the layout's measurer sets `style` in. */
function faceOf(style: ChartTextStyle, { measurer, theme }: ChartTextContext) {
  return measurer.resolve({
    family: chartTypeface(style.latin, theme),
    bold: style.bold,
    italic: style.italic,
  });
}

/** Advance width of `text` in points, measured with the layout's measurer. */
export function measureTextWidth(text: string, style: ChartTextStyle): number {
  if (!context) return 0;
  const measured = text.length > MAX_LABEL_UNITS ? text.slice(0, MAX_LABEL_UNITS) : text;
  if (measured === '') return 0;
  const face = faceOf(style, context);
  let total = 0;
  for (const advance of context.measurer.advances(face, measured)) total += advance;
  return total * style.sizePt;
}

/**
 * The ascent and descent of the face that sets `style`, as multiples of the size. The engine
 * lays out the lines of chart labels on the face's line (`LineBasis` `'face'` in
 * `@pptx-editor.dev/core/layout`), so the room that a chart keeps
 * for its labels follows the face, not one font.
 */
function faceExtent(style: ChartTextStyle): { readonly ascent: number; readonly descent: number } {
  const metrics = context ? faceOf(style, context).metrics : null;
  return metrics && metrics.ascent > 0 && metrics.descent >= 0
    ? metrics
    : { ascent: DEFAULT_ASCENT, descent: DEFAULT_DESCENT };
}

/** Height of one line of `style`, in points: the face's ascent plus descent. */
export function lineHeight(style: ChartTextStyle): number {
  const { ascent, descent } = faceExtent(style);
  return style.sizePt * (ascent + descent);
}

/** Distance from the top of a line of `style` to its baseline, in points: the face's ascent. */
export function baselineDrop(style: ChartTextStyle): number {
  return style.sizePt * faceExtent(style).ascent;
}

/** One run of label text with its style. */
export interface ChartTextRun {
  readonly text: string;
  readonly style: ChartTextStyle;
}

/** Height of paragraphs of runs: each line as tall as its largest run. */
export function paragraphsHeight(
  paragraphs: readonly (readonly ChartTextRun[])[],
  fallback: ChartTextStyle
): number {
  let total = 0;
  for (const runs of paragraphs) {
    total +=
      runs.length > 0
        ? Math.max(...runs.map((run) => lineHeight(run.style)))
        : lineHeight(fallback);
  }
  return total;
}

/** Measured width of the widest of `paragraphs`. */
export function paragraphsWidth(paragraphs: readonly (readonly ChartTextRun[])[]): number {
  let widest = 0;
  for (const runs of paragraphs) {
    widest = Math.max(
      widest,
      runs.reduce((sum, run) => sum + measureTextWidth(run.text, run.style), 0)
    );
  }
  return widest;
}

/** Paragraphs with every line break inside a run made a paragraph of its own. */
export function splitLines(
  paragraphs: readonly (readonly ChartTextRun[])[]
): (readonly ChartTextRun[])[] {
  const out: ChartTextRun[][] = [];
  for (const runs of paragraphs) {
    let current: ChartTextRun[] = [];
    for (const run of runs) {
      const lines = run.text.split(/\r\n|\r|\n/);
      for (const [index, line] of lines.entries()) {
        if (index > 0) {
          out.push(current);
          current = [];
        }
        if (line !== '') current.push({ text: line, style: run.style });
      }
    }
    out.push(current);
  }
  return out.slice(0, MAX_LINES);
}

export type ChartTextAlign = 'l' | 'ctr' | 'r';

/** Label frame and text in chart points; `chart-elements.ts` converts to EMU. */
export interface ChartLabel {
  /** Paragraphs of runs. */
  readonly paragraphs: readonly (readonly ChartTextRun[])[];
  readonly align: ChartTextAlign;
  /** Frame in chart points, before rotation about its center. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** Clockwise turn about the frame center, in degrees. */
  readonly rotation: number;
}
