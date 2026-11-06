// The chart part behind one chart graphic frame, after DOCX Editor's
// `packages/core/src/store/package/chart-parts.ts` (commit ba742d7d1).
//
// A chart `a:graphicData` holds `c:chart`, whose `r:id` names the chart part (`c:chartSpace`,
// ECMA-376 Part 1, 21.2.2.27) on the slide part. The chart part may name an embedded workbook
// or an external data source; neither is ever read: a chart renders from the values cached in
// its own XML (`c:numCache`, `c:strCache`). The package keeps a chart part as bytes, so its
// save is byte-identical; this module parses a read-only tree from those bytes, through the
// same bounded reader as every other part, and caches it by the bytes.

import {
  attr,
  binaryPartOf,
  child,
  contentTypeOf,
  decodeXml,
  NS,
  parseXmlPart,
  REL,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { contentElements } from './alternate-content.ts';
import type { ColorContext } from './color.ts';
import { readColorMap, type Theme } from './theme.ts';

/** The content type of a chart part (ECMA-376 Part 1, 14.2.1). */
export const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

/** The most bytes of chart XML read for one chart. Past it the chart keeps its placeholder. */
const MAX_CHART_PART_BYTES = 16 * 1024 * 1024;

/** A resolved chart part: its tree, and what its colors and fonts resolve against. */
export interface ResolvedChartPart {
  /** The chart part name. */
  readonly partName: string;
  /** The `c:chartSpace` root. Read-only: the package keeps the part's own bytes. */
  readonly chartSpace: Element;
  /**
   * The slide's color context, with the chart's own color map override
   * (`c:clrMapOvr`, ECMA-376 Part 1, 21.2.2.28) applied.
   */
  readonly colors: ColorContext;
  /** The theme, for the `+mj-lt` and `+mn-lt` fonts of chart text. */
  readonly theme: Theme;
}

const trees = new WeakMap<Uint8Array, Element | null>();

function chartSpaceOf(pkg: PptxPackage, partName: string): Element | undefined {
  const parsed = xmlPartOf(pkg, partName);
  if (parsed) return parsed.root;
  const bytes = binaryPartOf(pkg, partName);
  if (!bytes || bytes.length > MAX_CHART_PART_BYTES) return undefined;
  let root = trees.get(bytes);
  if (root === undefined) {
    const tree = parseXmlPart(partName, decodeXml(bytes));
    root = tree.ok ? tree.part.root : null;
    trees.set(bytes, root);
  }
  return root ?? undefined;
}

/**
 * The chart part a chart frame's `a:graphicData` names, or undefined when the reference, the
 * relationship, the part, or its root is missing or of another kind. An external target is
 * never fetched.
 */
export function resolveChartPart(
  pkg: PptxPackage,
  sourcePart: string,
  graphicData: Element | undefined,
  colors: ColorContext,
  theme: Theme
): ResolvedChartPart | undefined {
  const ref = contentElements(graphicData).find((e) => e.ns === NS.c && e.local === 'chart');
  const id = attr(ref, 'r:id');
  if (!id) return undefined;
  const rel = relationshipsOf(pkg, sourcePart).find((r) => r.id === id);
  if (!rel || rel.external || !rel.target || rel.type !== REL.chart) return undefined;
  const type = contentTypeOf(pkg, rel.target);
  if (type !== undefined && type.trim().toLowerCase() !== CHART_CONTENT_TYPE) return undefined;
  const chartSpace = chartSpaceOf(pkg, rel.target);
  if (!chartSpace || chartSpace.ns !== NS.c || chartSpace.local !== 'chartSpace') return undefined;
  const override = readColorMap(child(chartSpace, 'c:clrMapOvr'));
  return {
    partName: rel.target,
    chartSpace,
    colors: override ? { ...colors, colorMap: override } : colors,
    theme,
  };
}
