// Read and replace the data of the chart behind a chart frame: the `setChartData` command and
// the read that the Chart Data tab shows. The chart part, and an embedded workbook when its
// structure allows it, are written again whole as bytes, in one transaction (`removePart` and
// `addPart`).

import {
  attr,
  binaryPartOf,
  child,
  decodeXml,
  nodeIndex,
  NS,
  parseXmlPart,
  REL,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { contentElements } from '../../cascade/alternate-content.ts';
import type { ChartData, ChartDataRead, ChartWorkbookState } from '../../contracts/neutral.ts';
import { toBase64 } from './base64.ts';
import { readChartModel, type ChartModel } from './chart-data-model.ts';
import { planChartEdit, rewriteChartSpace } from './chart-data-write.ts';
import { chartWorkbook, updatedWorkbook } from './chart-workbook.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';

/** The most bytes of chart XML an edit reads, as the cascade reads (`chart-part.ts`). */
const MAX_CHART_PART_BYTES = 16 * 1024 * 1024;

/** The chart part behind a chart frame. */
export interface ChartFrame {
  readonly part: string;
  readonly chartSpace: Element;
}

/** The chart part that a graphic frame of `slidePart` shows, or undefined. */
export function chartFrameOf(
  pkg: PptxPackage,
  slidePart: string,
  nodeId: string
): ChartFrame | undefined {
  const xml = xmlPartOf(pkg, slidePart);
  const node = xml ? nodeIndex(xml).get(nodeId)?.node : undefined;
  if (node?.type !== 'element' || node.kind !== 'graphicFrame') return undefined;
  const data = child(child(node, 'a:graphic'), 'a:graphicData');
  const ref = contentElements(data).find((e) => e.ns === NS.c && e.local === 'chart');
  const id = attr(ref, 'r:id');
  const rel = id ? relationshipsOf(pkg, slidePart).find((r) => r.id === id) : undefined;
  if (!rel || rel.external || !rel.target || rel.type !== REL.chart) return undefined;
  let root = xmlPartOf(pkg, rel.target)?.root;
  if (!root) {
    const bytes = binaryPartOf(pkg, rel.target);
    if (!bytes || bytes.length > MAX_CHART_PART_BYTES) return undefined;
    const tree = parseXmlPart(rel.target, decodeXml(bytes));
    root = tree.ok ? tree.part.root : undefined;
  }
  if (!root || root.ns !== NS.c || root.local !== 'chartSpace') return undefined;
  return { part: rel.target, chartSpace: root };
}

/** The data of a read chart, in plot order, with each item's position as its source. */
function dataOf(model: ChartModel): Omit<ChartData, 'workbook'> {
  return {
    kind: model.kind,
    categories: model.categories.map((label, source) => ({ label, source })),
    series: model.series.map((s, source) => ({
      name: s.name,
      values: model.categories.map((_, i) => {
        const v = s.val?.values[i];
        const n = v === undefined || v.trim() === '' ? NaN : Number(v);
        return Number.isFinite(n) ? n : null;
      }),
      source,
    })),
  };
}

/** Where the chart's workbook is, and whether an edit can write it: a dry run of the edit. */
function workbookState(pkg: PptxPackage, frame: ChartFrame, model: ChartModel): ChartWorkbookState {
  const book = chartWorkbook(pkg, frame.part, frame.chartSpace);
  if (book.state !== 'embedded') return book.state;
  const bytes = book.part ? binaryPartOf(pkg, book.part) : undefined;
  if (!bytes || !model.layout) return 'kept';
  try {
    const plan = planChartEdit(model, dataOf(model));
    return updatedWorkbook(bytes, plan, model.layout) ? 'embedded' : 'kept';
  } catch {
    return 'kept';
  }
}

/** The data of the chart frame `nodeId` of `slidePart`, or why it cannot be edited. */
export function readChartData(pkg: PptxPackage, slidePart: string, nodeId: string): ChartDataRead {
  const frame = chartFrameOf(pkg, slidePart, nodeId);
  if (!frame) return { ok: false, reason: 'errors.chartData.noChart' };
  const read = readChartModel(frame.chartSpace);
  if (!read.ok) return read;
  return {
    ok: true,
    data: { ...dataOf(read.model), workbook: workbookState(pkg, frame, read.model) },
  };
}

/** Whether the chart frame's data can be edited, and why not. Reads no workbook. */
export function chartEditable(
  pkg: PptxPackage,
  slidePart: string,
  nodeId: string
): { readonly editable: boolean; readonly reason?: string } | undefined {
  const frame = chartFrameOf(pkg, slidePart, nodeId);
  if (!frame) return undefined;
  const read = readChartModel(frame.chartSpace);
  return read.ok ? { editable: true } : { editable: false, reason: read.reason };
}

/** `c:externalData` with its automatic update turned off (`c:autoUpdate val="0"`). */
function withoutAutoUpdate(chartSpace: Element): Element {
  const ext = child(chartSpace, 'c:externalData');
  if (!ext) return chartSpace;
  const off: Element = {
    type: 'element',
    id: '',
    kind: 'generic',
    ns: NS.c,
    local: 'autoUpdate',
    prefix: 'c',
    nsDecls: [],
    attrs: [{ ns: '', local: 'val', value: '0' }],
    children: [],
  };
  const has = ext.children.some((k) => k.type === 'element' && k.local === 'autoUpdate');
  const next: Element = {
    ...ext,
    children: has
      ? ext.children.map((k) => (k.type === 'element' && k.local === 'autoUpdate' ? off : k))
      : [off, ...ext.children],
  };
  return { ...chartSpace, children: chartSpace.children.map((k) => (k === ext ? next : k)) };
}

/** Write a part again whole in this transaction. */
function replacePart(
  batch: OpBatch,
  part: string,
  content: { xml: string } | { base64: string }
): void {
  batch.push({ op: 'removePart', part });
  batch.push({ op: 'addPart', part, ...content });
}

/**
 * Replace the data of the chart frame `nodeId` of `slidePart` (`setChartData`). Throws
 * `CommandRefusal` when the chart's data cannot be edited or the new data is not valid.
 */
export function setChartData(
  batch: OpBatch,
  slidePart: string,
  nodeId: string,
  data: Omit<ChartData, 'workbook'>
): void {
  const pkg = batch.pkg;
  const frame = chartFrameOf(pkg, slidePart, nodeId);
  if (!frame) throw new CommandRefusal('errors.chartData.noChart');
  const read = readChartModel(frame.chartSpace);
  if (!read.ok) throw new CommandRefusal(read.reason);
  const plan = planChartEdit(read.model, data);
  let chartSpace = rewriteChartSpace(frame.chartSpace, plan);
  const book = chartWorkbook(pkg, frame.part, frame.chartSpace);
  // A linked workbook is never opened: its automatic update would bring the old values back.
  if (book.state === 'external') chartSpace = withoutAutoUpdate(chartSpace);
  if (book.state === 'embedded' && book.part && read.model.layout) {
    const bytes = binaryPartOf(pkg, book.part);
    const next = bytes ? updatedWorkbook(bytes, plan, read.model.layout) : undefined;
    if (next) replacePart(batch, book.part, { base64: toBase64(next) });
    // The workbook took the chart's data before this edit: refuse an edit it cannot take, so
    // the chart and its workbook never part in silence.
    else if (workbookState(pkg, frame, read.model) === 'embedded')
      throw new CommandRefusal('errors.chartData.workbook');
  }
  writeChartPart(batch, frame.part, chartSpace);
}

/**
 * Write a chart part again whole, from its new `c:chartSpace`. Throws `CommandRefusal` when a
 * value cannot be written as XML.
 */
export function writeChartPart(batch: OpBatch, part: string, chartSpace: Element): void {
  let xml: string;
  try {
    xml = serializeXmlPart({ partName: part, root: chartSpace, nextId: 0 });
  } catch {
    throw new CommandRefusal('errors.chartData.invalidText');
  }
  // As bytes, as the reader keeps a chart part: a collaboration peer and a reopened copy then
  // hold the same part, with its attributes in the order written.
  replacePart(batch, part, { base64: toBase64(new TextEncoder().encode(xml)) });
}
