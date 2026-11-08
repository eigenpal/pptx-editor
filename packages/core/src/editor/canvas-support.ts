// Labels and tones the adapters show about the slide canvas: the Format panel's element kind,
// and the tone of the edit notice. Shared by React and Vue, so the two cannot drift.

import type {
  EditorRejection,
  EditorSnapshot,
  SelectionGeometry,
  TableStyleOption,
} from '../contracts/index.ts';

/** The i18n key that names a selected element in the Format panel. */
export type ElementKindLabelKey =
  | 'panel.format.elementKind.shape'
  | 'panel.format.elementKind.picture'
  | 'panel.format.elementKind.group'
  | 'panel.format.elementKind.connector'
  | 'panel.format.elementKind.graphicFrame'
  | 'panel.format.elementKind.table'
  | 'panel.format.elementKind.chart'
  | 'panel.format.elementKind.diagram';

/**
 * The Format panel's name for the selected element: a table, chart, or diagram frame is named
 * for what it shows (the snapshot's `table`, `chart`, and `diagram` name them), not as an
 * embedded object. @internal
 */
export function elementKindLabelKey(
  geometry: Pick<SelectionGeometry, 'kind' | 'nodeId'>,
  table: EditorSnapshot['table'],
  objects: Pick<EditorSnapshot, 'chart' | 'diagram'> = {}
): ElementKindLabelKey {
  if (geometry.kind === 'graphicFrame') {
    if (table?.nodeId === geometry.nodeId) return 'panel.format.elementKind.table';
    if (objects.chart?.nodeId === geometry.nodeId) return 'panel.format.elementKind.chart';
    if (objects.diagram?.nodeId === geometry.nodeId) return 'panel.format.elementKind.diagram';
  }
  return `panel.format.elementKind.${geometry.kind}`;
}

/**
 * The Format panel shows Rotation for the selected element: a table frame does not rotate, as
 * in slide editors. @internal
 */
export function elementRotates(
  geometry: Pick<SelectionGeometry, 'nodeId'>,
  table: EditorSnapshot['table']
): boolean {
  return table?.nodeId !== geometry.nodeId;
}

/**
 * The size the Format panel writes for a typed width (`cx`) or height (`cy`), in EMU. With the
 * aspect ratio locked (`noChangeAspect`), the other side follows, as a corner handle keeps it.
 * A side of 0 has no ratio to keep. @internal
 */
export function formatPanelSize(
  geometry: Pick<SelectionGeometry, 'cx' | 'cy' | 'lockAspectRatio'>,
  side: 'cx' | 'cy',
  value: number
): { cx: number; cy: number } {
  const v = Math.max(0, Math.round(value));
  const { cx, cy } = geometry;
  if (!geometry.lockAspectRatio || cx <= 0 || cy <= 0)
    return side === 'cx' ? { cx: v, cy } : { cx, cy: v };
  return side === 'cx'
    ? { cx: v, cy: Math.round((v * cy) / cx) }
    : { cx: Math.round((v * cx) / cy), cy: v };
}

/** One style option check box of the Format panel's Table section. @internal */
export interface TableOptionRow {
  readonly option: TableStyleOption;
  readonly labelKey:
    | 'table.headerRow'
    | 'table.totalRow'
    | 'table.bandedRows'
    | 'table.firstColumn'
    | 'table.lastColumn'
    | 'table.bandedColumns';
}

/** The Format panel's style option check boxes, in the Table Design order. @internal */
export const TABLE_OPTION_ROWS: readonly TableOptionRow[] = [
  { option: 'firstRow', labelKey: 'table.headerRow' },
  { option: 'lastRow', labelKey: 'table.totalRow' },
  { option: 'bandRow', labelKey: 'table.bandedRows' },
  { option: 'firstCol', labelKey: 'table.firstColumn' },
  { option: 'lastCol', labelKey: 'table.lastColumn' },
  { option: 'bandCol', labelKey: 'table.bandedColumns' },
];

/**
 * The tone of the edit notice for a refusal: `info` for a state the edit cannot run in
 * (`disabled.*`, such as a deck open for viewing), `error` for an edit that failed. @internal
 */
export function editNoticeTone(rejection: Pick<EditorRejection, 'reason'>): 'info' | 'error' {
  return rejection.reason.startsWith('disabled.') ? 'info' : 'error';
}
