// Labels and tones the adapters show about the slide canvas: the Format panel's element kind,
// and the tone of the edit notice. Shared by React and Vue, so the two cannot drift.

import type { EditorRejection, EditorSnapshot, SelectionGeometry } from '../contracts/index.ts';

/** The i18n key that names a selected element in the Format panel. */
export type ElementKindLabelKey =
  | 'panel.format.elementKind.shape'
  | 'panel.format.elementKind.picture'
  | 'panel.format.elementKind.group'
  | 'panel.format.elementKind.connector'
  | 'panel.format.elementKind.graphicFrame'
  | 'panel.format.elementKind.table';

/**
 * The Format panel's name for the selected element: a table frame is a table (the snapshot's
 * `table` names it), not an embedded object. @internal
 */
export function elementKindLabelKey(
  geometry: Pick<SelectionGeometry, 'kind' | 'nodeId'>,
  table: EditorSnapshot['table']
): ElementKindLabelKey {
  if (geometry.kind === 'graphicFrame' && table?.nodeId === geometry.nodeId)
    return 'panel.format.elementKind.table';
  return `panel.format.elementKind.${geometry.kind}`;
}

/**
 * The tone of the edit notice for a refusal: `info` for a state the edit cannot run in
 * (`disabled.*`, such as a deck open for viewing), `error` for an edit that failed. @internal
 */
export function editNoticeTone(rejection: Pick<EditorRejection, 'reason'>): 'info' | 'error' {
  return rejection.reason.startsWith('disabled.') ? 'info' : 'error';
}
