// The `chart` chrome group: Edit Data for the selected chart. The toolbar shows it while one
// chart frame is selected (`EditorSnapshot.chart`), and the chart context menu has the same
// row. Kept apart from chrome-controls.ts, which spreads it into `CHROME_GROUPS` and asks its
// slot state here.

import type { Editor } from '../contracts/index.ts';
import type { ChromeGroup, ChromeSlotId, ToolbarCommandState } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';

/** The `chart` group. @internal */
export const CHART_CHROME_GROUP: ChromeGroup = {
  id: 'chart',
  labelKey: 'toolbar.groups.chart',
  contextual: true,
  controls: [
    {
      id: 'editData',
      labelKey: 'chartData.editData',
      paths: ICON_PATHS['table_chart'] ?? [],
      state: { kind: 'dialog', dialog: 'chartData' },
    },
  ],
};

/** The `chart` slot ids. @public */
export type ChartSlotId = 'chart.editData';

/**
 * The state of a `chart` slot, or undefined for another slot. Edit Data is on while one chart
 * frame with editable data is selected in editing mode. @internal
 */
export function chartSlotState(editor: Editor, id: ChromeSlotId): ToolbarCommandState | undefined {
  if (id !== 'chart.editData') return undefined;
  const snap = editor.snapshot();
  const chart = snap.chart ?? null;
  const reason =
    snap.status !== 'ready'
      ? 'disabled.noEditor'
      : snap.mode === 'view'
        ? 'disabled.viewing'
        : !chart
          ? 'disabled.noChart'
          : chart.editable
            ? null
            : (chart.reason ?? 'errors.chartData.unsupportedType');
  return { id, enabled: reason === null, disabledReason: reason, active: false };
}
