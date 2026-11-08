// The `chart` chrome group: Edit Data for the selected chart, which shows the Chart Data tab of
// the side panel (`requestPane`). The toolbar shows it while one chart frame is selected
// (`EditorSnapshot.chart`), and the chart context menu has the same row. Kept apart from
// chrome-controls.ts, which spreads it into `CHROME_GROUPS` and asks its slot state here.

import type { Editor } from '../contracts/index.ts';
import type { ChromeGroup, ChromeSlotId, ToolbarCommandState } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { hasPaneListener } from './pane-requests.ts';

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
      state: { kind: 'pane', pane: 'chart' },
    },
  ],
};

/** The `chart` slot ids. @public */
export type ChartSlotId = 'chart.editData';

/**
 * The state of a `chart` slot, or undefined for another slot. Edit Data is on while one chart
 * frame is selected and a side panel or the host listens for pane requests. The tab says why a
 * chart's data cannot be edited, and shows the data read-only in view mode. @internal
 */
export function chartSlotState(editor: Editor, id: ChromeSlotId): ToolbarCommandState | undefined {
  if (id !== 'chart.editData') return undefined;
  const snap = editor.snapshot();
  const chart = snap.chart ?? null;
  const reason =
    snap.status !== 'ready'
      ? 'disabled.noEditor'
      : !chart
        ? 'disabled.noChart'
        : hasPaneListener(editor)
          ? null
          : 'disabled.noSidePanel';
  return { id, enabled: reason === null, disabledReason: reason, active: false };
}
