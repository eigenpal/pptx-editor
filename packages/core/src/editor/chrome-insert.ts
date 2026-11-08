// The `insert` chrome group: Text Box, Line, Shape, Picture, Chart, and Equation. Chart lists
// the chart types and inserts one with sample data (`insertChart`). Equation opens the
// Equation dialog, which inserts an equation, or edits the selected one (`equation-fields.ts`).
// Kept apart from chrome-controls.ts, which lists the group in `CHROME_GROUPS`.

import type { ChartType } from '../contracts/index.ts';
import type { ChromeControl, ChromeGroup } from './chrome-controls.ts';
import { LINE_KINDS } from './chrome-slot-commands.ts';
import { ICON_PATHS } from './icon-paths.ts';

/** The Equation control of the `insert` group. @internal */
const EQUATION_CONTROL: ChromeControl = {
  id: 'equation',
  labelKey: 'toolbar.equation',
  paths: ICON_PATHS['functions'] ?? [],
  state: { kind: 'dialog', dialog: 'equation' },
};

/** The chart types that Insert > Chart offers, with the Material Symbol of each. @internal */
export const CHART_INSERT_TYPES: readonly (readonly [ChartType, string])[] = [
  ['columnClustered', 'bar_chart'],
  ['barClustered', 'align_horizontal_left'],
  ['line', 'show_chart'],
  ['area', 'area_chart'],
  ['pie', 'pie_chart'],
  ['doughnut', 'donut_large'],
  ['scatter', 'scatter_plot'],
];

/** The Chart control of the `insert` group: one option for each chart type. @internal */
const CHART_CONTROL: ChromeControl = {
  id: 'chart',
  shape: 'dropdown',
  labelKey: 'toolbar.chart',
  paths: ICON_PATHS['insert_chart'] ?? [],
  options: CHART_INSERT_TYPES.map(([value, icon]) => ({
    value,
    labelKey: `chartFormat.types.${value}` as const,
    paths: ICON_PATHS[icon] ?? [],
  })),
  state: { kind: 'value' },
};

/** The `insert.chart` slot id. @public */
export type ChartInsertSlotId = 'insert.chart';

/** The `insert.equation` slot id. @public */
export type EquationSlotId = 'insert.equation';

/** The probe of the Equation slot: a well-formed linear form, only to ask `Editor.can`. @internal */
export const EQUATION_PROBES = { 'insert.equation': 'x' } as const;

/** The `insert` group. @internal */
export const INSERT_GROUP: ChromeGroup = {
  id: 'insert',
  labelKey: 'toolbar.groups.insert',
  controls: [
    {
      id: 'textBox',
      labelKey: 'toolbar.textBox',
      paths: ICON_PATHS['text_fields'] ?? [],
      state: { kind: 'command' },
    },
    {
      // Lines and connectors; the chrome draws each option as the line it inserts.
      id: 'line',
      shape: 'dropdown',
      labelKey: 'toolbar.line',
      paths: null,
      options: LINE_KINDS.map((k) => ({ value: k, labelKey: `toolbar.lines.${k}` as const })),
      state: { kind: 'value' },
    },
    {
      // A gallery of every preset in `SHAPE_GALLERY`, drawn from the spec geometry.
      id: 'shape',
      shape: 'gallery',
      labelKey: 'toolbar.shape',
      paths: ICON_PATHS['shapes'] ?? [],
      state: { kind: 'value' },
    },
    {
      id: 'image',
      labelKey: 'toolbar.image',
      paths: ICON_PATHS['image'] ?? [],
      state: { kind: 'load' },
    },
    CHART_CONTROL,
    EQUATION_CONTROL,
  ],
};
