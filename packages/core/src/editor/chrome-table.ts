// The table chrome: DOCX Editor's `table` group with the controls of the Table Design
// and Table Layout tabs, and their commands. Kept beside `chrome-controls.ts`, which spreads
// these into `CHROME_GROUPS`, `SLOT_COMMANDS`, and the probes; the slot names are in
// `ChromeSlotId` there. The group shows while the selection is in or on a table
// (`EditorSnapshot.table`, `tableChromeVisible`).

import type {
  Editor,
  EditorCommand,
  EditorSnapshot,
  TableBorderEdgeTarget,
  TableBorderTarget,
  TableContext,
  TableStyleOption,
} from '../contracts/index.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../cascade/index.ts';
import type { ChromeGroup, ChromeSlotId, ToolbarCommandState } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { isColorChoice } from '../automation/commands/style-commands.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

// The Borders menu order: No Border first.
const BORDERS: readonly TableBorderTarget[] = [
  'none',
  'all',
  'outside',
  'inside',
  'top',
  'bottom',
  'left',
  'right',
];
const BORDER_ICONS: Record<TableBorderTarget, string> = {
  all: 'border_all',
  outside: 'border_outer',
  inside: 'border_inner',
  top: 'border_top',
  bottom: 'border_bottom',
  left: 'border_left',
  right: 'border_right',
  none: 'border_clear',
};

/** Split Cells presets: columns × rows. */
const SPLITS = {
  cols2: { cols: 2, rows: 1 },
  rows2: { cols: 1, rows: 2 },
  cols2rows2: { cols: 2, rows: 2 },
  cols3: { cols: 3, rows: 1 },
  rows3: { cols: 1, rows: 3 },
} as const;

/** The Cell Margins presets, EMU (0.05" is 45 720; 0.1" is 91 440; 0.15" is 137 160). */
const MARGINS = {
  normal: { left: 91_440, right: 91_440, top: 45_720, bottom: 45_720 },
  none: { left: 0, right: 0, top: 0, bottom: 0 },
  narrow: { left: 45_720, right: 45_720, top: 45_720, bottom: 45_720 },
  wide: { left: 137_160, right: 137_160, top: 137_160, bottom: 137_160 },
} as const;

/** The style options, by the Table Design check box that turns each on. */
const OPTIONS = {
  headerRow: 'firstRow',
  totalRow: 'lastRow',
  bandedRows: 'bandRow',
  firstColumn: 'firstCol',
  lastColumn: 'lastCol',
  bandedColumns: 'bandCol',
} as const satisfies Record<string, TableStyleOption>;

/**
 * One icon for each style option: a box with the part the style sets apart (a top bar, a bottom
 * bar, a left or right strip, bands). The Borders menu keeps the `border_*` icons.
 */
const OPTION_ICONS: Record<keyof typeof OPTIONS, string> = {
  headerRow: 'toolbar',
  totalRow: 'call_to_action',
  bandedRows: 'table_rows',
  // `dock_to_right` draws its strip on the left, `dock_to_left` on the right.
  firstColumn: 'dock_to_right',
  lastColumn: 'dock_to_left',
  bandedColumns: 'view_column',
};

/** The pen the Borders menu draws with: 1 pt in the theme's Text 1, the default. */
const PEN = { color: { theme: 'tx1' }, width: 1 } as const;

/**
 * The `table` group, as DOCX Editor's: Insert > Table, then the controls of the Table
 * Design and Table Layout tabs. It is contextual (not in the default bar): a menu places
 * `table.insert`, and a toolbar shows the group while the selection is in a table
 * (`tableChromeVisible`). The ribbon arranges these slots on its Insert, Table Design, and
 * Table Layout tabs; slot names never take a tab name (`editor-ui-parity` design, "Names").
 */
export const TABLE_GROUPS: readonly ChromeGroup[] = [
  {
    id: 'table',
    labelKey: 'toolbar.groups.table',
    contextual: true,
    controls: [
      {
        // A grid of rows and columns.
        id: 'insert',
        shape: 'tableGrid',
        labelKey: 'table.insert',
        paths: icon('table'),
        state: { kind: 'value' },
      },
      // The Table Style Options come first, then the Table Styles gallery.
      ...(Object.keys(OPTIONS) as (keyof typeof OPTIONS)[]).map((id) => ({
        id,
        labelKey: `table.${id}` as const,
        toggle: 'pressed' as const,
        paths: icon(OPTION_ICONS[id]),
        state: { kind: 'command' as const },
      })),
      {
        // The deck's styles: the adapter lists them (`tableStyleChoices`), as the font picker.
        id: 'style',
        shape: 'dropdown',
        labelKey: 'table.style',
        toggle: 'choice',
        paths: icon('table'),
        state: { kind: 'value' },
      },
      {
        id: 'cellFill',
        shape: 'colorSplit',
        labelKey: 'table.cellFill',
        paths: icon('format_color_fill'),
        state: { kind: 'value' },
      },
      {
        id: 'borderTarget',
        shape: 'dropdown',
        labelKey: 'table.borderTarget',
        paths: icon('border_all'),
        options: BORDERS.map((b) => ({
          value: b,
          labelKey: `table.border.${b}` as const,
          paths: icon(BORDER_ICONS[b]),
        })),
        state: { kind: 'value' },
      },
      {
        id: 'select',
        shape: 'dropdown',
        labelKey: 'table.select',
        paths: icon('select_all'),
        // The order: Select Table, Select Column, Select Row.
        options: (['table', 'column', 'row'] as const).map((r) => ({
          value: r,
          labelKey: `table.region.${r}` as const,
        })),
        state: { kind: 'value' },
      },
      {
        id: 'delete',
        shape: 'dropdown',
        labelKey: 'table.delete',
        paths: icon('delete_sweep'),
        // The order: Delete Columns, Delete Rows, Delete Table.
        options: (['columns', 'rows', 'table'] as const).map((r) => ({
          value: r,
          labelKey: `table.deleteRegion.${r}` as const,
        })),
        state: { kind: 'value' },
      },
      {
        id: 'insertRowAbove',
        labelKey: 'table.insertRowAbove',
        paths: icon('add_row_above'),
        state: { kind: 'command' },
      },
      {
        id: 'insertRowBelow',
        labelKey: 'table.insertRowBelow',
        paths: icon('add_row_below'),
        state: { kind: 'command' },
      },
      {
        id: 'insertColumnLeft',
        labelKey: 'table.insertColumnLeft',
        paths: icon('add_column_left'),
        state: { kind: 'command' },
      },
      {
        id: 'insertColumnRight',
        labelKey: 'table.insertColumnRight',
        paths: icon('add_column_right'),
        state: { kind: 'command' },
      },
      {
        id: 'mergeCells',
        labelKey: 'table.mergeCells',
        paths: icon('call_merge'),
        state: { kind: 'command' },
      },
      {
        id: 'splitCell',
        shape: 'dropdown',
        labelKey: 'table.splitCell',
        paths: icon('call_split'),
        options: (Object.keys(SPLITS) as (keyof typeof SPLITS)[]).map((v) => ({
          value: v,
          labelKey: `table.split.${v}` as const,
        })),
        state: { kind: 'value' },
      },
      {
        id: 'distributeRows',
        labelKey: 'table.distributeRows',
        paths: icon('height'),
        state: { kind: 'command' },
      },
      {
        id: 'distributeColumns',
        labelKey: 'table.distributeColumns',
        paths: icon('fit_width'),
        state: { kind: 'command' },
      },
      {
        id: 'cellMargins',
        shape: 'dropdown',
        labelKey: 'table.cellMargins',
        paths: icon('padding'),
        options: (Object.keys(MARGINS) as (keyof typeof MARGINS)[]).map((v) => ({
          value: v,
          labelKey: `table.margins.${v}` as const,
        })),
        state: { kind: 'value' },
      },
    ],
  },
];

/** The fixed commands of the table's command slots. */
export const TABLE_SLOT_COMMANDS: Readonly<Record<string, EditorCommand>> = {
  ...Object.fromEntries(
    Object.entries(OPTIONS).map(([id, option]) => [
      `table.${id}`,
      { type: 'toggleTableStyleOption', option } satisfies EditorCommand,
    ])
  ),
  'table.insertRowAbove': { type: 'insertRow', where: 'above' },
  'table.insertRowBelow': { type: 'insertRow', where: 'below' },
  'table.insertColumnLeft': { type: 'insertColumn', where: 'left' },
  'table.insertColumnRight': { type: 'insertColumn', where: 'right' },
  'table.mergeCells': { type: 'mergeCells' },
  'table.distributeRows': { type: 'distributeRows' },
  'table.distributeColumns': { type: 'distributeColumns' },
};

/** The command a table value slot dispatches for a picked value, or null. */
export function tableCommandForSlotValue(slot: string, value: unknown): EditorCommand | null {
  switch (slot) {
    case 'table.insert': {
      const v = value as { rows?: unknown; cols?: unknown } | null;
      return v && typeof v.rows === 'number' && typeof v.cols === 'number'
        ? { type: 'insertTable', rows: v.rows, cols: v.cols }
        : null;
    }
    case 'table.style':
      return typeof value === 'string' && value ? { type: 'setTableStyle', styleId: value } : null;
    case 'table.cellFill':
      return isColorChoice(value) || value === null ? { type: 'setCellFill', color: value } : null;
    case 'table.borderTarget':
      if (!BORDERS.includes(value as TableBorderTarget)) return null;
      return value === 'none'
        ? { type: 'setTableBorders', scope: 'none' }
        : {
            type: 'setTableBorders',
            scope: value as TableBorderEdgeTarget,
            spec: PEN,
          };
    case 'table.select':
      return value === 'table' || value === 'row' || value === 'column'
        ? { type: 'selectTableRegion', region: value }
        : null;
    case 'table.delete':
      return value === 'rows'
        ? { type: 'deleteRow' }
        : value === 'columns'
          ? { type: 'deleteColumn' }
          : value === 'table'
            ? { type: 'deleteTable' }
            : null;
    case 'table.splitCell': {
      const split =
        typeof value === 'string' && Object.hasOwn(SPLITS, value)
          ? SPLITS[value as keyof typeof SPLITS]
          : null;
      return split ? { type: 'splitCell', ...split } : null;
    }
    case 'table.cellMargins': {
      const margins =
        typeof value === 'string' && Object.hasOwn(MARGINS, value)
          ? MARGINS[value as keyof typeof MARGINS]
          : null;
      return margins ? { type: 'setCellMargins', margins } : null;
    }
    default:
      return null;
  }
}

/** Probe values for the table value slots (never executed). */
export const TABLE_PROBES: Readonly<Record<string, unknown>> = {
  'table.insert': { rows: 2, cols: 2 },
  'table.style': DEFAULT_TABLE_STYLE_ID,
  'table.cellFill': '#000000',
  'table.borderTarget': 'all',
  'table.select': 'table',
  'table.delete': 'rows',
  'table.cellMargins': 'normal',
};

/** Whether a table toggle slot is on: the style option it controls. */
export function tableSlotActive(slot: string, snap: EditorSnapshot): boolean {
  const option = OPTIONS[slot.slice('table.'.length) as keyof typeof OPTIONS];
  return slot.startsWith('table.') && !!option && !!snap.table?.options.includes(option);
}

/** The value a table slot shows: the table style, or the first selected cell's shading. */
export function tableSlotValue(slot: string, snap: EditorSnapshot): string | undefined {
  if (slot === 'table.style') return snap.table?.styleId;
  if (slot === 'table.cellFill') return snap.table?.cellFill;
  return undefined;
}

/**
 * The state of a table slot that one probe cannot give: Split Cells is enabled when any of its
 * presets can run, as the Split Cells dialog checks its numbers. @internal
 */
export function tableSlotState(editor: Editor, id: ChromeSlotId): ToolbarCommandState | null {
  if (id !== 'table.splitCell') return null;
  const results = Object.values(SPLITS).map((s) => editor.can({ type: 'splitCell', ...s }));
  const ok = results.some((r) => r.ok);
  const reason = results.find((r) => !r.ok);
  return {
    id,
    enabled: ok,
    disabledReason: ok || !reason || reason.ok ? null : reason.reason,
    active: false,
  };
}

/** A slot of the `table` group that a table selection shows (every one but Insert > Table). @public */
export type TableChromeSlotId = Exclude<Extract<ChromeSlotId, `table.${string}`>, 'table.insert'>;

/** Every {@link TableChromeSlotId}, in the group's order. @public */
export const TABLE_CHROME_SLOT_IDS: readonly TableChromeSlotId[] = TABLE_GROUPS[0].controls
  .filter((c) => c.id !== 'insert')
  .map((c) => `table.${c.id}` as TableChromeSlotId);

/** Whether a slot is one of the table-only ones, as DOCX Editor's `isTableChromeSlot`. @public */
export function isTableChromeSlot(slot: ChromeSlotId): slot is TableChromeSlotId {
  return (TABLE_CHROME_SLOT_IDS as readonly string[]).includes(slot);
}

/** Contextual table chrome shows while the engine reports a table, as DOCX Editor's. @public */
export function tableChromeVisible(table: TableContext | null | undefined): boolean {
  return table != null;
}
