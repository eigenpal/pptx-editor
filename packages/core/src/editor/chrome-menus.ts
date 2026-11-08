// The menu registry: the menu bar and the canvas context menus as rows over chrome slots, after
// DOCX Editor's `CHROME_MENUS` (docx-editor/packages/core/src/editor/chrome-controls.ts,
// packages/core/src/editor/chrome-menu-entries.ts). A row names a slot, so its label, icon, enabled
// state, and action come from the same registry entry as the toolbar control: a menu row and its
// toolbar twin cannot describe one capability two ways. The adapters render these rows; a host
// replaces one with `Menu.Item` or `ContextMenu.Slot`, adds its own, or turns the preset off. DOCX
// Editor commit adf0e178.
//
// The menu bar includes rows whose commands the engine provides.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type { Editor, EditorSnapshot } from '../contracts/index.ts';
import { chromeControlForSlot, chromeSlotIsToggle, type ChromeSlotId } from './chrome-controls.ts';
import { commentRowShown } from './chrome-review.ts';
import { tidySeparators } from './chrome-forms.ts';
import { ICON_PATHS } from './icon-paths.ts';

/** A row that runs one chrome slot. @public */
export interface ChromeMenuItemEntry {
  readonly kind: 'item';
  readonly slot: ChromeSlotId;
  /** The value a value slot runs with (`arrange.order` with `front`). */
  readonly value?: string;
  /**
   * A plain label for the row, in place of the slot's own label. A slot's label is a tooltip,
   * and some fold a shortcut into the text; a menu shows the shortcut in its own column.
   */
  readonly labelKey?: TranslationKey;
  /** The shortcut shown at the end of the row. */
  readonly shortcutKey?: TranslationKey;
  /**
   * For a value slot without `value`: the row runs the slot's own command instead of opening
   * its choices. The canvas menu's New slide adds a slide on the layout
   * `newSlideLayout` picks.
   */
  readonly direct?: true;
}

/** A row that opens a nested panel of rows. It has no command of its own. @public */
export interface ChromeMenuSubmenuEntry {
  readonly kind: 'submenu';
  /** A stable id, so a host can find the submenu. */
  readonly id: string;
  readonly labelKey: TranslationKey;
  readonly paths: readonly string[] | null;
  readonly items: readonly ChromeMenuEntry[];
}

/** A line between groups of rows. @public */
export interface ChromeMenuSeparatorEntry {
  readonly kind: 'separator';
}

/** One row of a chrome menu. @public */
export type ChromeMenuEntry =
  ChromeMenuItemEntry | ChromeMenuSubmenuEntry | ChromeMenuSeparatorEntry;

/**
 * Every menu id of the menu bar. Stable public API: renaming one is a breaking change, as for a
 * slot id. @public
 */
export type ChromeMenuId = 'file' | 'edit' | 'view' | 'insert' | 'format' | 'arrange' | 'slideShow';

/** One menu of the menu bar. @public */
export interface ChromeMenu {
  readonly id: ChromeMenuId;
  readonly labelKey: TranslationKey;
  readonly entries: readonly ChromeMenuEntry[];
}

const SEPARATOR: ChromeMenuSeparatorEntry = { kind: 'separator' };

const item = (
  slot: ChromeSlotId,
  labelKey?: TranslationKey,
  shortcutKey?: TranslationKey
): ChromeMenuItemEntry => ({
  kind: 'item',
  slot,
  ...(labelKey ? { labelKey } : {}),
  ...(shortcutKey ? { shortcutKey } : {}),
});

const ORDER_ROWS: readonly ChromeMenuItemEntry[] = [
  { kind: 'item', slot: 'arrange.order', value: 'front', labelKey: 'toolbar.bringToFront' },
  { kind: 'item', slot: 'arrange.order', value: 'back', labelKey: 'toolbar.sendToBack' },
  { kind: 'item', slot: 'arrange.order', value: 'forward', labelKey: 'toolbar.bringForward' },
  { kind: 'item', slot: 'arrange.order', value: 'backward', labelKey: 'toolbar.sendBackward' },
];

const ALIGN_ROWS: readonly ChromeMenuItemEntry[] = (
  ['left', 'center', 'right', 'top', 'middle', 'bottom'] as const
).map((v) => ({
  kind: 'item',
  slot: 'arrange.align',
  value: v,
  labelKey: `toolbar.objectAlign.${v}` as const,
}));

/** The menu bar, in this order: File, Edit, View, Insert, Format, Arrange, Slide Show. @public */
export const CHROME_MENUS: readonly ChromeMenu[] = [
  { id: 'file', labelKey: 'menu.file', entries: [item('file.pageSetup')] },
  {
    id: 'edit',
    labelKey: 'menu.edit',
    entries: [
      item('history.undo', 'menu.undo', 'contextMenu.shortcuts.undo'),
      item('history.redo', 'menu.redo', 'contextMenu.shortcuts.redo'),
      SEPARATOR,
      item('edit.cut', undefined, 'contextMenu.shortcuts.cut'),
      item('edit.copy', undefined, 'contextMenu.shortcuts.copy'),
      item('edit.paste', undefined, 'contextMenu.shortcuts.paste'),
      item('edit.selectAll', undefined, 'contextMenu.shortcuts.selectAll'),
      SEPARATOR,
      item('more.duplicate', undefined, 'contextMenu.shortcuts.duplicate'),
      item('more.delete', undefined, 'contextMenu.shortcuts.delete'),
      SEPARATOR,
      // Find and Replace sit in a Find submenu of Edit.
      {
        kind: 'submenu',
        id: 'find',
        labelKey: 'menu.find',
        paths: ICON_PATHS['search'] ?? null,
        items: [
          item('find.find', 'menu.findRow', 'contextMenu.shortcuts.find'),
          item('find.replace', 'menu.replaceRow', 'contextMenu.shortcuts.replace'),
        ],
      },
    ],
  },
  {
    id: 'view',
    labelKey: 'menu.view',
    entries: [
      // The View menu starts with the presentation views: Normal and Slide Sorter.
      { kind: 'item', slot: 'view.presentationView', value: 'normal' },
      { kind: 'item', slot: 'view.presentationView', value: 'slideSorter' },
      SEPARATOR,
      item('view.smartGuides'),
    ],
  },
  {
    id: 'insert',
    labelKey: 'menu.insert',
    entries: [
      // The order: New Slide, then Text Box, Table, Picture, Shape (Line sits in the Shape menu).
      // New Slide adds a slide on the next layout, as the slide context menu's row; Table opens
      // the size grid.
      {
        ...item('slide.add', 'contextMenu.newSlide', 'contextMenu.shortcuts.newSlide'),
        direct: true,
      },
      SEPARATOR,
      item('insert.textBox'),
      item('table.insert'),
      item('insert.image', 'menu.picture'),
      item('insert.shape', 'menu.shape'),
      item('insert.line', 'menu.line'),
      item('insert.equation', 'menu.equation'),
      SEPARATOR,
      item('text.link', 'menu.link', 'contextMenu.shortcuts.link'),
    ],
  },
  {
    id: 'format',
    labelKey: 'menu.format',
    entries: [
      // Font…, Paragraph…, Bullets and Numbering…, and Columns… open their dialogs, at the top
      // of the menu in this order.
      item('text.dialog', 'contextMenu.font'),
      item('paragraph.dialog', 'contextMenu.paragraph'),
      {
        kind: 'item',
        slot: 'list.bulletStyle',
        value: 'dialog',
        labelKey: 'toolbar.bulletsAndNumbering',
      },
      { kind: 'item', slot: 'paragraph.columns', value: 'dialog', labelKey: 'menu.columns' },
      SEPARATOR,
      item('text.bold', 'contextMenu.bold', 'contextMenu.shortcuts.bold'),
      item('text.italic', 'contextMenu.italic', 'contextMenu.shortcuts.italic'),
      item('text.underline', 'contextMenu.underline', 'contextMenu.shortcuts.underline'),
      item('text.strike'),
      SEPARATOR,
      item('list.bullet'),
      item('list.numbered'),
      SEPARATOR,
      item('format.painter', 'contextMenu.formatPainter', 'contextMenu.shortcuts.formatPainter'),
      item('format.clear', 'contextMenu.clearFormatting', 'contextMenu.shortcuts.clearFormatting'),
      SEPARATOR,
      // The Format menu names it Slide Background; the context menus keep the dialog's name.
      item('slide.background', 'menu.slideBackground'),
      item('image.crop'),
      // Format picture: it formats pictures only, so the row says so in both menus.
      item('image.properties'),
    ],
  },
  {
    id: 'arrange',
    labelKey: 'menu.arrange',
    entries: [
      ...ORDER_ROWS,
      SEPARATOR,
      item('arrange.group', 'contextMenu.group', 'contextMenu.shortcuts.group'),
      item('arrange.ungroup', 'contextMenu.ungroup', 'contextMenu.shortcuts.ungroup'),
      SEPARATOR,
      {
        kind: 'submenu',
        id: 'align',
        labelKey: 'menu.alignOrDistribute',
        paths: ICON_PATHS['align_horizontal_center'] ?? null,
        items: ALIGN_ROWS,
      },
    ],
  },
  {
    id: 'slideShow',
    labelKey: 'menu.slideShow',
    entries: [
      item('view.playFromStart', undefined, 'contextMenu.shortcuts.present'),
      item('view.playFromCurrent', undefined, 'contextMenu.shortcuts.presentFromCurrent'),
      SEPARATOR,
      // Slide Show > Hide Slide: a check row for the selected slides.
      item('slide.hide'),
    ],
  },
];

/**
 * What a canvas context menu opened on: shapes, pictures only, one chart, text being edited, a
 * table cell (a caret in a cell, or a cell range), or empty slide area. @public
 */
export type CanvasMenuKind =
  'shapes' | 'picture' | 'chart' | 'diagram' | 'text' | 'table' | 'tableFrame' | 'slide';

/**
 * The context menu kind for a snapshot: a table cell (a caret in a cell, or a cell range), text
 * being edited, a selection of pictures only (the snapshot has `image`), other shapes, or the
 * slide. @internal
 */
export function canvasMenuKind(
  snapshot: Pick<EditorSnapshot, 'selection' | 'image'> &
    Partial<Pick<EditorSnapshot, 'table' | 'chart' | 'diagram'>>
): CanvasMenuKind {
  const sel = snapshot.selection;
  if (sel.kind === 'cells' || (sel.kind === 'text' && sel.cell)) return 'table';
  if (snapshot.diagram) return 'diagram';
  if (sel.kind === 'text') return 'text';
  if (sel.kind === 'shapes' && snapshot.table) return 'tableFrame';
  if (sel.kind === 'shapes' && snapshot.chart) return 'chart';
  if (sel.kind === 'shapes') return snapshot.image ? 'picture' : 'shapes';
  return 'slide';
}

/** The table Insert submenu: rows above and below, columns left and right. */
const TABLE_INSERT: ChromeMenuEntry = {
  kind: 'submenu',
  id: 'tableInsert',
  labelKey: 'table.insertSubmenu',
  paths: ICON_PATHS['add'] ?? null,
  items: [
    item('table.insertRowAbove'),
    item('table.insertRowBelow'),
    item('table.insertColumnLeft'),
    item('table.insertColumnRight'),
  ],
};

/** The table Delete submenu, in the toolbar's order. */
const TABLE_DELETE: ChromeMenuEntry = {
  kind: 'submenu',
  id: 'tableDelete',
  labelKey: 'table.delete',
  paths: ICON_PATHS['delete_sweep'] ?? null,
  // The toolbar's order: Delete Columns, Delete Rows, Delete Table.
  items: (['columns', 'rows', 'table'] as const).map((v): ChromeMenuItemEntry => ({
    kind: 'item',
    slot: 'table.delete',
    value: v,
    labelKey: `table.deleteRegion.${v}`,
  })),
};

/**
 * New comment, after the clipboard rows of the text, shape, picture, and table menus. It shows
 * only while a comments pane can open a draft (`canvasMenuRows`).
 */
const NEW_COMMENT = item('review.addComment');

/**
 * The canvas context menus, one per kind of target, in a fixed order. Rows the selection
 * cannot run stay in the menu, disabled with the engine's reason. The New comment row
 * (`review.addComment`) is the exception: without a comments pane, `canvasMenuRows` leaves it
 * out. @public
 */
export const CANVAS_CONTEXT_MENUS: Readonly<Record<CanvasMenuKind, readonly ChromeMenuEntry[]>> = {
  text: [
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    SEPARATOR,
    NEW_COMMENT,
    SEPARATOR,
    item('text.bold', 'contextMenu.bold', 'contextMenu.shortcuts.bold'),
    item('text.italic', 'contextMenu.italic', 'contextMenu.shortcuts.italic'),
    item('text.underline', 'contextMenu.underline', 'contextMenu.shortcuts.underline'),
    item('text.link', 'contextMenu.link', 'contextMenu.shortcuts.link'),
    item('format.clear', 'contextMenu.clearFormatting', 'contextMenu.shortcuts.clearFormatting'),
    SEPARATOR,
    item('text.dialog', 'contextMenu.font'),
    item('paragraph.dialog', 'contextMenu.paragraph'),
  ],
  shapes: [
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    item('more.duplicate', 'contextMenu.duplicate', 'contextMenu.shortcuts.duplicate'),
    item('more.delete', 'contextMenu.delete', 'contextMenu.shortcuts.delete'),
    SEPARATOR,
    NEW_COMMENT,
    SEPARATOR,
    ...ORDER_ROWS,
    SEPARATOR,
    item('arrange.group', 'contextMenu.group', 'contextMenu.shortcuts.group'),
    item('arrange.ungroup', 'contextMenu.ungroup', 'contextMenu.shortcuts.ungroup'),
    SEPARATOR,
    item('format.painter', 'contextMenu.formatPainter', 'contextMenu.shortcuts.formatPainter'),
    SEPARATOR,
    item('image.altText', 'contextMenu.altText'),
  ],
  // The table cell menu, on the cell (or the cell range): the clipboard rows, then the row and
  // column commands of the Table Layout tab in three short submenus (Insert, Delete, Select),
  // then Merge and Split, then the text dialogs. A table selected as an object takes the
  // table frame menu.
  table: [
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    SEPARATOR,
    NEW_COMMENT,
    SEPARATOR,
    TABLE_INSERT,
    TABLE_DELETE,
    {
      kind: 'submenu',
      id: 'tableSelect',
      labelKey: 'table.select',
      paths: ICON_PATHS['select_all'] ?? null,
      // The toolbar's order: Select Table, Select Column, Select Row.
      items: (['table', 'column', 'row'] as const).map((v): ChromeMenuItemEntry => ({
        kind: 'item',
        slot: 'table.select',
        value: v,
        labelKey: `table.region.${v}`,
      })),
    },
    SEPARATOR,
    item('table.mergeCells'),
    {
      kind: 'submenu',
      id: 'splitCell',
      labelKey: 'table.splitCell',
      paths: ICON_PATHS['call_split'] ?? null,
      items: (['cols2', 'rows2', 'cols2rows2', 'cols3', 'rows3'] as const).map(
        (v): ChromeMenuItemEntry => ({
          kind: 'item',
          slot: 'table.splitCell',
          value: v,
          labelKey: `table.split.${v}`,
        })
      ),
    },
    SEPARATOR,
    item('text.dialog', 'contextMenu.font'),
    item('paragraph.dialog', 'contextMenu.paragraph'),
  ],
  // A table selected as an object (from the select handle): the shape rows without the format
  // painter, with the table's Insert submenu after the clipboard rows. On the whole table, the
  // Delete row removes the table, so it has no Delete submenu.
  tableFrame: [
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    item('more.duplicate', 'contextMenu.duplicate', 'contextMenu.shortcuts.duplicate'),
    item('more.delete', 'contextMenu.delete', 'contextMenu.shortcuts.delete'),
    SEPARATOR,
    NEW_COMMENT,
    SEPARATOR,
    TABLE_INSERT,
    SEPARATOR,
    ...ORDER_ROWS,
    SEPARATOR,
    item('arrange.group', 'contextMenu.group', 'contextMenu.shortcuts.group'),
    item('arrange.ungroup', 'contextMenu.ungroup', 'contextMenu.shortcuts.ungroup'),
    SEPARATOR,
    item('image.altText', 'contextMenu.altText'),
  ],
  // The picture menu: the shape rows, with Change Picture after Paste and Format Picture at the
  // end. This order is not checked yet.
  picture: [
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    item('image.change'),
    item('more.duplicate', 'contextMenu.duplicate', 'contextMenu.shortcuts.duplicate'),
    item('more.delete', 'contextMenu.delete', 'contextMenu.shortcuts.delete'),
    SEPARATOR,
    NEW_COMMENT,
    SEPARATOR,
    ...ORDER_ROWS,
    SEPARATOR,
    item('arrange.group', 'contextMenu.group', 'contextMenu.shortcuts.group'),
    item('arrange.ungroup', 'contextMenu.ungroup', 'contextMenu.shortcuts.ungroup'),
    SEPARATOR,
    item('format.painter', 'contextMenu.formatPainter', 'contextMenu.shortcuts.formatPainter'),
    SEPARATOR,
    item('image.altText', 'contextMenu.altText'),
    item('image.properties'),
  ],
  // The chart menu: Edit Data first, then the shape rows without the format painter.
  chart: [
    item('chart.editData'),
    SEPARATOR,
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    item('more.duplicate', 'contextMenu.duplicate', 'contextMenu.shortcuts.duplicate'),
    item('more.delete', 'contextMenu.delete', 'contextMenu.shortcuts.delete'),
    SEPARATOR,
    ...ORDER_ROWS,
    SEPARATOR,
    item('arrange.group', 'contextMenu.group', 'contextMenu.shortcuts.group'),
    item('arrange.ungroup', 'contextMenu.ungroup', 'contextMenu.shortcuts.ungroup'),
    SEPARATOR,
    item('image.altText', 'contextMenu.altText'),
  ],
  // The diagram menu: the shape commands first, then the object rows.
  diagram: [
    item('diagram.addAfter'),
    item('diagram.addBefore'),
    item('diagram.addBelow'),
    item('diagram.delete'),
    SEPARATOR,
    item('edit.cut', 'contextMenu.cut', 'contextMenu.shortcuts.cut'),
    item('edit.copy', 'contextMenu.copy', 'contextMenu.shortcuts.copy'),
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    item('more.duplicate', 'contextMenu.duplicate', 'contextMenu.shortcuts.duplicate'),
    SEPARATOR,
    ...ORDER_ROWS,
    SEPARATOR,
    item('image.altText', 'contextMenu.altText'),
  ],
  slide: [
    item('edit.paste', 'contextMenu.paste', 'contextMenu.shortcuts.paste'),
    SEPARATOR,
    item('slide.background', 'contextMenu.background'),
    SEPARATOR,
    { ...item('slide.add', 'contextMenu.newSlide'), direct: true },
  ],
};

/**
 * The rows of a canvas context menu as `editor` shows them: without the New comment row while no
 * comments pane can open a draft (`commentRowShown`), so core without the pro review pane shows
 * no dead row. Separators left alone go with it. @internal
 */
export function canvasMenuRows(
  entries: readonly ChromeMenuEntry[],
  editor: Editor | null | undefined
): readonly ChromeMenuEntry[] {
  if (commentRowShown(editor)) return entries;
  const kept = entries.filter((e) => e.kind !== 'item' || e.slot !== 'review.addComment');
  return kept.length === entries.length ? entries : tidySeparators(kept);
}

/** The key that names a menu row for host overrides: the slot, and its value if any. @internal */
export function chromeMenuRowKey(entry: Pick<ChromeMenuItemEntry, 'slot' | 'value'>): string {
  return entry.value === undefined ? entry.slot : `${entry.slot}.${entry.value}`;
}

/** A menu row with its label, shortcut, and icon read from the registry. @internal */
export interface ResolvedMenuRow {
  /** The override key: the slot, and its value if any (`arrange.order.front`). */
  readonly key: string;
  readonly slot: ChromeSlotId;
  readonly value?: string;
  readonly labelKey: TranslationKey;
  /** Values for the label's placeholders (a line width option's `{value}`). */
  readonly labelParams?: Readonly<Record<string, string | number>>;
  readonly shortcutKey?: TranslationKey;
  readonly paths: readonly string[] | null;
  /** A two-state slot (Bold): the row shows a check while it is on. */
  readonly toggle: boolean;
}

/**
 * The icon of a menu row that opens a dialog from a gallery slot, whose option has no icon of
 * its own (Format > Bullets and numbering…).
 */
function dialogRowPaths(dialog: string | undefined): readonly string[] | undefined {
  return dialog === 'bulletsAndNumbering' ? ICON_PATHS['list_alt'] : undefined;
}

/**
 * A row's label, shortcut, and icon: the entry's own, else the option's for a value row, else
 * the slot's. Null for a slot the registry does not hold. @internal
 */
export function resolveChromeMenuRow(entry: ChromeMenuItemEntry): ResolvedMenuRow | null {
  const control = chromeControlForSlot(entry.slot);
  if (!control) return null;
  const option =
    entry.value === undefined ? undefined : control.options?.find((o) => o.value === entry.value);
  return {
    key: chromeMenuRowKey(entry),
    slot: entry.slot,
    ...(entry.value !== undefined ? { value: entry.value } : {}),
    labelKey: entry.labelKey ?? option?.labelKey ?? control.labelKey,
    ...(!entry.labelKey && option?.labelParams ? { labelParams: option.labelParams } : {}),
    ...(entry.shortcutKey ? { shortcutKey: entry.shortcutKey } : {}),
    paths: option?.paths ?? dialogRowPaths(option?.dialog) ?? control.paths,
    toggle: entry.value === undefined && chromeSlotIsToggle(entry.slot),
  };
}
