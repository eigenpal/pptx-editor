// Framework-free rules of the menu bar and the context menus, shared by the React and Vue
// adapters (`menu-parts`, `context-menu-parts`, `ContextMenu`, `CanvasContextMenu`,
// `RailContextMenu`), so the two cannot drift: which registry rows a host row replaces or
// hides, which menus the bar shows in which order, how the bar's arrow keys move, where a
// context menu opens, and how the slide and section menu rows read the engine.

import type { Translate } from '@pptx-editor.dev/i18n';
import type { EditorCommand, SlideLayoutSummary } from '../contracts/index.ts';
import {
  CHROME_MENUS,
  chromeMenuRowKey,
  type CanvasMenuKind,
  type ChromeMenu,
  type ChromeMenuEntry,
} from './chrome-menus.ts';
import type { ChromeSlotId } from './chrome-controls.ts';
import { tidySeparators } from './chrome-forms.ts';
import { warnMenuPart } from './dev-warnings.ts';
import { translateKeyOrText } from './adapter-support.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';

// ---------------------------------------------------------------------------------------------
// Context menu placement

/** The gap a context menu keeps from the window edges, in pixels. */
const EDGE = 8;

/**
 * Place a panel of size `w` × `h` at the point `at`, inside a window of size `vw` × `vh`. The
 * panel keeps a corner at the point: it opens below and to the right, flips above or to the
 * left when it does not fit, and is clamped at the side edges. When it fits neither above nor
 * below, it opens on the side with more room and `maxHeight` caps it there, so its rows scroll
 * and it never slides away from the point (over the toolbar, for a menu opened in a rail at the
 * bottom of a short window). @internal
 */
export function placeInViewport(
  at: { readonly x: number; readonly y: number },
  w: number,
  h: number,
  vw: number,
  vh: number
): { left: number; top: number; maxHeight?: number } {
  const left = at.x + w + EDGE > vw ? Math.max(EDGE, Math.min(at.x - w, vw - w - EDGE)) : at.x;
  const below = vh - EDGE - at.y;
  const above = at.y - EDGE;
  if (h <= below) return { left, top: at.y };
  if (h <= above) return { left, top: at.y - h };
  return below >= above
    ? { left, top: at.y, maxHeight: Math.max(0, below) }
    : { left, top: EDGE, maxHeight: above };
}

/**
 * Where a canvas context menu opens: the pointer, or for a keyboard open (`at` null) the middle
 * of the selection outline, else of the canvas. @internal
 */
export function canvasMenuPoint(
  owner: Element,
  at: { readonly x: number; readonly y: number } | null
): { x: number; y: number } {
  if (at) return { x: at.x, y: at.y };
  const outline = owner.querySelector('.pptx-surface__outline')?.getBoundingClientRect();
  const box = outline ?? owner.getBoundingClientRect();
  return { x: Math.round(box.left + box.width / 2), y: Math.round(box.top + box.height / 2) };
}

/** The translation key of a canvas context menu's accessible name. @internal */
export function canvasMenuLabelKey(
  kind: CanvasMenuKind
):
  | 'contextMenu.textLabel'
  | 'contextMenu.tableLabel'
  | 'contextMenu.shapeLabel'
  | 'contextMenu.pictureLabel'
  | 'contextMenu.chartLabel'
  | 'contextMenu.diagramLabel'
  | 'contextMenu.canvasLabel' {
  if (kind === 'table' || kind === 'tableFrame') return 'contextMenu.tableLabel';
  return kind === 'text'
    ? 'contextMenu.textLabel'
    : kind === 'shapes'
      ? 'contextMenu.shapeLabel'
      : kind === 'picture'
        ? 'contextMenu.pictureLabel'
        : kind === 'chart'
          ? 'contextMenu.chartLabel'
          : kind === 'diagram'
            ? 'contextMenu.diagramLabel'
            : 'contextMenu.canvasLabel';
}

/**
 * Whether a key opens the canvas context menu: the Menu key, or Shift+F10. @internal
 */
export function isContextMenuKey(e: { readonly key: string; readonly shiftKey: boolean }): boolean {
  return e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10');
}

// ---------------------------------------------------------------------------------------------
// Host rows over registry rows

/** The registry row a named menu row part stands for: its slot, and its value if any. @internal */
export interface MenuNamedRow {
  readonly slot: ChromeSlotId;
  readonly value?: string;
}

const row = (slot: ChromeSlotId, value?: string): MenuNamedRow =>
  value === undefined ? { slot } : { slot, value };

/** The order and grouping rows both menus share. */
const ARRANGE_ROWS = {
  BringToFront: row('arrange.order', 'front'),
  SendToBack: row('arrange.order', 'back'),
  BringForward: row('arrange.order', 'forward'),
  SendBackward: row('arrange.order', 'backward'),
  GroupObjects: row('arrange.group'),
  UngroupObjects: row('arrange.ungroup'),
} as const;

/**
 * The named rows of the canvas context menu (`ContextMenu.Cut`, …) and the row each one stands
 * for, after DOCX Editor's context menu parts. `ContextMenu.Cut` is shorthand for
 * `ContextMenu.Slot slotId="edit.cut"`, with the same props. The names follow DOCX Editor's
 * where the row is the same (`Cut`, `Copy`, `Paste`, `Delete`, `InsertRowAbove`, `MergeCells`,
 * `DeleteTable`), and the toolbar's named parts elsewhere. @internal
 */
export const CONTEXT_MENU_NAMED_PARTS = {
  Cut: row('edit.cut'),
  Copy: row('edit.copy'),
  Paste: row('edit.paste'),
  Delete: row('more.delete'),
  Duplicate: row('more.duplicate'),
  Bold: row('text.bold'),
  Italic: row('text.italic'),
  Underline: row('text.underline'),
  Link: row('text.link'),
  ClearFormatting: row('format.clear'),
  FormatPainter: row('format.painter'),
  Font: row('text.dialog'),
  ParagraphDialog: row('paragraph.dialog'),
  ...ARRANGE_ROWS,
  ImageAltText: row('image.altText'),
  ImageChange: row('image.change'),
  ImageProperties: row('image.properties'),
  ChartEditData: row('chart.editData'),
  DiagramAddAfter: row('diagram.addAfter'),
  DiagramAddBefore: row('diagram.addBefore'),
  DiagramAddBelow: row('diagram.addBelow'),
  DiagramDelete: row('diagram.delete'),
  FormatBackground: row('slide.background'),
  NewSlide: row('slide.add'),
  InsertRowAbove: row('table.insertRowAbove'),
  InsertRowBelow: row('table.insertRowBelow'),
  InsertColumnLeft: row('table.insertColumnLeft'),
  InsertColumnRight: row('table.insertColumnRight'),
  MergeCells: row('table.mergeCells'),
  DeleteTableRow: row('table.delete', 'rows'),
  DeleteTableColumn: row('table.delete', 'columns'),
  DeleteTable: row('table.delete', 'table'),
  SelectTable: row('table.select', 'table'),
  SelectTableRow: row('table.select', 'row'),
  SelectTableColumn: row('table.select', 'column'),
} as const satisfies Record<string, MenuNamedRow>;

/**
 * The named rows of the menu bar (`Menu.PageSetup`, `Menu.Cut`, …) and the row each one stands
 * for, after DOCX Editor's menu row parts (`PageSetup`, `ParagraphDialog`, `ImageInsert`). Put
 * one in the menu that holds its row, such as `Menu.Edit`. @internal
 */
export const MENU_NAMED_PARTS = {
  PageSetup: row('file.pageSetup'),
  Undo: row('history.undo'),
  Redo: row('history.redo'),
  Cut: row('edit.cut'),
  Copy: row('edit.copy'),
  Paste: row('edit.paste'),
  SelectAll: row('edit.selectAll'),
  Duplicate: row('more.duplicate'),
  Delete: row('more.delete'),
  Find: row('find.find'),
  Replace: row('find.replace'),
  NormalView: row('view.presentationView', 'normal'),
  SlideSorterView: row('view.presentationView', 'slideSorter'),
  SmartGuides: row('view.smartGuides'),
  NewSlide: row('slide.add'),
  TextBox: row('insert.textBox'),
  TableInsert: row('table.insert'),
  ImageInsert: row('insert.image'),
  Shapes: row('insert.shape'),
  Line: row('insert.line'),
  Link: row('text.link'),
  Font: row('text.dialog'),
  ParagraphDialog: row('paragraph.dialog'),
  BulletsAndNumbering: row('list.bulletStyle', 'dialog'),
  Columns: row('paragraph.columns', 'dialog'),
  Bold: row('text.bold'),
  Italic: row('text.italic'),
  Underline: row('text.underline'),
  Strike: row('text.strike'),
  BulletList: row('list.bullet'),
  NumberedList: row('list.numbered'),
  FormatPainter: row('format.painter'),
  ClearFormatting: row('format.clear'),
  FormatBackground: row('slide.background'),
  ImageCrop: row('image.crop'),
  ImageProperties: row('image.properties'),
  ...ARRANGE_ROWS,
  PlayFromStart: row('view.playFromStart'),
  PlayFromCurrent: row('view.playFromCurrent'),
  HideSlide: row('slide.hide'),
} as const satisfies Record<string, MenuNamedRow>;

/** The name of a named context menu row. @internal */
export type ContextMenuPartName = keyof typeof CONTEXT_MENU_NAMED_PARTS;
/** The name of a named menu bar row. @internal */
export type MenuRowPartName = keyof typeof MENU_NAMED_PARTS;

/** A host row that names a registry slot (`Menu.Item`, `ContextMenu.Slot`). @internal */
export interface MenuSlotRowSpec {
  readonly slotId: string;
  readonly value?: string | undefined;
  readonly hidden?: boolean | undefined;
}

/** The rows of a menu after the host's children. @internal */
export interface ArrangedMenuRows<C> {
  /** The registry rows still shown, with the separators a hidden row left behind removed. */
  readonly entries: ChromeMenuEntry[];
  /** Host rows that replace registry rows, by row key; `null` removes the row. */
  readonly overrides: Map<string, C | null>;
  /** Host rows that replace no registry row: they follow the registry rows. */
  readonly extras: C[];
}

/**
 * Split a menu's host children against its registry rows: a slot row (`slotOf` answers for
 * it) whose slot and value match a registry row replaces that row in place, or removes it when
 * `hidden`; other children are extra rows. A hidden slot row that matches nothing calls
 * `onUnmatchedHidden` and adds no row. @internal
 */
export function arrangeMenuRows<C>(
  entries: readonly ChromeMenuEntry[],
  children: readonly C[],
  slotOf: (child: C) => MenuSlotRowSpec | null,
  onUnmatchedHidden?: (slotId: string) => void
): ArrangedMenuRows<C> {
  const keys = new Set<string>();
  const collect = (list: readonly ChromeMenuEntry[]): void => {
    for (const e of list)
      if (e.kind === 'item') keys.add(chromeMenuRowKey(e));
      else if (e.kind === 'submenu') collect(e.items);
  };
  collect(entries);
  const overrides = new Map<string, C | null>();
  const extras: C[] = [];
  for (const child of children) {
    const spec = slotOf(child);
    if (spec) {
      const key = chromeMenuRowKey({
        slot: spec.slotId as never,
        ...(spec.value !== undefined ? { value: spec.value } : {}),
      });
      if (keys.has(key)) {
        overrides.set(key, spec.hidden ? null : child);
        continue;
      }
      if (spec.hidden) {
        onUnmatchedHidden?.(spec.slotId);
        continue;
      }
    }
    extras.push(child);
  }
  // A hidden row takes its separators with it when they would stand alone.
  const shown = tidySeparators(
    entries.filter((e) => e.kind !== 'item' || overrides.get(chromeMenuRowKey(e)) !== null)
  );
  return { entries: shown, overrides, extras };
}

// ---------------------------------------------------------------------------------------------
// The menu bar

/** What the bar reads from a named menu or `Menu.Entry` child. @internal */
export interface MenuBarSpec {
  readonly id: string;
  readonly label?: string | undefined;
  readonly labelKey?: string | undefined;
  readonly hidden?: boolean | undefined;
  readonly after?: string | undefined;
}

/** One menu of the bar, before the adapter renders its rows. @internal */
export interface ArrangedBarMenu<S extends MenuBarSpec> {
  readonly id: string;
  readonly title: string;
  /** The host's description of the menu, if any. */
  readonly spec: S | undefined;
  /** The registry menu, for a built-in menu: its rows come from `arrangeMenuRows`. */
  readonly registry: ChromeMenu | undefined;
}

/**
 * The menus of the bar, in order: with `preset`, the registry menus (minus hidden ones),
 * changed in place by the host's specs; then the host's other menus, each after its `after`
 * menu or at the end. With `preset` off, a named built-in menu still brings its title and rows.
 * @internal
 */
export function arrangeMenuBar<S extends MenuBarSpec>(
  specs: ReadonlyMap<string, S>,
  preset: boolean,
  t: Translate
): ArrangedBarMenu<S>[] {
  const titleOf = (spec: S | undefined, fallback: string): string =>
    spec?.label ?? (spec?.labelKey ? translateKeyOrText(t, spec.labelKey) : fallback);
  const menus: ArrangedBarMenu<S>[] = [];
  if (preset)
    for (const menu of CHROME_MENUS) {
      const spec = specs.get(menu.id);
      if (spec?.hidden) continue;
      menus.push({ id: menu.id, title: titleOf(spec, t(menu.labelKey)), spec, registry: menu });
    }
  const builtIn = new Set<string>(CHROME_MENUS.map((m) => m.id));
  for (const spec of specs.values()) {
    if ((preset && builtIn.has(spec.id)) || spec.hidden) continue;
    const registry = CHROME_MENUS.find((m) => m.id === spec.id);
    const entry: ArrangedBarMenu<S> = {
      id: spec.id,
      title: titleOf(spec, registry ? t(registry.labelKey) : spec.id),
      spec,
      registry,
    };
    const at = spec.after === undefined ? -1 : menus.findIndex((m) => m.id === spec.after);
    if (spec.after !== undefined && at === -1)
      warnMenuPart('Menu.Entry', 'unknownAnchor', spec.after);
    if (at === -1) menus.push(entry);
    else menus.splice(at + 1, 0, entry);
  }
  return menus;
}

/** The part name of a built-in menu, for warnings (`Menu.SlideShow`). @internal */
export function menuPartName(id: string): string {
  return id === 'slideShow' ? 'Menu.SlideShow' : `Menu.${id[0]?.toUpperCase() ?? ''}${id.slice(1)}`;
}

/** The step ArrowLeft or ArrowRight takes along the bar (reversed right to left), or 0. @internal */
export function menuBarArrow(key: string, rtl: boolean): -1 | 0 | 1 {
  const by = key === 'ArrowRight' ? 1 : key === 'ArrowLeft' ? -1 : 0;
  return (rtl ? -by : by) as -1 | 0 | 1;
}

/** What a key on a bar trigger does. @internal */
export type MenuBarTriggerKey =
  | { readonly kind: 'step'; readonly by: -1 | 1 }
  | { readonly kind: 'open' }
  | { readonly kind: 'edge'; readonly edge: 'first' | 'last' }
  | null;

/**
 * A key on a menu bar trigger, as DOCX Editor's menu bar: the arrows along the bar step,
 * ArrowDown, ArrowUp, Enter, and Space open the menu, Home and End go to the ends. @internal
 */
export function menuBarTriggerKey(key: string, rtl: boolean): MenuBarTriggerKey {
  const by = menuBarArrow(key, rtl);
  if (by) return { kind: 'step', by };
  if (['ArrowDown', 'ArrowUp', 'Enter', ' '].includes(key)) return { kind: 'open' };
  if (key === 'Home' || key === 'End')
    return { kind: 'edge', edge: key === 'Home' ? 'first' : 'last' };
  return null;
}

/**
 * Whether closing the bar gives focus back to `target`: only while focus is nowhere else on
 * purpose (a row that opened a dialog, the find bar, or the link popup keeps it). @internal
 */
export function menuBarRestoresFocus(target: HTMLElement, bar: Element | null): boolean {
  if (!target.isConnected) return false;
  const active = target.ownerDocument.activeElement;
  return (
    !active ||
    active === target.ownerDocument.body ||
    !!bar?.contains(active) ||
    active.closest('.pptx-menu') !== null
  );
}

/**
 * Whether a menu bar panel just closed on Escape: the popover gave focus to its trigger before
 * it called `onClose`, and only Escape does that. The WAI-ARIA menubar pattern, as DOCX
 * Editor's menu bar: Escape closes the menu and leaves focus on its menubar item, also after a
 * pointer open, so the bar does not give focus back to the slide. @internal
 */
export function menuBarClosedOnEscape(trigger: HTMLElement | null | undefined): boolean {
  return !!trigger?.isConnected && trigger.ownerDocument.activeElement === trigger;
}

// ---------------------------------------------------------------------------------------------
// Slide and section menu rows

/**
 * Why a slide or section menu row cannot run `command`, as a translation key or the engine's
 * reason, or null when it can. With no editor, `disabled.noEditor`. @internal
 */
export function commandRowReason(
  editor: PptxEditorInstance | null,
  command: EditorCommand | null
): string | null {
  const can = command && editor ? editor.can(command) : null;
  return !can ? 'disabled.noEditor' : can.ok ? null : can.reason;
}

/**
 * The command that tells whether the slide menu's Layout row can pick any layout: the slide's
 * own layout when the deck lists it, else the first one, else a reset. @internal
 */
export function layoutRowProbe(
  layouts: readonly SlideLayoutSummary[],
  currentLayoutId: string | undefined,
  selected: readonly number[]
): EditorCommand {
  const listed = layouts.some((l) => l.id === currentLayoutId) ? currentLayoutId : layouts[0]?.id;
  return listed
    ? { type: 'applyLayout', indices: selected, layoutId: listed }
    : { type: 'resetSlides', indices: selected };
}

/** The slide menu's Add Section command at slide `index`. @internal */
export function addSectionCommand(index: number, t: Translate): EditorCommand {
  return {
    type: 'addSection',
    index,
    name: t('sections.untitled'),
    defaultSectionName: t('sections.default'),
  };
}

/**
 * Run Add Section at slide `index` and return the id of the section it added there, for the
 * rail to start renaming it; null when the command was refused. @internal
 */
export function addSectionAt(
  editor: PptxEditorInstance,
  command: EditorCommand,
  index: number
): string | null {
  const before = new Set((editor.snapshot().sections ?? []).map((s) => s.id));
  if (!editor.exec(command).ok) return null;
  const added = (editor.snapshot().sections ?? []).filter((s) => !before.has(s.id));
  return (added.find((s) => s.firstSlide === index) ?? added[0])?.id ?? null;
}
