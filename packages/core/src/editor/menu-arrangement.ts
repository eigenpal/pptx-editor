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
import { tidySeparators } from './chrome-forms.ts';
import { warnMenuPart } from './dev-warnings.ts';
import { translateKeyOrText } from './adapter-support.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';

// ---------------------------------------------------------------------------------------------
// Context menu placement

/** The gap a context menu keeps from the window edges, in pixels. */
const EDGE = 8;

/**
 * Place a panel of size `w` × `h` at the point `at`, flipped and clamped to stay inside a
 * window of size `vw` × `vh`. @internal
 */
export function placeInViewport(
  at: { readonly x: number; readonly y: number },
  w: number,
  h: number,
  vw: number,
  vh: number
): { left: number; top: number } {
  const left = at.x + w + EDGE > vw ? Math.max(EDGE, Math.min(at.x - w, vw - w - EDGE)) : at.x;
  const top = at.y + h + EDGE > vh ? Math.max(EDGE, Math.min(at.y - h, vh - h - EDGE)) : at.y;
  return { left, top };
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
  | 'contextMenu.canvasLabel' {
  if (kind === 'table') return 'contextMenu.tableLabel';
  return kind === 'text'
    ? 'contextMenu.textLabel'
    : kind === 'shapes'
      ? 'contextMenu.shapeLabel'
      : kind === 'picture'
        ? 'contextMenu.pictureLabel'
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
