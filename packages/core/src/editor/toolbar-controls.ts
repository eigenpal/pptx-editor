// The framework-neutral half of the toolbar controls, shared by the React and Vue adapters so
// the two cannot drift: which part renders a slot, the rows of a value dropdown, the color a
// split button applies first, the size box text, the font list filter, and the More menu fit
// (the room the bar may use, the cost of a separator, and which groups leave the bar). After
// DOCX Editor's `toolbar-measure.ts` and `useToolbarOverflow`.

import type { ChromeActionHost } from './chrome-actions.ts';
import type {
  ChromeControl,
  ChromeControlShape,
  ChromeDialogKey,
  ChromeSlotId,
} from './chrome-controls.ts';
import type { ColorChoice } from '../contracts/index.ts';
import { slotOptionState } from './adapter-support.ts';
import { toolbarCollapse } from './toolbar-layout.ts';

/**
 * Which part draws a control: a registry shape, or an adapter part for a control that loads a
 * file (`load`), opens a panel the adapter owns (`panel`), or runs an adapter action (`action`:
 * a registry dialog, the clipboard, the slide show). @internal
 */
export type ToolbarPartKind = ChromeControlShape | 'load' | 'panel' | 'action';

/** The part kind of `control`, as both adapters' `ToolbarControl` pick it. @internal */
export function toolbarPartKind(control: ChromeControl): ToolbarPartKind {
  const kind = control.state.kind;
  if (kind === 'load' || kind === 'replace') return 'load';
  if (kind === 'panel') return 'panel';
  if (kind === 'dialog' || kind === 'clipboard' || kind === 'slideshow') return 'action';
  // A task pane slot (Format shape, Selection pane) shows its state: not built yet.
  return control.shape ?? 'icon';
}

/** A color a split color button applies, and how its preview shows it. @internal */
export interface ToolbarPickedColor {
  readonly choice: ColorChoice;
  readonly color: string;
}

/**
 * The color a split color button applies before the user picks one: Dark Red text and a
 * Yellow highlight from the Standard Colors, and the theme's Text 1 (`text1`) for outlines.
 * @internal
 */
export function toolbarSplitDefaultColor(
  slot: ChromeSlotId,
  text1: string | undefined
): ToolbarPickedColor {
  if (slot === 'text.highlight') return { choice: '#FFFF00', color: '#FFFF00' };
  if (slot === 'shape.lineColor' && text1) return { choice: { theme: 'tx1' }, color: text1 };
  if (slot === 'shape.lineColor') return { choice: '#000000', color: '#000000' };
  return { choice: '#C00000', color: '#C00000' };
}

/**
 * The translation key of a color menu's reset row: shapes and table cells can be transparent;
 * text returns to its inherited (automatic) color. @internal
 */
export function toolbarColorResetKey(
  slot: ChromeSlotId
): 'toolbar.transparent' | 'toolbar.resetColor' {
  return slot.startsWith('shape.') || slot === 'table.cellFill'
    ? 'toolbar.transparent'
    : 'toolbar.resetColor';
}

/**
 * The tab a dialog row of a slot opens on: More Spacing… opens the Font dialog on Character
 * Spacing; Bullets and Numbering… opens on the tab of the gallery it comes from.
 */
const ROW_SECTIONS: Readonly<Partial<Record<string, string>>> = {
  'text.characterSpacing': 'spacing',
  'list.bulletStyle': 'bullet',
  'list.numberStyle': 'number',
};

/** The dialog section a dialog row of `slot` opens on, if any. @internal */
export function toolbarRowSection(slot: ChromeSlotId): string | undefined {
  return ROW_SECTIONS[slot];
}

/** One option row of a value dropdown, as both adapters draw it. @internal */
export interface ToolbarDropdownRow {
  /** The registry dialog the row opens, for a dialog row ("Line spacing options…"). */
  readonly dialog: ChromeDialogKey | undefined;
  /** Whether the row shows as checked: `undefined` for a row that is not a radio. */
  readonly checked: boolean | undefined;
  /** Whether the row is off: the engine or the host refuses it. */
  readonly refused: boolean;
  /** Why, as a translation key or the engine's text, when the row is off. */
  readonly reason: string | null;
}

/**
 * The state of one dropdown option. A dialog row opens its registry dialog; with the dialog
 * turned off (`popups.<key>: false`) it stays, disabled with the reason, as in the menus. An
 * option the engine would refuse is disabled, with the reason; a `choice` option with no
 * command of its own stays on. @internal
 */
export function toolbarDropdownRow(
  editor: Parameters<typeof slotOptionState>[0],
  host: Pick<ChromeActionHost, 'unavailable'> | null,
  slot: ChromeSlotId,
  control: ChromeControl,
  option: { readonly value: unknown; readonly dialog?: ChromeDialogKey | undefined },
  current: unknown
): ToolbarDropdownRow {
  // Only a `choice` dropdown is a radio group; the others run an action per option.
  const choice = control.toggle === 'choice';
  const dialog = option.dialog;
  const off = dialog ? (host?.unavailable?.(slot, dialog) ?? null) : null;
  const state = dialog
    ? { enabled: off === null, reason: off }
    : slotOptionState(editor, slot, option.value);
  const refused = !state.enabled && (!choice || state.reason !== 'disabled.notWired');
  return {
    dialog,
    // A number value (the line spacing `1.5`) matches its option's text (`'1.5'`).
    checked:
      choice && !dialog
        ? current !== undefined && current !== null && String(option.value) === String(current)
        : undefined,
    refused,
    reason: refused ? state.reason : null,
  };
}

/** The families of the font list that match a search `query`, case-insensitive. @internal */
export function filterFontFamilies(all: readonly string[], query: string): readonly string[] {
  const q = query.trim().toLowerCase();
  return q ? all.filter((f) => f.toLowerCase().includes(q)) : all;
}

/** The text of the font size box: blank over a mixed size, as the Font dialog's field. @internal */
export function toolbarSizeText(value: unknown, mixed: boolean): string {
  return mixed || value === undefined ? '' : String(Math.round(Number(value) * 10) / 10);
}

/** The sizes the font size box takes, in points: the range `setFontSize` accepts. @internal */
export const TOOLBAR_SIZE_RANGE = { min: 1, max: 4000 } as const;

/**
 * Why the font size box refuses its text, as a message key with the range, or null when the
 * text is a size in `TOOLBAR_SIZE_RANGE` or blank (no change). The box says why instead of
 * dropping the size, as the Font dialog's size field does. @internal
 */
export function toolbarSizeRefusal(draft: string | null): {
  readonly key: 'toolbar.fontSizeRange';
  readonly params: { readonly min: number; readonly max: number };
} | null {
  if (draft === null || !draft.trim()) return null;
  const n = Number(draft);
  const { min, max } = TOOLBAR_SIZE_RANGE;
  return Number.isFinite(n) && n >= min && n <= max
    ? null
    : { key: 'toolbar.fontSizeRange', params: { min, max } };
}

/** How long the font size box shows why it refused a size, in milliseconds. @internal */
export const SIZE_ERROR_MS = 5000;

/** Where the font size box's message goes: under the box, in viewport pixels. @internal */
export function sizeErrorPlace(box: HTMLElement | null): { left: number; top: number } {
  const rect = box?.getBoundingClientRect();
  return { left: Math.round(rect?.left ?? 0), top: Math.round((rect?.bottom ?? 0) + 6) };
}

/** The size typed in the font size box, or null when it is not a positive number. @internal */
export function parseToolbarSize(draft: string | null): number | null {
  if (draft === null) return null;
  const n = Number(draft);
  return Number.isFinite(n) && n > 0 ? n : null;
}

/**
 * The room the bar may use: its own content width, or more when its parent has room, after
 * DOCX Editor's `readAvailableWidth`. The packaged bar is `width: max-content`, so after a
 * collapse its own width is only what it shows; the parent's room lets groups come back when
 * the window grows. The parent counts only where it decides the bar's width the same way
 * (a block parent, or a flex row whose growing siblings count at their minimum width): an
 * out-of-flow or floated bar, and a bar with a fixed `max-width` keep their own box, and a
 * grid cell gives its own width. The caller also caps the room after the controls overflowed
 * it. @internal
 */
export function toolbarRoom(bar: HTMLElement, style: CSSStyleDeclaration): number {
  const px = (value: string): number => parseFloat(value) || 0;
  const padding = px(style.paddingLeft) + px(style.paddingRight);
  const own = bar.clientWidth - padding;
  const parent = bar.parentElement;
  if (!parent || style.position === 'absolute' || style.position === 'fixed') return own;
  if ((style.cssFloat || 'none') !== 'none') return own;
  const view = bar.ownerDocument.defaultView;
  if (!view) return own;
  const ps = view.getComputedStyle(parent);
  // A grid or table cell sizes the bar from its own track: a full width there is the cell's
  // room, read with a moment's `width: 100%` (no paint happens in between).
  if (ps.display.includes('grid') || ps.display === 'table-cell') {
    const width = bar.style.width;
    bar.style.width = '100%';
    const cell = bar.clientWidth - padding;
    bar.style.width = width;
    return cell;
  }
  const flexRow = ps.display.includes('flex') && !ps.flexDirection.startsWith('column');
  const block = ps.display === 'block' || ps.display === 'flow-root';
  if (!flexRow && !block) return own;
  if (flexRow && px(style.flexGrow) > 0) return own;
  let siblings = 0;
  if (flexRow) {
    const others = [...parent.children].filter((c) => c !== bar) as HTMLElement[];
    for (const other of others) {
      const os = view.getComputedStyle(other);
      if (os.position === 'absolute' || os.position === 'fixed' || os.display === 'none') continue;
      // A growing sibling (a `flex: 1` spacer) gives its room back: count its minimum.
      siblings +=
        (px(os.flexGrow) > 0 ? px(os.minWidth) : other.offsetWidth) +
        px(os.marginLeft) +
        px(os.marginRight) +
        px(ps.columnGap);
    }
  }
  const parentContent = parent.clientWidth - px(ps.paddingLeft) - px(ps.paddingRight);
  const chrome =
    Math.max(0, bar.offsetWidth - bar.clientWidth) + px(style.marginLeft) + px(style.marginRight);
  let room = parentContent - siblings - chrome - padding;
  // A fixed `max-width` (a host class) bounds the bar; a percentage resolves against the parent.
  const max = style.maxWidth;
  if (max.endsWith('px')) room = Math.min(room, px(max) - padding - chrome);
  return Math.max(own, room);
}

/** What one separator costs in the bar: its width, its margins, and the gaps beside it. @internal */
export function separatorCost(bar: HTMLElement, style: CSSStyleDeclaration): number {
  const sep = bar.querySelector<HTMLElement>(':scope > .pptx-toolbar__sep');
  const gap = parseFloat(style.columnGap) || 0;
  const view = bar.ownerDocument.defaultView;
  if (!sep || !view) return 15;
  const s = view.getComputedStyle(sep);
  return (
    sep.offsetWidth + (parseFloat(s.marginLeft) || 0) + (parseFloat(s.marginRight) || 0) + 2 * gap
  );
}

/** One group of a bar as the fit sees it. @internal */
export interface ToolbarFitGroup {
  readonly id: string;
  readonly priority: number;
  readonly pinned: boolean;
}

/**
 * What a bar keeps between fits: each group's measured width by id (a group that comes or goes
 * does not hand its width to its neighbour), and the room cap after the controls overflowed
 * the bar's own box. @internal
 */
export interface ToolbarFitMemory {
  readonly widths: Map<string, number>;
  cap: { parent: number; room: number } | null;
}

/** A new, empty fit memory for one bar. @internal */
export function createToolbarFitMemory(): ToolbarFitMemory {
  return { widths: new Map(), cap: null };
}

/**
 * The ids of the groups that leave `bar` for the More menu, in bar order. Measured, not
 * assumed: each group after the first costs a separator and two gaps, and the More button
 * costs one of each plus its own width (`more`, else the packaged button width). `fixed` is the
 * block of loose host children, which never collapses. The groups on the bar carry
 * `data-group-index` and `data-group`. @internal
 */
export function toolbarCollapsedIds(
  bar: HTMLElement,
  groups: readonly ToolbarFitGroup[],
  memory: ToolbarFitMemory,
  more: HTMLElement | null,
  fixed: HTMLElement | null
): string[] {
  const view = bar.ownerDocument.defaultView;
  if (!view) return [];
  const style = view.getComputedStyle(bar);
  const sep = separatorCost(bar, style);
  for (const g of bar.querySelectorAll<HTMLElement>('[data-group-index]'))
    if (g.dataset.group) memory.widths.set(g.dataset.group, g.offsetWidth + sep);
  const button = more?.offsetWidth || 30;
  // After the controls overflowed the bar's own box (a host width, a grid cell), trust that
  // box until the parent's width changes, as DOCX Editor caps its room.
  const own =
    bar.clientWidth - (parseFloat(style.paddingLeft) || 0) - (parseFloat(style.paddingRight) || 0);
  const parentWidth = bar.parentElement?.clientWidth ?? 0;
  if (memory.cap && memory.cap.parent !== parentWidth) memory.cap = null;
  if (bar.scrollWidth > bar.clientWidth + 1) memory.cap = { parent: parentWidth, room: own };
  const inner =
    Math.min(toolbarRoom(bar, style), memory.cap?.room ?? Number.POSITIVE_INFINITY) -
    (fixed ? fixed.offsetWidth + sep : 0) +
    // The first group has no separator before it.
    sep;
  return [
    ...toolbarCollapse(
      groups.map((g) => ({
        width: memory.widths.get(g.id) ?? 0,
        priority: g.priority,
        pinned: g.pinned,
      })),
      inner,
      button + sep
    ),
  ]
    .sort((a, b) => a - b)
    .map((i) => groups[i]!.id);
}

/**
 * Run `fit` when the bar, its parent, a group, or the fixed block changes size: the bar is
 * only as wide as its content, so a wider window shows on its parent; a group or the fixed
 * block that grows (a host action added) shows on itself. Returns the cleanup. @internal
 */
export function observeToolbarFit(bar: HTMLElement, fit: () => void): () => void {
  const Observer = bar.ownerDocument.defaultView?.ResizeObserver ?? globalThis.ResizeObserver;
  if (typeof Observer !== 'function') return () => undefined;
  const observer = new Observer(fit);
  observer.observe(bar);
  if (bar.parentElement) observer.observe(bar.parentElement);
  for (const part of bar.querySelectorAll('[data-group-index], [data-fixed]'))
    observer.observe(part);
  return () => observer.disconnect();
}
