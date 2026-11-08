// Framework-free rules of the shared chrome parts, used by the React and Vue adapters (moved from
// packages/react/src/ColorSwatches.tsx, dialog-color.tsx, dialog-form.tsx, menu-rows.tsx, and
// notes-target.ts): the swatch grid's keys and selection, the color field's labels, the form
// dialog's error marks and tab keys, the menu row's check state, and the focused notes field.
// The adapters keep only their glue, so the two cannot drift.

import type { Translate } from '@pptx-editor.dev/i18n';
import type { ColorChoice, NotesFormat, ThemeColorSlot, ThemeSwatch } from '../contracts/index.ts';
import type { ChromeCommandState } from './adapter-support.ts';
import type { ChromeControl, ChromeSlotId } from './chrome-controls.ts';
import type { ChromeMenuEntry, ResolvedMenuRow } from './chrome-menus.ts';
import { chromeOptionChecked } from './chrome-shapes.ts';
import { colorChoiceOf, type ColorPick } from './color-fields.ts';
import { STANDARD_COLORS, swatchName } from './palette.ts';

// --- Swatch grids -------------------------------------------------------------------------

/** The swatch that arrow key `key` moves to from `at`, in a grid of `columns`, or -1. @internal */
export function swatchStep(key: string, at: number, count: number, columns: number): number {
  const to =
    key === 'ArrowRight'
      ? at + 1
      : key === 'ArrowLeft'
        ? at - 1
        : key === 'ArrowDown'
          ? at + columns
          : key === 'ArrowUp'
            ? at - columns
            : key === 'Home'
              ? 0
              : key === 'End'
                ? count - 1
                : -1;
  return to >= 0 && to < count ? to : -1;
}

const gridSwatches = (grid: HTMLElement | null): HTMLButtonElement[] => [
  ...(grid?.querySelectorAll<HTMLButtonElement>(':scope > .pptx-swatch') ?? []),
];

/**
 * Keep one Tab stop in a roving swatch grid: the focused swatch, else the chosen one, else the
 * first. Run after each render. @internal
 */
export function syncSwatchTabStops(grid: HTMLElement | null): void {
  const all = gridSwatches(grid);
  const stop =
    all.find((b) => b === b.ownerDocument.activeElement) ??
    all.find((b) => b.getAttribute('aria-pressed') === 'true') ??
    all[0];
  for (const b of all) b.tabIndex = b === stop ? 0 : -1;
}

/** `#RRGGBB` as the `rgb(r, g, b)` text that a browser gives back for an inline color. */
function rgbText(hex: string): string | null {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  return m
    ? `rgb(${m
        .slice(1)
        .map((h) => parseInt(h, 16))
        .join(', ')})`
    : null;
}

/**
 * Focus the chosen swatch of an open color grid, else the first swatch of `shown` (the color
 * that the field's button shows, such as an automatic color), else its first button, and scroll
 * the panel into view. The panel can hold more than one grid. @internal
 */
export function focusCurrentSwatch(panel: HTMLElement | null, shown: string | null = null): void {
  const rgb = shown ? rgbText(shown) : null;
  const buttons = [...(panel?.querySelectorAll<HTMLElement>('button') ?? [])];
  const to =
    panel?.querySelector<HTMLElement>('button[aria-pressed="true"]') ??
    (rgb ? buttons.find((b) => b.style.background === rgb) : undefined) ??
    buttons[0];
  to?.focus({ preventScroll: true });
  // The whole panel, with its custom color field, scrolls into view: the grid can open at the
  // bottom of a side panel.
  panel?.scrollIntoView?.({ block: 'nearest' });
}

/**
 * After a row of the Chart tab is added or removed, move the focus to row `index` of `list`
 * (`data-chart-row`), else to its last row: its first field (the summary of a closed card),
 * with the text selected for a new row. `from` is the tab or an element in it; the focus moves
 * after the tab renders again, so a keyboard user keeps their place. @internal
 */
export function focusChartRow(
  from: Element | null,
  list: 'category' | 'series',
  index: number,
  select = false
): void {
  const pane = from?.closest('.pptx-chart-pane');
  const view = pane?.ownerDocument.defaultView;
  if (!pane || !view) return;
  view.requestAnimationFrame(() => {
    const rows = pane.querySelectorAll(`[data-chart-row="${list}"]`);
    const row = rows[Math.min(index, rows.length - 1)];
    // A closed series card shows its summary only.
    const closed = row instanceof view.HTMLDetailsElement && !row.open;
    const field = row?.querySelector<HTMLElement>(closed ? 'summary' : 'input:not([disabled])');
    field?.focus();
    if (select && field instanceof view.HTMLInputElement) field.select();
  });
}

/**
 * The arrow keys, Home, and End in a roving swatch grid, as a WAI-ARIA grid of buttons: they
 * move by column and row, and the Tab stop moves with the focus. @internal
 */
export function swatchGridKeyDown(
  grid: HTMLElement | null,
  event: { readonly key: string; readonly target: EventTarget | null; preventDefault(): void },
  columns: number
): void {
  const all = gridSwatches(grid);
  const at = all.indexOf(event.target as HTMLButtonElement);
  const to = swatchStep(event.key, at, all.length, columns);
  if (to === -1 || at === -1) return;
  event.preventDefault();
  for (const b of all) b.tabIndex = b === all[to] ? 0 : -1;
  all[to]!.focus();
}

/**
 * Which swatch shows as chosen. A theme choice matches its slot and shift. A hex choice marks
 * one swatch only: the standard swatch of that color, else the first theme swatch of it.
 * @internal
 */
export function swatchSelection(
  selected: ColorChoice | null | undefined,
  rows: readonly (readonly ThemeSwatch[])[]
): (color: string, slot?: string, shift?: number) => boolean {
  const hex = typeof selected === 'string' ? selected.toUpperCase() : null;
  const hexIsStandard = hex !== null && STANDARD_COLORS.some((s) => s.color.toUpperCase() === hex);
  const firstTheme =
    hex === null || hexIsStandard
      ? null
      : (rows.flat().find((s) => s.color.toUpperCase() === hex) ?? null);
  return (color, slot, shift) => {
    if (hex !== null)
      return slot === undefined
        ? color.toUpperCase() === hex
        : firstTheme !== null &&
            firstTheme.slot === slot &&
            (firstTheme.shift ?? 0) === (shift ?? 0);
    return (
      !!selected &&
      typeof selected !== 'string' &&
      selected.theme === slot &&
      (selected.shift ?? 0) === (shift ?? 0)
    );
  };
}

/** The color choice a theme swatch picks: its slot, with its shift when it has one. @internal */
export function themeSwatchChoice(swatch: ThemeSwatch): ColorChoice {
  return swatch.shift ? { theme: swatch.slot, shift: swatch.shift } : { theme: swatch.slot };
}

/** The ten theme colors (the grid's first row) with their names, for gradient presets. @internal */
export function themeBaseSwatches(
  t: Translate,
  first: readonly ThemeSwatch[] | undefined
): readonly { color: string; name: string }[] {
  return (first ?? []).map((s) => ({ color: s.color, name: swatchName(t, s) }));
}

// --- The color field ----------------------------------------------------------------------

/** The label of a color pick in the color field's list. @internal */
export function colorPickLabel(t: Translate, pick: ColorPick): string {
  if (pick === 'auto') return t('colors.automatic');
  if (pick === 'keep') return t('colors.keep');
  if (pick === 'mixed') return t('colors.mixed');
  if (pick.startsWith('#')) {
    const standard = STANDARD_COLORS.find((s) => s.color.toUpperCase() === pick.toUpperCase());
    return standard ? t(`colors.standard.${standard.name}`) : pick;
  }
  const [slot, shift] = pick.split(':') as [ThemeColorSlot, string?];
  return swatchName(t, { slot, shift: Number(shift ?? 0), color: '' });
}

/**
 * A color pick as the slide renders it, or null for the inherited color. No change (`keep`)
 * shows `current`, the color the selection has now, so its swatch is never empty. The automatic
 * color (`auto`) shows `automatic` when the field knows it (a chart series). @internal
 */
export function renderedColorOf(
  pick: ColorPick,
  rows: readonly (readonly ThemeSwatch[])[],
  current: string | null = null,
  automatic: string | null = null
): string | null {
  if (pick === 'keep') return current;
  if (pick === 'auto') return automatic;
  const choice = colorChoiceOf(pick);
  if (choice === null) return null;
  if (typeof choice === 'string') return choice;
  const shift = choice.shift ?? 0;
  return rows.flat().find((s) => s.slot === choice.theme && s.shift === shift)?.color ?? null;
}

// --- Form dialogs -------------------------------------------------------------------------

/** A problem a form dialog shows: the message, and the field to mark and focus (by `name`). @internal */
export interface DialogError {
  readonly message: string;
  readonly field?: string;
}

/** A dialog error as the per-field record of `UseDialogReturn.errors`. @internal */
export function dialogErrorsOf(error: DialogError | null): Readonly<Record<string, string>> {
  return error ? { [error.field ?? 'form']: error.message } : {};
}

/**
 * Mark the field an error names, tie it to the message (`errorId`), and move focus there. The
 * dialog carries `data-invalid` while an error shows, for host styles. @internal
 */
export function markDialogError(
  dialog: HTMLElement | null | undefined,
  error: DialogError | null,
  errorId: string
): void {
  if (!dialog) return;
  dialog.toggleAttribute('data-invalid', error !== null);
  for (const el of dialog.querySelectorAll('[aria-invalid="true"]')) {
    el.removeAttribute('aria-invalid');
    el.removeAttribute('aria-describedby');
  }
  if (!error?.field) return;
  const field = dialog.querySelector<HTMLElement>(`[name="${CSS.escape(error.field)}"]`);
  if (!field) return;
  field.setAttribute('aria-invalid', 'true');
  field.setAttribute('aria-describedby', errorId);
  field.focus();
}

/**
 * The markers that make a `FormSection` or one field the dialog's `Field` part of `name`.
 * @internal
 */
export const fieldPart = (name: string) =>
  ({ 'data-pptx-part': 'field', 'data-pptx-field': name }) as const;

/** The tab the arrow keys, Home, and End move to in a form dialog's tab list, or -1. @internal */
export function formTabStep(key: string, at: number, count: number): number {
  return key === 'ArrowRight'
    ? (at + 1) % count
    : key === 'ArrowLeft'
      ? (at - 1 + count) % count
      : key === 'Home'
        ? 0
        : key === 'End'
          ? count - 1
          : -1;
}

/** The id of a form tab, and of its panel. @internal */
export const formTabId = (base: string, id: string): string => `${base}-tab-${id}`;
/** @internal */
export const formPanelId = (base: string, id: string): string => `${base}-panel-${id}`;

/** A font family as a CSS `font-family` value, with quotes and backslashes dropped. @internal */
export const cssFontFamily = (font: string): string => `"${font.replace(/["\\]/g, '')}"`;

// --- Menu rows ----------------------------------------------------------------------------

/**
 * Drop separators a hidden row left behind: at the start, at the end, and two in a row.
 * @internal
 */
export function tidySeparators(entries: readonly ChromeMenuEntry[]): ChromeMenuEntry[] {
  const out: ChromeMenuEntry[] = [];
  for (const entry of entries) {
    if (
      entry.kind === 'separator' &&
      (out.length === 0 || out[out.length - 1]!.kind === 'separator')
    )
      continue;
    out.push(entry);
  }
  while (out.length && out[out.length - 1]!.kind === 'separator') out.pop();
  return out;
}

/**
 * How a slot's menu row shows its state. A two-state row (Bold, Ruler) is a checkbox row; a
 * choice row (View > Normal, a Shadow preset) is a radio row; a grouped value (Shape Effects)
 * holds one choice for each group. `checked` is undefined for a plain row. @internal
 */
export function menuRowCheck(
  row: ResolvedMenuRow,
  control: ChromeControl,
  state: ChromeCommandState,
  value: string | undefined
): { readonly checked: boolean | undefined; readonly radio: boolean } {
  const checked = row.toggle
    ? state.isActive
    : control.toggle === 'choice' && value !== undefined
      ? state.value === value
      : control.optionGroups && value !== undefined
        ? chromeOptionChecked(state.value, value)
        : undefined;
  return {
    checked,
    radio: (!!control.optionGroups || control.toggle === 'choice') && value !== undefined,
  };
}

/** The key of a menu entry's row, for host overrides. @internal */
export function menuEntryKey(entry: { readonly slot: ChromeSlotId; readonly value?: string }) {
  return entry.value === undefined ? entry.slot : `${entry.slot}.${entry.value}`;
}

// --- The focused notes field --------------------------------------------------------------

/** The text slots that format the speaker notes while the notes field has focus. @internal */
export const NOTES_SLOTS: Partial<Record<ChromeSlotId, NotesFormat>> = {
  'text.bold': 'bold',
  'text.italic': 'italic',
  'text.underline': 'underline',
  'text.strike': 'strike',
};

/** What the toolbar needs from a focused notes field. @internal */
export interface NotesTarget {
  /** The formats of the selected notes text, as the field's own buttons show them. */
  readonly formats: Readonly<Record<NotesFormat, boolean>>;
  readonly readOnly: boolean;
  /** Whether text is selected, or the caret is in a word: else a format has nothing to take. */
  readonly canFormat: boolean;
  /** Format the selected notes text, or the word that holds the caret; false on a refusal. */
  readonly apply: (format: NotesFormat) => boolean;
}

const notesTargets = new WeakMap<object, NotesTarget>();
const notesListeners = new Set<() => void>();

/** Set, or clear with null, the focused notes field of `editor`. @internal */
export function setNotesTarget(editor: object, target: NotesTarget | null): void {
  if (target) notesTargets.set(editor, target);
  else if (notesTargets.has(editor)) notesTargets.delete(editor);
  else return;
  for (const listen of notesListeners) listen();
}

/** The focused notes field of `editor`, or null. @internal */
export function notesTargetOf(editor: object | null | undefined): NotesTarget | null {
  return editor ? (notesTargets.get(editor) ?? null) : null;
}

/** Be told when any editor's focused notes field changes. Returns the unsubscribe. @internal */
export function subscribeNotesTargets(listen: () => void): () => void {
  notesListeners.add(listen);
  return () => void notesListeners.delete(listen);
}
