// The color field of the form dialogs (Font color, Underline color, the bullet color), shared by
// the adapters: the choices it offers, and the color OK sends. A field offers the inherited
// color, "no change" when the selection already has a color of its own, and the ten theme
// colors in the order of the theme color row. A field over runs or paragraphs whose
// colors differ opens at "Mixed", which also sends no change. The swatch grid beside it picks
// any theme tint or shade (`accent1:40`) or a standard color (`#C00000`).

import type {
  ColorChoice,
  EditorSnapshot,
  ThemeColorSlot,
  ThemeSwatch,
} from '../contracts/index.ts';

/** The theme colors a color field offers, in the order of the theme color row. @internal */
export const COLOR_FIELD_SLOTS: readonly ThemeColorSlot[] = [
  'tx1',
  'bg1',
  'tx2',
  'bg2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

/**
 * A color field's value: the inherited color (`auto`), no change (`keep`, or `mixed` when the
 * selection's colors differ), a theme color, a theme tint or shade (`accent1:40` is Lighter 40%,
 * `accent1:-25` Darker 25%), or `#RRGGBB`. @public
 */
export type ColorPick =
  'auto' | 'keep' | 'mixed' | ThemeColorSlot | `${ThemeColorSlot}:${number}` | `#${string}`;

/**
 * The picks a color field offers. `keep` and `mixed` only when the field seeded with them, and
 * a tint, shade, or standard color only while it is the value (the swatch grid picks it).
 * @internal
 */
export function colorPicks(seed: ColorPick | undefined, value?: ColorPick): readonly ColorPick[] {
  const own = seed === 'keep' || seed === 'mixed' ? [seed] : [];
  const extra = value && !/^(auto|keep|mixed|[a-z]+[0-9]?)$/.test(value) ? [value] : [];
  return ['auto', ...own, ...COLOR_FIELD_SLOTS, ...extra];
}

/** The pick of a color choice from the swatch grid. @internal */
export function colorPickOf(choice: ColorChoice): ColorPick {
  if (typeof choice === 'string') return `#${choice.replace(/^#/, '').toUpperCase()}`;
  return choice.shift ? `${choice.theme}:${choice.shift}` : choice.theme;
}

/** The color choice of a pick, or null for `auto`, `keep`, and `mixed`. @internal */
export function colorChoiceOf(pick: ColorPick): ColorChoice | null {
  if (pick === 'auto' || pick === 'keep' || pick === 'mixed') return null;
  if (pick.startsWith('#')) return pick;
  const [theme, shift] = pick.split(':') as [ThemeColorSlot, string?];
  return shift ? { theme, shift: Number(shift) } : { theme };
}

/** The color a field sends: undefined when it did not change, null for the inherited color. @internal */
export function colorChange(seed: ColorPick, pick: ColorPick): ColorChoice | null | undefined {
  if (pick === seed || pick === 'keep' || pick === 'mixed') return undefined;
  return pick === 'auto' ? null : colorChoiceOf(pick);
}

/**
 * The text a custom color field keeps: up to six hex digits, upper case, without `#`, as DOCX
 * Editor's color picker keeps its custom hex field. @internal
 */
export function customHexDraft(text: string): string {
  return text
    .replace(/[^0-9A-Fa-f]/g, '')
    .slice(0, 6)
    .toUpperCase();
}

/** The color of a complete custom hex draft (`#RRGGBB`), or null while it is not complete. @internal */
export function customHexColor(draft: string): `#${string}` | null {
  const hex = customHexDraft(draft);
  return hex.length === 6 ? `#${hex}` : null;
}

/** The draft a custom color field opens with: the selected color when it is a hex color. @internal */
export function customHexSeed(selected: ColorChoice | null | undefined): string {
  return typeof selected === 'string' ? customHexDraft(selected) : '';
}

/**
 * The six hex digits that stand for the current color in a custom color field: its placeholder
 * while the field is empty, and the color the browser's selector shows. A hex choice shows
 * itself, a theme choice shows its swatch as rendered, and no color shows black, the
 * selector's own default. @internal
 */
export function customHexCurrent(
  selected: ColorChoice | null | undefined,
  rows: readonly (readonly ThemeSwatch[])[]
): string {
  if (typeof selected === 'string') return customHexDraft(selected).padEnd(6, '0');
  const swatch = selected
    ? rows.flat().find((s) => s.slot === selected.theme && (s.shift ?? 0) === (selected.shift ?? 0))
    : undefined;
  return swatch ? customHexDraft(swatch.color) : '000000';
}

/** The `#RRGGBB` of a snapshot color, or null for `none`, `gradient`, `other`, and no value. */
function hexOf(color: string | undefined): string | null {
  return color && /^#[0-9A-Fa-f]{6}$/.test(color) ? color.toUpperCase() : null;
}

/**
 * The color a color menu marks as current: the selected text's color for Font Color, and the
 * selected shapes' solid fill or outline color for Shape Fill and Shape Outline, as `#RRGGBB`.
 * Null when the snapshot has none (a mixed selection, a gradient, a highlight). @internal
 */
export function colorMenuSelection(slot: string, snapshot: EditorSnapshot): ColorChoice | null {
  if (slot === 'text.color') return hexOf(snapshot.format?.color);
  if (slot === 'shape.fill') return hexOf(snapshot.geometry?.fill);
  if (slot === 'shape.lineColor') return hexOf(snapshot.geometry?.lineColor);
  return null;
}
