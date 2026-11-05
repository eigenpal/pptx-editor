// Framework-free rules of the insert, shape, table, and slide chrome, shared by the React and
// Vue adapters (`ShapeControls`, `ShapeFormatControls`, `TableControls`, `SlideControls`), so the
// two cannot drift: the drawn line glyphs, the glow grid's rows, the custom gradient dialog's
// stop edits, the table size grid's moves, and the layout gallery's groups and keys.

import type { GradientFill, SlideLayoutSummary } from '../contracts/index.ts';
import type { ChromeControl } from './chrome-controls.ts';
import { handleMenuKey } from './menu-keyboard.ts';
import { presetGradient } from './palette.ts';

// ---------------------------------------------------------------------------------------------
// Line glyphs

/**
 * One drawn part of a line glyph: a stroked path (`join` adds a round line join), a solid
 * arrowhead, or a solid dot. @internal
 */
export type LineGlyphPart =
  | { readonly kind: 'stroke'; readonly d: string; readonly join?: true }
  | { readonly kind: 'solid'; readonly d: string }
  | { readonly kind: 'dot'; readonly cx: number; readonly cy: number; readonly r: number };

/** The stroke of a line glyph's stroked parts. @internal */
export const LINE_GLYPH_STROKE = {
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 2,
  strokeLinecap: 'round',
} as const;

/**
 * The parts of a line kind's or a line end's glyph, drawn at the right of a 48×24 box. An
 * `arrow` without `start` is a line kind (an arrow); with `start`, an open arrowhead. @internal
 */
export function lineGlyphParts(value: string, start?: boolean): readonly LineGlyphPart[] {
  switch (value) {
    case 'line':
      return [{ kind: 'stroke', d: 'M8 18 40 6' }];
    case 'elbow':
      return [{ kind: 'stroke', d: 'M8 18H24V6H40', join: true }];
    case 'curved':
      return [{ kind: 'stroke', d: 'M8 18C26 18 22 6 40 6' }];
    case 'arrow':
      return start === undefined
        ? [
            { kind: 'stroke', d: 'M8 18 36 7.5' },
            { kind: 'solid', d: 'M41 6 31.6 5.7 35.4 13.9Z' },
          ]
        : [
            { kind: 'stroke', d: 'M8 12H39' },
            { kind: 'stroke', d: 'M33 6 40 12 33 18', join: true },
          ];
    case 'triangle':
      return [
        { kind: 'stroke', d: 'M8 12H32' },
        { kind: 'solid', d: 'M42 12 31 6.5V17.5Z' },
      ];
    case 'stealth':
      return [
        { kind: 'stroke', d: 'M8 12H32' },
        { kind: 'solid', d: 'M42 12 30 6 33.5 12 30 18Z' },
      ];
    case 'oval':
      return [
        { kind: 'stroke', d: 'M8 12H33' },
        { kind: 'dot', cx: 37, cy: 12, r: 4.5 },
      ];
    case 'diamond':
      return [
        { kind: 'stroke', d: 'M8 12H32' },
        { kind: 'solid', d: 'M37 6.5 42.5 12 37 17.5 31.5 12Z' },
      ];
    default:
      return [{ kind: 'stroke', d: 'M8 12H40' }];
  }
}

/** A line start draws mirrored: the end sits at the left. @internal */
export function lineGlyphTransform(start?: boolean): string | undefined {
  return start ? 'translate(48 0) scale(-1 1)' : undefined;
}

/**
 * Which end a line slot draws: `true` for the line start, `false` for the line end, and
 * `undefined` for the line tool, which shows its kinds. @internal
 */
export function lineSlotEnd(slot: string): boolean | undefined {
  return slot === 'shape.lineStart' ? true : slot === 'shape.lineEnd' ? false : undefined;
}

// ---------------------------------------------------------------------------------------------
// Shape effects

type EffectOption = NonNullable<ChromeControl['options']>[number];

/** One row of the Glow Variations grid: a size, with a cell per accent. @internal */
export interface GlowGridRow {
  readonly size: string;
  readonly cells: readonly { readonly option: EffectOption; readonly accent: string }[];
}

/** The Glow Variations grid: `glow.8.accent2` is the size, then the accent. @internal */
export function glowGridRows(options: readonly EffectOption[]): readonly GlowGridRow[] {
  // A size may hold a decimal point, so the split is before `accent`.
  const parts = (value: string) => {
    const [size = '', accent = 'accent1'] = value.slice('glow.'.length).split(/\.(?=accent)/);
    return { size, accent };
  };
  const sizes = [...new Set(options.map((o) => parts(o.value).size))];
  return sizes.map((size) => ({
    size,
    cells: options
      .filter((o) => parts(o.value).size === size)
      .map((option) => ({ option, accent: parts(option.value).accent })),
  }));
}

/**
 * Whether a separator goes before option `i` of an effect group: in Shadow, a line between the
 * outer and the inner presets. @internal
 */
export function effectSeparatorBefore(
  group: string,
  options: readonly EffectOption[],
  i: number
): boolean {
  return (
    group === 'shadow' &&
    i > 0 &&
    options[i]!.value.startsWith('shadow.inner') &&
    !options[i - 1]!.value.startsWith('shadow.inner')
  );
}

// ---------------------------------------------------------------------------------------------
// Fill and the custom gradient dialog

/** The default custom gradient: Accent 2 of the default theme, light to dark. @internal */
export const DEFAULT_GRADIENT: GradientFill = presetGradient('#ED7D31', 0.35, 0.35);

/** The fill a fill control applies before a pick: the theme's Accent 1, the new-shape fill. @internal */
export function defaultFillChoice(accent1: string | undefined): {
  readonly choice: { readonly theme: 'accent1' } | string;
  readonly color: string;
} {
  return accent1
    ? { choice: { theme: 'accent1' }, color: accent1 }
    : { choice: '#4472C4', color: '#4472C4' };
}

/** The gradient the dialog starts from, with its stops in order. @internal */
export function gradientDialogStart(initial: GradientFill | null | undefined): GradientFill {
  const start = initial ?? DEFAULT_GRADIENT;
  return { ...start, stops: [...start.stops].sort((a, b) => a.pos - b.pos) };
}

/** Change stop `i` of a gradient. @internal */
export function gradientSetStop(
  g: GradientFill,
  i: number,
  patch: Partial<{ pos: number; color: string }>
): GradientFill {
  return { ...g, stops: g.stops.map((s, j) => (j === i ? { ...s, ...patch } : s)) };
}

/** Add a stop after the selected one, half way to the next; it becomes selected. @internal */
export function gradientAddStop(
  g: GradientFill,
  selected: number
): { readonly gradient: GradientFill; readonly selected: number } {
  const stops = [...g.stops].sort((a, b) => a.pos - b.pos);
  const i = Math.min(selected, stops.length - 2);
  const a = stops[i]!;
  const b = stops[i + 1]!;
  const mid = { pos: (a.pos + b.pos) / 2, color: a.color };
  return {
    gradient: { ...g, stops: [...stops.slice(0, i + 1), mid, ...stops.slice(i + 1)] },
    selected: i + 1,
  };
}

/** Remove the selected stop; a gradient keeps two at least. @internal */
export function gradientRemoveStop(
  g: GradientFill,
  selected: number
): { readonly gradient: GradientFill; readonly selected: number } {
  if (g.stops.length <= 2) return { gradient: g, selected };
  return {
    gradient: { ...g, stops: g.stops.filter((_, j) => j !== selected) },
    selected: Math.max(0, selected - 1),
  };
}

/** The stop position (0 to 1) under a pointer on the stop bar. @internal */
export function gradientPosAt(
  rect: { left: number; width: number } | undefined,
  x: number
): number {
  return rect && rect.width > 0 ? Math.max(0, Math.min(1, (x - rect.left) / rect.width)) : 0;
}

/** The angle a typed value sets (0 to 359), or null when it is not a number. @internal */
export function gradientAngleFrom(value: string): number | null {
  const v = Number(value);
  return Number.isFinite(v) ? ((v % 360) + 360) % 360 : null;
}

/** A stop handle's arrow keys: the new position, or null for another key. @internal */
export function gradientHandleKey(key: string, shift: boolean, pos: number): number | null {
  const step = shift ? 0.1 : 0.01;
  if (key === 'ArrowLeft') return Math.max(0, pos - step);
  if (key === 'ArrowRight') return Math.min(1, pos + step);
  return null;
}

// ---------------------------------------------------------------------------------------------
// Insert > Table size grid

/** The size grid: 10 columns by 8 rows. @internal */
export const TABLE_GRID_COLUMNS = 10;
/** @internal */
export const TABLE_GRID_ROWS = 8;

/** A size in the table grid. @internal */
export interface TableGridSize {
  readonly rows: number;
  readonly cols: number;
}

/** The size grid's keys: a move, or null for a key the grid does not take. @internal */
export function tableGridKeyStep(
  key: string
): { readonly rows?: number; readonly cols?: number; readonly toCol?: number } | null {
  if (key === 'ArrowRight') return { cols: 1 };
  if (key === 'ArrowLeft') return { cols: -1 };
  if (key === 'ArrowDown') return { rows: 1 };
  if (key === 'ArrowUp') return { rows: -1 };
  if (key === 'Home') return { toCol: 1 };
  if (key === 'End') return { toCol: TABLE_GRID_COLUMNS };
  return null;
}

/** The cell a move reaches, inside the grid. @internal */
export function tableGridMove(
  current: TableGridSize,
  step: { readonly rows?: number; readonly cols?: number; readonly toCol?: number }
): TableGridSize {
  return {
    rows: Math.min(TABLE_GRID_ROWS, Math.max(1, current.rows + (step.rows ?? 0))),
    cols: Math.min(TABLE_GRID_COLUMNS, Math.max(1, step.toCol ?? current.cols + (step.cols ?? 0))),
  };
}

// ---------------------------------------------------------------------------------------------
// Layout gallery

/** One slide master's layouts in the gallery. @internal */
export interface LayoutMasterGroup {
  readonly id: string;
  readonly name: string;
  readonly layouts: readonly SlideLayoutSummary[];
}

/** The gallery's groups: one per slide master, in the order the layouts list them. @internal */
export function layoutMasterGroups(
  layouts: readonly SlideLayoutSummary[]
): readonly LayoutMasterGroup[] {
  const masters: { id: string; name: string; layouts: SlideLayoutSummary[] }[] = [];
  for (const l of layouts) {
    const group = masters.find((m) => m.id === l.masterId);
    if (group) group.layouts.push(l);
    else masters.push({ id: l.masterId, name: l.masterName, layouts: [l] });
  }
  return masters;
}

/** The tile that takes Tab focus: the current layout, or the first. @internal */
export function layoutGalleryTabStop(
  layouts: readonly SlideLayoutSummary[],
  current: string | undefined
): string | undefined {
  return layouts.some((l) => l.id === current) ? current : layouts[0]?.id;
}

/**
 * The gallery's first focus: a menu opened from the keyboard focuses its first row, so the focus
 * moves on to the current layout. @internal
 */
export function layoutGalleryEnter(gallery: Element, target: EventTarget | null): void {
  const tiles = gallery.querySelectorAll<HTMLElement>('.pptx-layout-gallery__item');
  const marked = gallery.querySelector<HTMLElement>('[data-current]');
  if (marked && target === tiles[0] && marked !== tiles[0]) marked.focus();
}

/**
 * The gallery's keys: the arrows, Home, and End move between tiles. Left from the first column
 * leaves the gallery, so a submenu then closes. Returns true when the key was taken. @internal
 */
export function layoutGalleryKey(
  gallery: HTMLElement,
  target: EventTarget | null,
  key: string
): boolean {
  const tile = (target as HTMLElement | null)?.closest?.<HTMLElement>('.pptx-layout-gallery__item');
  const row = tile?.parentElement ? [...tile.parentElement.children].indexOf(tile) : -1;
  if (key === 'ArrowLeft' && row >= 0 && row % 3 === 0) return false;
  return handleMenuKey(gallery, key);
}
