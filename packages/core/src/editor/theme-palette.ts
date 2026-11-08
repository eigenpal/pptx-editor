// The theme color grid of the color menus (Font Color, Shape Fill, Shape Outline,
// Format Background): the ten theme colors of the current slide, each with five lighter or
// darker variants below it.
//
// Colors resolve through the same cascade as paint: the slide's theme, the master `p:clrMap`
// with any layout or slide `p:clrMapOvr`, and the ECMA-376 `lumMod`/`lumOff` transform math
// that `cascade/color.ts` applies (`drawingml-color.ts` in `@eigenpal/ooxml-core`). A swatch is
// therefore the color a slide renders for the same choice.

import {
  applyTransforms,
  DEFAULT_COLOR_MAP,
  parseTheme,
  readColorMap,
  schemeColor,
  type ColorContext,
  type Rgba,
} from '../cascade/index.ts';
import {
  child,
  NS,
  path,
  xmlPartOf,
  type Element,
  type PptxPackage,
  slideRefAt,
} from '../store/index.ts';
import type { ThemeColorSlot, ThemeSwatch } from '../contracts/index.ts';
import { hexOf } from './selection-state.ts';

/** The theme columns, in menu order. */
const SLOTS: readonly ThemeColorSlot[] = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

/** The grid of a palette. */
export type ThemeColorGrid = readonly (readonly ThemeSwatch[])[];

/**
 * The five variants offered under a theme color, chosen by the color's HSL
 * lightness: lighter steps for black and dark colors, darker steps for white and light
 * colors, and both for the rest. Positive is "Lighter n%", negative is "Darker n%".
 */
export function variantsFor(color: Rgba): readonly number[] {
  const max = Math.max(color.r, color.g, color.b) / 255;
  const min = Math.min(color.r, color.g, color.b) / 255;
  const l = (max + min) / 2;
  if (l === 0) return [50, 35, 25, 15, 5];
  if (l === 1) return [-5, -15, -25, -35, -50];
  if (l < 0.2) return [90, 75, 50, 25, 10];
  if (l > 0.8) return [-10, -25, -50, -75, -90];
  return [80, 60, 40, -25, -50];
}

/** An `a:` transform element (`a:lumMod`, `a:lumOff`) with a value in thousandths of a percent. */
function transform(local: string, percent: number): Element {
  return {
    type: 'element',
    id: '',
    kind: 'generic',
    ns: NS.a,
    local,
    nsDecls: [],
    attrs: [{ ns: '', local: 'val', value: String(percent * 1000) }],
    children: [],
  };
}

/** A theme color with a variant applied, as `a:lumMod`/`a:lumOff` would render it. */
export function shiftColor(base: Rgba, shift: number): Rgba {
  if (shift === 0) return base;
  const amount = Math.abs(shift);
  const transforms = [transform('lumMod', 100 - amount)];
  if (shift > 0) transforms.push(transform('lumOff', amount));
  return applyTransforms(base, transforms);
}

/** The grid for a theme color context: one row of theme colors, then five variant rows. */
export function themeColorGrid(ctx: ColorContext): ThemeColorGrid {
  const columns = SLOTS.map((slot) => {
    const base = schemeColor(slot, ctx) ?? { r: 0, g: 0, b: 0, a: 1 };
    const scheme = ctx.colorMap.get(slot) ?? slot;
    return [0, ...variantsFor(base)].map((shift): ThemeSwatch => ({
      slot,
      color: hexOf(shiftColor(base, shift)),
      shift,
      scheme,
    }));
  });
  return [0, 1, 2, 3, 4, 5].map((row) => columns.map((column) => column[row]));
}

/** The color context of slide `index`: its theme and its effective color map. */
function slideColorContext(pkg: PptxPackage, index: number): ColorContext | null {
  const ref = slideRefAt(pkg, index);
  if (!ref) return null;
  const theme = parseTheme(xmlPartOf(pkg, ref.themePart));
  // A master in the slide master view has no master part: its own color map applies.
  const master = xmlPartOf(pkg, ref.masterPart ?? ref.partName)?.root;
  let colorMap = readColorMap(child(master, 'p:clrMap')) ?? DEFAULT_COLOR_MAP;
  for (const part of [ref.layoutPart, ref.partName]) {
    const override = readColorMap(
      path(xmlPartOf(pkg, part)?.root, 'p:clrMapOvr', 'a:overrideClrMapping')
    );
    if (override) colorMap = override;
  }
  return { scheme: theme.colors, colorMap };
}

/**
 * A reference-stable source of the grid: it returns the previous grid while the colors are
 * the same, so a `useEditorState` selector on it does not re-render on every edit.
 */
export function createThemeColorSource(): (
  pkg: PptxPackage | undefined,
  index: number
) => ThemeColorGrid {
  let last: ThemeColorGrid = [];
  let lastKey = '';
  return (pkg, index) => {
    const ctx = pkg ? slideColorContext(pkg, index) : null;
    const next = ctx ? themeColorGrid(ctx) : [];
    const key = next.map((row) => row.map((s) => s.color).join(',')).join(';');
    if (key !== lastKey) {
      last = next;
      lastKey = key;
    }
    return last;
  };
}
