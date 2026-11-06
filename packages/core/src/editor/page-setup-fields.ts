// The slide size fields of the Page setup dialog, shared by the adapters: a preset size, the
// width and height in inches as the user types them, the orientation, and whether the content
// scales. OK sends one `setSlideSize` command, so the change is one undo step. The field types
// are public, as the dialog drafts that hold them; the functions are adapter glue.

import type { EditorCommand } from '../contracts/index.ts';
import { MAX_SLIDE_SIZE, MIN_SLIDE_SIZE } from '../automation/commands/slide-size.ts';
import { EMU_PER_INCH, amountText, parseAmount } from './paragraph-dialog-fields.ts';

/** A named slide size, or `custom`. @public */
export type SlideSizePreset = 'widescreen' | 'standard' | 'a4' | 'letter' | 'custom';

/** Landscape: the width is the longer side. @public */
export type SlideOrientation = 'landscape' | 'portrait';

/** The named sizes, landscape, in EMU. A4 is ISO 216 (297 × 210 mm). @internal */
export const SLIDE_SIZE_PRESETS: Readonly<
  Record<Exclude<SlideSizePreset, 'custom'>, { readonly width: number; readonly height: number }>
> = {
  widescreen: { width: 12_192_000, height: 6_858_000 },
  standard: { width: 9_144_000, height: 6_858_000 },
  a4: { width: 10_692_000, height: 7_560_000 },
  letter: { width: 10_058_400, height: 7_772_400 },
};

/** The choices of the size list, in display order. @internal */
export const SLIDE_SIZE_CHOICES: readonly SlideSizePreset[] = [
  'widescreen',
  'standard',
  'a4',
  'letter',
  'custom',
];

/** The slide size fields of the Page setup draft. Width and height are typed inches. @public */
export interface SlideSizeFields {
  readonly size: SlideSizePreset;
  readonly width: string;
  readonly height: string;
  readonly orientation: SlideOrientation;
  /** Whether the boxes on every slide scale to the new size (`content: 'scale'`). */
  readonly scaleContent: boolean;
}

const inches = (emu: number): string => amountText(emu / EMU_PER_INCH);

/** The preset of an exact size in either orientation, or `custom`. @internal */
export function slideSizePresetOf(width: number, height: number): SlideSizePreset {
  const long = Math.max(width, height);
  const short = Math.min(width, height);
  for (const name of SLIDE_SIZE_CHOICES) {
    if (name === 'custom') continue;
    const p = SLIDE_SIZE_PRESETS[name];
    if (p.width === long && p.height === short) return name;
  }
  return 'custom';
}

/** The fields a dialog opens with, for the deck's slide size in EMU. @internal */
export function seedSlideSizeFields(width: number, height: number): SlideSizeFields {
  return {
    size: slideSizePresetOf(width, height),
    width: inches(width),
    height: inches(height),
    orientation: height > width ? 'portrait' : 'landscape',
    scaleContent: true,
  };
}

/** The fields after a pick in the size list: a preset fills in its size. @internal */
export function withSlideSizePreset(d: SlideSizeFields, size: SlideSizePreset): SlideSizeFields {
  if (size === 'custom') return { ...d, size };
  const p = SLIDE_SIZE_PRESETS[size];
  const portrait = d.orientation === 'portrait';
  return {
    ...d,
    size,
    width: inches(portrait ? p.height : p.width),
    height: inches(portrait ? p.width : p.height),
  };
}

/** The fields after an orientation pick: the width and height swap when it changes. @internal */
export function withSlideOrientation(
  d: SlideSizeFields,
  orientation: SlideOrientation
): SlideSizeFields {
  if (orientation === d.orientation) return d;
  return { ...d, orientation, width: d.height, height: d.width };
}

/** The fields after typing in Width or Height: the size becomes custom. @internal */
export function withSlideDimension(
  d: SlideSizeFields,
  field: 'width' | 'height',
  text: string
): SlideSizeFields {
  const next = { ...d, [field]: text };
  const w = parseAmount(next.width);
  const h = parseAmount(next.height);
  const orientation: SlideOrientation =
    Number.isFinite(w) && Number.isFinite(h) && w !== h
      ? h > w
        ? 'portrait'
        : 'landscape'
      : d.orientation;
  return { ...next, size: 'custom', orientation };
}

/** The size the fields ask for, in EMU: a preset's exact size, or the typed inches. */
function emuOf(d: SlideSizeFields): { readonly width: number; readonly height: number } {
  if (d.size !== 'custom') {
    const p = SLIDE_SIZE_PRESETS[d.size];
    return d.orientation === 'portrait'
      ? { width: p.height, height: p.width }
      : { width: p.width, height: p.height };
  }
  return {
    width: Math.round(parseAmount(d.width) * EMU_PER_INCH),
    height: Math.round(parseAmount(d.height) * EMU_PER_INCH),
  };
}

const unchanged = (seed: SlideSizeFields, d: SlideSizeFields): boolean =>
  d.width === seed.width && d.height === seed.height && d.size === seed.size;

/**
 * The field that blocks OK, with the translation key of its message, or null. The range is
 * `ST_SlideSizeCoordinate`, 1 to 56 inches. @internal
 */
export function slideSizeFieldError(
  seed: SlideSizeFields,
  d: SlideSizeFields
): { readonly field: 'width' | 'height'; readonly key: 'errors.slideSizeRange' } | null {
  if (unchanged(seed, d)) return null;
  const { width, height } = emuOf(d);
  const bad = (v: number): boolean =>
    !Number.isFinite(v) || v < MIN_SLIDE_SIZE || v > MAX_SLIDE_SIZE;
  if (bad(width)) return { field: 'width', key: 'errors.slideSizeRange' };
  if (bad(height)) return { field: 'height', key: 'errors.slideSizeRange' };
  return null;
}

/**
 * The command OK sends, or null when the size did not change. A size typed back to the opening
 * text sends nothing, so the rounding of the inches never moves a slide. @internal
 */
export function slideSizeCommandOf(
  seed: SlideSizeFields,
  d: SlideSizeFields
): EditorCommand | null {
  if (unchanged(seed, d)) return null;
  const { width, height } = emuOf(d);
  return {
    type: 'setSlideSize',
    width,
    height,
    ...(d.scaleContent ? {} : { content: 'keep' as const }),
  };
}
