// Text Effects (`shape.textEffects`): the Transform gallery, which sets or removes the text warp
// of the selected shapes (`a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19). The registry holds the
// options, the command, and the value; each adapter draws the gallery cells from
// `textWarpPreview`, so React and Vue show the same paths.

import { evaluatePresetTextWarp, TEXT_WARP_NAMES, type PathCommand } from '@eigenpal/ooxml-core';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type {
  EditorCommand,
  EditorSnapshot,
  TextEffectsValue,
  TextWarpPreset,
} from '../contracts/index.ts';
import type { ChromeControl } from './chrome-controls.ts';

/** Every warp the gallery offers, in schema order, without `textNoShape`. */
const WARPS = TEXT_WARP_NAMES.filter((w) => w !== 'textNoShape') as Exclude<
  TextWarpPreset,
  'textNoShape'
>[];

/** The warps whose text follows one path at its own size: the first gallery section. */
const FOLLOW_PATH: ReadonlySet<string> = new Set([
  'textArchUp',
  'textArchDown',
  'textCircle',
  'textButton',
]);

/**
 * The Transform gallery: its sections, top to bottom, each with its warps. Follow Path holds the
 * warps that lay text along a path; Warp holds the warps that stretch text between paths.
 * @internal
 */
export const TEXT_WARP_GALLERY: readonly {
  readonly section: 'followPath' | 'warp';
  readonly labelKey: TranslationKey;
  readonly warps: readonly Exclude<TextWarpPreset, 'textNoShape'>[];
}[] = [
  {
    section: 'followPath',
    labelKey: 'toolbar.textWarpSections.followPath',
    warps: WARPS.filter((w) => FOLLOW_PATH.has(w)),
  },
  {
    section: 'warp',
    labelKey: 'toolbar.textWarpSections.warp',
    warps: WARPS.filter((w) => !FOLLOW_PATH.has(w)),
  },
];

/**
 * Text Effects: one Transform submenu with No Transform and every text warp. Not in the default
 * formatting bar; the ribbon and a host place it.
 */
export const TEXT_EFFECTS_CONTROL: ChromeControl = {
  id: 'textEffects',
  shape: 'dropdown',
  labelKey: 'toolbar.textEffects',
  paths: null,
  defaultToolbar: false,
  optionGroups: [{ id: 'transform', labelKey: 'toolbar.transform' }],
  options: [
    { value: 'transform.none', labelKey: 'toolbar.noTransform', group: 'transform' },
    ...WARPS.map((w) => ({
      value: `transform.${w}`,
      labelKey: `toolbar.textWarps.${w}` as TranslationKey,
      group: 'transform',
    })),
  ],
  state: { kind: 'value' },
};

/** The command of a Text Effects option (`transform.textArchUp`, `transform.none`). @internal */
export function textEffectsCommand(value: unknown): EditorCommand | null {
  if (typeof value !== 'string' || !value.startsWith('transform.')) return null;
  const pick = value.slice('transform.'.length);
  if (pick === 'none') return { type: 'setTextWarp', preset: null };
  return (WARPS as readonly string[]).includes(pick)
    ? { type: 'setTextWarp', preset: pick as TextWarpPreset }
    : null;
}

/** The value Text Effects shows: the selection's warp, or `transform.none`. @internal */
export function textEffectsValue(snap: EditorSnapshot): TextEffectsValue | undefined {
  const warp = snap.geometry?.textWarp;
  if (warp === undefined) return undefined;
  return warp === 'textNoShape' ? 'transform.none' : `transform.${warp}`;
}

/** SVG path data for commands, with numbers rounded for a small preview. */
function previewData(commands: readonly PathCommand[]): string {
  const n = (v: number): string => String(Math.round(v * 10) / 10);
  let d = '';
  for (const c of commands) {
    if (c.op === 'Z') d += 'Z';
    else if (c.op === 'M' || c.op === 'L') d += `${c.op}${n(c.x)} ${n(c.y)}`;
    else if (c.op === 'Q') d += `Q${n(c.x1)} ${n(c.y1)} ${n(c.x)} ${n(c.y)}`;
    else if (c.op === 'C') d += `C${n(c.x1)} ${n(c.y1)} ${n(c.x2)} ${n(c.y2)} ${n(c.x)} ${n(c.y)}`;
  }
  return d;
}

const previews = new Map<string, readonly string[]>();

/**
 * The paths of a warp as SVG path data in a 40 × 28 box (4 units of margin), for a gallery
 * cell: the cell strokes them. Empty for a warp with no paths. @internal
 */
export function textWarpPreview(preset: string): readonly string[] {
  if (!TEXT_WARP_NAMES.includes(preset)) return [];
  let out = previews.get(preset);
  if (!out) {
    const paths = evaluatePresetTextWarp(preset, 32, 20, new Map()) ?? [];
    out = paths.map((p) =>
      previewData(
        p.commands.map((c) =>
          c.op === 'Z'
            ? c
            : c.op === 'C'
              ? {
                  ...c,
                  x1: c.x1 + 4,
                  y1: c.y1 + 4,
                  x2: c.x2 + 4,
                  y2: c.y2 + 4,
                  x: c.x + 4,
                  y: c.y + 4,
                }
              : c.op === 'Q'
                ? { ...c, x1: c.x1 + 4, y1: c.y1 + 4, x: c.x + 4, y: c.y + 4 }
                : { ...c, x: c.x + 4, y: c.y + 4 }
        )
      )
    );
    previews.set(preset, out);
  }
  return out;
}
