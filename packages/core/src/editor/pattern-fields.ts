// The pattern fill dialog that both adapters show: the gallery order of the 54 preset patterns,
// their names, a CSS swatch of each from the bitmaps the renderer draws, and the pattern the
// dialog opens with. Framework-free, so the React and Vue dialogs agree. `@internal`.

import type { Translate, TranslationKey } from '@pptx-editor.dev/i18n';
import type { PatternFill, PatternPreset } from '../contracts/index.ts';
import { patternBitmap } from '../output/preset-patterns.ts';

/** The presets in gallery order: the percentages, then the lines, then the figures. @internal */
export const PATTERN_CHOICES: readonly PatternPreset[] = [
  'pct5',
  'pct10',
  'pct20',
  'pct25',
  'pct30',
  'pct40',
  'pct50',
  'pct60',
  'pct70',
  'pct75',
  'pct80',
  'pct90',
  'ltDnDiag',
  'ltUpDiag',
  'dkDnDiag',
  'dkUpDiag',
  'wdDnDiag',
  'wdUpDiag',
  'ltVert',
  'ltHorz',
  'narVert',
  'narHorz',
  'dkVert',
  'dkHorz',
  'dashDnDiag',
  'dashUpDiag',
  'dashHorz',
  'dashVert',
  'smConfetti',
  'lgConfetti',
  'zigZag',
  'wave',
  'diagBrick',
  'horzBrick',
  'weave',
  'plaid',
  'divot',
  'dotGrid',
  'dotDmnd',
  'shingle',
  'trellis',
  'sphere',
  'smGrid',
  'lgGrid',
  'smCheck',
  'lgCheck',
  'openDmnd',
  'solidDmnd',
  'horz',
  'vert',
  'dnDiag',
  'upDiag',
  'cross',
  'diagCross',
];

/** The pattern a pattern dialog opens with: the shown one, or 50% in `base` on white. @internal */
export function patternDialogStart(
  shown: PatternFill | null | undefined,
  base = '#4472C4'
): PatternFill {
  return shown ?? { preset: 'pct50', foreground: base, background: '#FFFFFF' };
}

/** The accessible name of a preset pattern, such as "25%" or "Dark downward diagonal". @internal */
export function patternLabel(t: Translate, preset: PatternPreset): string {
  const percent = /^pct(\d+)$/.exec(preset);
  if (percent) return t('pattern.percent', { percent: Number(percent[1]) });
  return t(`pattern.name.${preset}` as TranslationKey);
}

const HEX = /^#[0-9a-f]{6}$/i;

/** The color between two `#RRGGBB` colors at coverage `k` of the first. */
function mix(fg: string, bg: string, k: number): string {
  const ch = (hex: string, i: number): number => parseInt(hex.slice(1 + 2 * i, 3 + 2 * i), 16);
  const out = [0, 1, 2].map((i) => Math.round(ch(fg, i) * k + ch(bg, i) * (1 - k)));
  return `#${out.map((v) => v.toString(16).padStart(2, '0')).join('')}`;
}

/**
 * A CSS `background` of a pattern: its 8 × 8 bitmap as an SVG tile of `tile` CSS pixels, from
 * the bitmaps the slide renderer draws. Colors that are not `#RRGGBB` draw as black on white.
 * @internal
 */
export function patternCss(p: PatternFill, tile = 8): string {
  const bits = patternBitmap(p.preset) ?? [];
  const fg = HEX.test(p.foreground) ? p.foreground : '#000000';
  const bg = HEX.test(p.background) ? p.background : '#FFFFFF';
  let rects = '';
  bits.forEach((row, y) =>
    row.forEach((k, x) => {
      if (k > 0) rects += `<rect x="${x}" y="${y}" width="1" height="1" fill="${mix(fg, bg, k)}"/>`;
    })
  );
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8" shape-rendering="crispEdges">` +
    `<rect width="8" height="8" fill="${bg}"/>${rects}</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 0 0 / ${tile}px ${tile}px repeat`;
}
