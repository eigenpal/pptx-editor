// Shared by the React and Vue adapters (moved from packages/react/src/palette.ts).
//
// The one palette every color menu offers: the deck's theme colors
// (from the engine snapshot, through the slide's color map) and the ten Standard Colors. Every
// swatch has a name from en.json, never a hex code.

import type { GradientFill, ThemeSwatch } from '../contracts/index.ts';
import type { Translate } from '@pptx-editor.dev/i18n';

/** The Standard Colors row, in menu order. @internal */
export const STANDARD_COLORS = [
  { name: 'darkRed', color: '#C00000' },
  { name: 'red', color: '#FF0000' },
  { name: 'orange', color: '#FFC000' },
  { name: 'yellow', color: '#FFFF00' },
  { name: 'lightGreen', color: '#92D050' },
  { name: 'green', color: '#00B050' },
  { name: 'lightBlue', color: '#00B0F0' },
  { name: 'blue', color: '#0070C0' },
  { name: 'darkBlue', color: '#002060' },
  { name: 'purple', color: '#7030A0' },
] as const;

/** A swatch's name: "Accent 1", "Accent 1, Lighter 40%", "Text 1, Darker 25%". @internal */
export function swatchName(t: Translate, swatch: ThemeSwatch): string {
  const color = t(`colors.theme.${swatch.slot}`);
  if (swatch.shift > 0) return t('colors.lighter', { color, percent: swatch.shift });
  if (swatch.shift < 0) return t('colors.darker', { color, percent: -swatch.shift });
  return color;
}

/** Mix a `#RRGGBB` color toward white (`amount` > 0) or black (`amount` < 0). */
function mix(hex: string, amount: number): string {
  const n = parseInt(hex.replace(/^#/, ''), 16);
  const channel = (shift: number): number => {
    const v = (n >> shift) & 255;
    return Math.round(amount >= 0 ? v + (255 - v) * amount : v * (1 + amount));
  };
  return `#${[16, 8, 0]
    .map((s) => channel(s).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;
}

/** A top-to-bottom preset gradient from `base`: lighter by `light`, darker by `dark`. @internal */
export function presetGradient(base: string, light: number, dark: number): GradientFill {
  return {
    kind: 'linear',
    angle: 90,
    stops: [
      { pos: 0, color: mix(base, light) },
      { pos: 1, color: mix(base, -dark) },
    ],
    rotWithShape: true,
  };
}

/**
 * CSS for a gradient preview (DrawingML angles start at left-to-right; CSS at bottom-to-top).
 * Not a stable API. @internal
 */
export function gradientCss(g: GradientFill): string {
  const stops = [...g.stops]
    .sort((a, b) => a.pos - b.pos)
    .map((s) => `${s.color} ${Math.round(s.pos * 1000) / 10}%`)
    .join(', ');
  return g.kind === 'radial'
    ? `radial-gradient(circle, ${stops})`
    : `linear-gradient(${(g.angle + 90) % 360}deg, ${stops})`;
}
