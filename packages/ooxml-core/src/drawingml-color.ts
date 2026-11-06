// DrawingML color models and color transforms (ECMA-376 Part 1, 20.1.2.3), format-neutral: the
// six color models, the transform math, sRGB and linear (scRGB) components, and HSL. It reads
// element names and attribute strings only, so it works over any XML tree. The theme lookup is
// the caller's, because each document format maps scheme color names in its own way.
//
// This is PPTX Editor's module (from `packages/core/src/cascade/color.ts`), not a copy of a
// DOCX Editor file. DOCX Editor resolves the same colors in
// docx-editor/packages/core/src/store/package/drawing-shape-color.ts at 5993bd223, with other
// math. The fixes below also apply to DOCX Editor. The color
// values are measured on the reference PDF of the `colors` fixture of PPTX Editor:
// - `a:shade` uses the sRGB transfer curve and rounds once at the end. A 2.3 gamma with a floor
//   draws accent2 ED7D31 + `shade 50000` as AF5C24; the reference PDF draws AE5A21.
// - `a:inv` inverts linear components: accent1 4472C4 gives F8EBB3.
// - `a:gray` uses Rec. 709 weights on sRGB components: accent2 ED7D31 gives 8F8F8F.
// - `a:hslClr`, `a:lum`, `a:sat`, `a:hue`, `a:comp`, `a:inv`, `a:gray`, the
//   `a:red`/`a:green`/`a:blue` family, `a:gamma`, and `a:invGamma` resolve here;
//   drawing-shape-color.ts refuses the whole color.
// - The integer spelling of a percentage is an `xsd:int`, so it can carry a leading `+`.
//   `transformRatio` in drawing-shape-color.ts refuses `a:tint`, `a:shade`, and `a:alpha` with
//   `val="+50000"`.
// Both modules read the percentage spelling of `val` (`50%`). The `a:ST_Percentage` family is a
// union of an `xsd:int` in thousandths of a percent and the `s:ST_Percentage` pattern
// `-?[0-9]+(\.[0-9]+)?%`, in the Strict and the Transitional schemas alike.
// drawing-shape-color.ts has two rules that this module does not have. Neither is a spec rule:
// the schema leaves an invalid value and a missing `lastClr` to the consumer.
// - a malformed or out-of-range `val` refuses the color, where this module reads a malformed
//   one as 0 and clamps the result of an out-of-range one;
// - `a:sysClr` with no `lastClr` takes a default for each system color, where this module
//   draws white for `window` and black for every other name.
//
// Space policy (one place, so calibration against reference PDFs changes one line):
// - `lum*`, `sat*`, `hue*`, `comp` work in HSL over sRGB components.
// - `tint`, `shade`, and the `red/green/blue*` family work in linear (scRGB) components, which
//   is the documented implementer behavior for those transforms.
// - `inv` inverts linear (scRGB) components; `gray` uses Rec. 709 weights on sRGB components.
//   Both are measured on the `colors` fixture (see the transform cases below).

import { PRESET_COLORS } from './drawingml-preset-colors.ts';

export { PRESET_COLORS } from './drawingml-preset-colors.ts';

/** A resolved color: sRGB 0–255 integers and alpha 0–1. */
export interface Rgba {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}

/** One color transform child (`a:lumMod`, `a:inv`, …): its local name and its raw `val`. */
export interface DrawingColorTransform {
  readonly name: string;
  readonly val?: string;
}

/** Reads one attribute of a color model element by local name (`val`, `lastClr`, `r`, …). */
export type DrawingColorAttributeReader = (name: string) => string | undefined;

/** Resolves a scheme color name (`accent1`, `tx1`, `phClr`) against the caller's theme. */
export type DrawingSchemeColorResolver = (name: string) => Rgba | undefined;

/** Opaque black, the fallback color. Frozen, because every caller shares it. */
export const BLACK: Rgba = Object.freeze({ r: 0, g: 0, b: 0, a: 1 });
const WHITE: Rgba = Object.freeze({ r: 255, g: 255, b: 255, a: 1 });

/** The local names of the six DrawingML color models (ECMA-376 Part 1, 20.1.2.3). */
const COLOR_MODELS: ReadonlySet<string> = new Set([
  'srgbClr',
  'schemeClr',
  'sysClr',
  'prstClr',
  'scrgbClr',
  'hslClr',
]);

const clamp01 = (v: number): number => (v < 0 ? 0 : v > 1 ? 1 : v);
const toByte = (v: number): number => Math.round(clamp01(v) * 255);

/**
 * An `xsd:int` value (`ST_Angle`, `ST_PositiveFixedAngle`): an optional sign and up to 15
 * digits, after the schema's whitespace collapse. Anything else reads as `undefined`, so a
 * corrupt file falls back to the default instead of propagating NaN.
 */
function intValue(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim();
  if (!/^[+-]?\d{1,15}$/.test(text)) return undefined;
  const v = Number(text);
  return Number.isSafeInteger(v) ? v : undefined;
}

/**
 * A percentage value of the `a:ST_Percentage` family, in thousandths of a percent. The schema
 * type is a union of two spellings, and files use both: an `xsd:int` (`50000`) and the
 * `s:ST_Percentage` pattern `-?[0-9]+(\.[0-9]+)?%` (`50%`, `12.5%`). The pattern check is
 * anchored and bounded to 10 digits on each side of the point, so a hostile value cannot
 * backtrack or overflow. Anything else reads as `undefined`.
 */
function percentageValue(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const text = raw.trim();
  if (!text.endsWith('%')) return intValue(text);
  if (!/^-?\d{1,10}(?:\.\d{1,10})?%$/.test(text)) return undefined;
  return Number(text.slice(0, -1)) * 1000;
}

/** sRGB component (0–1) to linear. */
export function srgbToLinear(c: number): number {
  return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
}

/** Linear component (0–1) to sRGB. */
export function linearToSrgb(c: number): number {
  return c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055;
}

interface Hsl {
  h: number; // 0–360
  s: number; // 0–1
  l: number; // 0–1
}

function rgbToHsl(r: number, g: number, b: number): Hsl {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return { h: h * 60, s, l };
}

function hslToRgb({ h, s, l }: Hsl): [number, number, number] {
  if (s === 0) return [l, l, l];
  const hue = (((h % 360) + 360) % 360) / 360;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number): number => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return [f(hue + 1 / 3), f(hue), f(hue - 1 / 3)];
}

/** Parse `RRGGBB`. Anything else reads as undefined. */
export function parseHex(hex: string | undefined): Rgba | undefined {
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  const n = parseInt(hex, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

/** `#rrggbb`. */
export function toHex(c: Rgba): string {
  const h = (v: number): string =>
    toByte(v / 255)
      .toString(16)
      .padStart(2, '0');
  return `#${h(c.r)}${h(c.g)}${h(c.b)}`;
}

/** `rgba(...)` CSS, for paint. */
export function toCss(c: Rgba): string {
  if (c.a >= 1) return toHex(c);
  return `rgba(${c.r},${c.g},${c.b},${Math.round(c.a * 1000) / 1000})`;
}

/** Whether a DrawingML local name is one of the six color models (`srgbClr`, `schemeClr`, …). */
export function isDrawingColorModel(local: string): boolean {
  return COLOR_MODELS.has(local);
}

/**
 * The base color of one color model element, before its transforms, or undefined if it cannot
 * resolve. `model` is the element's local name in the DrawingML main namespace.
 */
export function drawingColorBase(
  model: string,
  readAttribute: DrawingColorAttributeReader,
  resolveScheme: DrawingSchemeColorResolver
): Rgba | undefined {
  switch (model) {
    case 'srgbClr':
      return parseHex(readAttribute('val'));
    case 'schemeClr':
      return resolveScheme(readAttribute('val') ?? '');
    case 'sysClr':
      return (
        parseHex(readAttribute('lastClr')) ?? (readAttribute('val') === 'window' ? WHITE : BLACK)
      );
    case 'prstClr':
      return parseHex(PRESET_COLORS[readAttribute('val') ?? '']);
    case 'scrgbClr': {
      const pct = (n: string): number =>
        clamp01((percentageValue(readAttribute(n)) ?? 0) / 100_000);
      return {
        r: toByte(linearToSrgb(pct('r'))),
        g: toByte(linearToSrgb(pct('g'))),
        b: toByte(linearToSrgb(pct('b'))),
        a: 1,
      };
    }
    case 'hslClr': {
      const [r, g, b] = hslToRgb({
        h: (intValue(readAttribute('hue')) ?? 0) / 60_000,
        s: clamp01((percentageValue(readAttribute('sat')) ?? 0) / 100_000),
        l: clamp01((percentageValue(readAttribute('lum')) ?? 0) / 100_000),
      });
      return { r: toByte(r), g: toByte(g), b: toByte(b), a: 1 };
    }
  }
  return undefined;
}

/**
 * Apply color transforms in document order, as the spec requires (each transform sees the
 * result of the previous one). The caller passes only the children in the DrawingML main
 * namespace; an unknown name is skipped.
 */
export function applyColorTransforms(
  base: Rgba,
  transforms: readonly DrawingColorTransform[]
): Rgba {
  let r = base.r / 255;
  let g = base.g / 255;
  let b = base.b / 255;
  let a = base.a;
  for (const t of transforms) {
    // `a:hue` and `a:hueOff` are angles (`ST_PositiveFixedAngle`, `ST_Angle`) in 60000ths of a
    // degree. Every other `val` is a percentage in thousandths of a percent.
    const isAngle = t.name === 'hue' || t.name === 'hueOff';
    const raw = (isAngle ? intValue(t.val) : percentageValue(t.val)) ?? 0;
    const v = raw / 100_000;
    switch (t.name) {
      case 'alpha':
        a = clamp01(v);
        break;
      case 'alphaMod':
        a = clamp01(a * v);
        break;
      case 'alphaOff':
        a = clamp01(a + v);
        break;
      case 'lum':
      case 'lumMod':
      case 'lumOff':
      case 'sat':
      case 'satMod':
      case 'satOff':
      case 'hue':
      case 'hueMod':
      case 'hueOff':
      case 'comp': {
        const hsl = rgbToHsl(r, g, b);
        if (t.name === 'lum') hsl.l = clamp01(v);
        else if (t.name === 'lumMod') hsl.l = clamp01(hsl.l * v);
        else if (t.name === 'lumOff') hsl.l = clamp01(hsl.l + v);
        else if (t.name === 'sat') hsl.s = clamp01(v);
        else if (t.name === 'satMod') hsl.s = clamp01(hsl.s * v);
        else if (t.name === 'satOff') hsl.s = clamp01(hsl.s + v);
        else if (t.name === 'hue') hsl.h = raw / 60_000;
        else if (t.name === 'hueMod') hsl.h = hsl.h * v;
        else if (t.name === 'hueOff') hsl.h = hsl.h + raw / 60_000;
        else hsl.h = hsl.h + 180;
        [r, g, b] = hslToRgb(hsl);
        break;
      }
      case 'tint':
      case 'shade':
      case 'red':
      case 'redMod':
      case 'redOff':
      case 'green':
      case 'greenMod':
      case 'greenOff':
      case 'blue':
      case 'blueMod':
      case 'blueOff': {
        let lr = srgbToLinear(r);
        let lg = srgbToLinear(g);
        let lb = srgbToLinear(b);
        switch (t.name) {
          case 'tint':
            lr = 1 - (1 - lr) * v;
            lg = 1 - (1 - lg) * v;
            lb = 1 - (1 - lb) * v;
            break;
          case 'shade':
            lr *= v;
            lg *= v;
            lb *= v;
            break;
          case 'red':
            lr = v;
            break;
          case 'redMod':
            lr *= v;
            break;
          case 'redOff':
            lr += v;
            break;
          case 'green':
            lg = v;
            break;
          case 'greenMod':
            lg *= v;
            break;
          case 'greenOff':
            lg += v;
            break;
          case 'blue':
            lb = v;
            break;
          case 'blueMod':
            lb *= v;
            break;
          case 'blueOff':
            lb += v;
            break;
        }
        r = linearToSrgb(clamp01(lr));
        g = linearToSrgb(clamp01(lg));
        b = linearToSrgb(clamp01(lb));
        break;
      }
      // ECMA-376 Part 1, 20.1.2.3.17 inv: the inverse of the input color. The reference PDF of
      // the `colors` fixture inverts linear components: accent1 4472C4 with `a:inv` draws
      // F8EBB3 (248, 235, 179), not the sRGB inverse BB8D3B.
      case 'inv':
        r = linearToSrgb(1 - srgbToLinear(r));
        g = linearToSrgb(1 - srgbToLinear(g));
        b = linearToSrgb(1 - srgbToLinear(b));
        break;
      // ECMA-376 Part 1, 20.1.2.3.9 gray: a grayscale of the input color that takes into account
      // the relative intensities of the red, green, and blue primaries. sRGB uses the Rec. 709
      // primaries, so the weights are 0.2126, 0.7152, and 0.0722 on sRGB components. The
      // reference PDF of the `colors` fixture draws accent2 ED7D31 with `a:gray` as 8F8F8F (143);
      // Rec. 601 weights give 150, and weights on linear components give 156.
      case 'gray': {
        const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
        r = g = b = y;
        break;
      }
      case 'gamma':
        r = linearToSrgb(r);
        g = linearToSrgb(g);
        b = linearToSrgb(b);
        break;
      case 'invGamma':
        r = srgbToLinear(r);
        g = srgbToLinear(g);
        b = srgbToLinear(b);
        break;
    }
  }
  return { r: toByte(r), g: toByte(g), b: toByte(b), a: clamp01(a) };
}

/**
 * One color model element in a neutral shape, so each editor can pass its own XML tree: the
 * element's local name, a reader for its attributes, and its transform children in the
 * DrawingML main namespace, in document order.
 */
export interface DrawingColorElement {
  readonly model: string;
  readonly readAttribute: DrawingColorAttributeReader;
  readonly transforms: readonly DrawingColorTransform[];
}

/**
 * Resolves one color model element and its transforms to RGBA: `drawingColorBase`, then
 * `applyColorTransforms`. It returns `undefined` when the model is unknown or its base color
 * does not resolve.
 */
export function resolveDrawingColor(
  color: DrawingColorElement,
  resolveScheme: DrawingSchemeColorResolver
): Rgba | undefined {
  const base = drawingColorBase(color.model, color.readAttribute, resolveScheme);
  return base && applyColorTransforms(base, color.transforms);
}
