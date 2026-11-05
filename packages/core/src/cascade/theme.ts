// The theme: color scheme, font scheme, and format scheme (fill, line, effect, and background
// style lists that `p:style` references index into).

import { attr, child, elements, path, type Element, type XmlPart } from '../store/index.ts';
import { colorChild, parseHex, resolveColor, type ColorContext, type Rgba } from './color.ts';
import { complexScriptOf, eastAsianScriptOf } from './script-fonts.ts';

/** One script slot of a font scheme collection. */
export interface FontCollection {
  readonly latin?: string;
  readonly ea?: string;
  readonly cs?: string;
  /**
   * The script fonts (`a:font script="Jpan" typeface="…"`, ECMA-376 Part 1, 20.1.4.1.16), by
   * ISO 15924 script code. They name the East Asian and complex-script faces when `a:ea` or
   * `a:cs` is empty. Absent when the collection has none.
   */
  readonly scripts?: ReadonlyMap<string, string>;
}

/** The most script fonts read from one collection. ECMA-376 themes list about 50. */
const MAX_SCRIPT_FONTS = 256;

/** A parsed theme. */
export interface Theme {
  readonly colors: ReadonlyMap<string, Rgba>;
  readonly majorFont: FontCollection;
  readonly minorFont: FontCollection;
  /** `a:fmtScheme` lists. Index 0 is `idx="1"`, as `p:style` references count from 1. */
  readonly fillStyles: readonly Element[];
  readonly lineStyles: readonly Element[];
  readonly effectStyles: readonly Element[];
  /** `a:bgFillStyleLst`. `bgRef idx` 1001 is entry 0. */
  readonly backgroundFillStyles: readonly Element[];
}

const SCHEME_SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/** The default theme colors, used when a package has no theme part. */
const FALLBACK_COLORS: Record<(typeof SCHEME_SLOTS)[number], string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

const cache = new WeakMap<XmlPart, Theme>();

function fonts(el: Element | undefined): FontCollection {
  const pick = (name: string): string | undefined => {
    const v = attr(child(el, name), 'typeface');
    return v ? v : undefined;
  };
  const latin = pick('a:latin');
  const ea = pick('a:ea');
  const cs = pick('a:cs');
  const scripts = new Map<string, string>();
  for (const f of elements(el)) {
    if (f.local !== 'font' || scripts.size >= MAX_SCRIPT_FONTS) continue;
    const script = attr(f, 'script');
    const typeface = attr(f, 'typeface');
    if (script && typeface && !scripts.has(script)) scripts.set(script, typeface);
  }
  return {
    ...(latin ? { latin } : {}),
    ...(ea ? { ea } : {}),
    ...(cs ? { cs } : {}),
    ...(scripts.size > 0 ? { scripts } : {}),
  };
}

/** Parse a theme part. A missing theme yields the default theme. */
export function parseTheme(part: XmlPart | undefined): Theme {
  if (part) {
    const hit = cache.get(part);
    if (hit) return hit;
  }
  const elements0 = path(part?.root, 'a:themeElements');
  const scheme = child(elements0, 'a:clrScheme');
  const colors = new Map<string, Rgba>();
  const empty: ColorContext = { scheme: new Map(), colorMap: new Map() };
  for (const slot of SCHEME_SLOTS) {
    const resolved = resolveColor(colorChild(child(scheme, `a:${slot}`)), empty);
    colors.set(slot, resolved ?? parseHex(FALLBACK_COLORS[slot])!);
  }
  const fontScheme = child(elements0, 'a:fontScheme');
  const fmt = child(elements0, 'a:fmtScheme');
  const theme: Theme = {
    colors,
    majorFont: fonts(child(fontScheme, 'a:majorFont')),
    minorFont: fonts(child(fontScheme, 'a:minorFont')),
    fillStyles: elements(child(fmt, 'a:fillStyleLst')),
    lineStyles: elements(child(fmt, 'a:lnStyleLst')),
    effectStyles: elements(child(fmt, 'a:effectStyleLst')),
    backgroundFillStyles: elements(child(fmt, 'a:bgFillStyleLst')),
  };
  if (part) cache.set(part, theme);
  return theme;
}

/** Default color map, used when a master has none: bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2. */
export const DEFAULT_COLOR_MAP: ReadonlyMap<string, string> = new Map([
  ['bg1', 'lt1'],
  ['tx1', 'dk1'],
  ['bg2', 'lt2'],
  ['tx2', 'dk2'],
  ['accent1', 'accent1'],
  ['accent2', 'accent2'],
  ['accent3', 'accent3'],
  ['accent4', 'accent4'],
  ['accent5', 'accent5'],
  ['accent6', 'accent6'],
  ['hlink', 'hlink'],
  ['folHlink', 'folHlink'],
]);

/** Read a `p:clrMap` (or `a:overrideClrMapping`) element into a map. */
export function readColorMap(el: Element | undefined): ReadonlyMap<string, string> | undefined {
  if (!el) return undefined;
  const map = new Map(DEFAULT_COLOR_MAP);
  for (const a of el.attrs) if (a.ns === '') map.set(a.local, a.value);
  return map;
}

/**
 * The East Asian (`ea`) or complex-script (`cs`) face of a collection for a language: the
 * slot's own typeface, else the script font of the language's script (ECMA-376 Part 1,
 * 20.1.4.1.16 `font`). Undefined when the collection names neither.
 */
export function collectionScriptFont(
  collection: FontCollection,
  slot: 'ea' | 'cs',
  lang?: string
): string | undefined {
  const own = slot === 'ea' ? collection.ea : collection.cs;
  if (own) return own;
  const script = slot === 'ea' ? eastAsianScriptOf(lang) : complexScriptOf(lang);
  return script ? collection.scripts?.get(script) : undefined;
}

/**
 * Resolve a theme font reference (`+mj-lt`, `+mn-ea`, …) to a typeface. Returns the input
 * unchanged when it is a plain typeface name. `lang` picks the script font of an `ea` or `cs`
 * reference when the theme's slot is empty.
 */
export function resolveThemeFont(
  typeface: string,
  theme: Theme,
  lang?: string
): string | undefined {
  if (!typeface.startsWith('+')) return typeface;
  const collection = typeface.startsWith('+mj')
    ? theme.majorFont
    : typeface.startsWith('+mn')
      ? theme.minorFont
      : null;
  if (!collection) return undefined;
  if (typeface.endsWith('-lt')) return collection.latin;
  if (typeface.endsWith('-ea'))
    return collectionScriptFont(collection, 'ea', lang) || collection.latin;
  if (typeface.endsWith('-cs'))
    return collectionScriptFont(collection, 'cs', lang) || collection.latin;
  return collection.latin;
}
