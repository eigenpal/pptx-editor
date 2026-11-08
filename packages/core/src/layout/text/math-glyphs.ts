// The characters of Office Math runs (ECMA-376 Part 1, 22.1): the math alphanumeric form of a
// letter in each math style, the spacing class of each character, and the spacing between
// classes. Internal to the layout lane.

import type { MathStyle } from '../../cascade/math.ts';

/** The spacing class of a math atom, after the classes of TeX. */
export type AtomClass = 'ord' | 'op' | 'bin' | 'rel' | 'open' | 'close' | 'punct';

const REL = new Set([...'=<>≤≥≠≈≡≅∼≃∝≪≫≺≻⪯⪰∈∉∋⊂⊃⊆⊇⊄⊅⊊⊋∥⊥⊢⊣⊨≔≕≐≜→←↔⇒⇐⇔↦⟶⟵⟷⟹⟸⟺↑↓⇑⇓↗↘↙↖∣∤']);
const BIN = new Set([...'+−-±∓×÷·∘∗⋅∙∪∩∧∨⊕⊖⊗⊘⊙∖⋆⋄◦†‡⊎⊓⊔*']);
const OPEN = new Set([...'([{⟨⌊⌈⟦⟪']);
const CLOSE = new Set([...')]}⟩⌋⌉⟧⟫!']);
const PUNCT = new Set([...',;']);
/** Operators that a run draws larger in display math when they stand alone (n-ary glyphs). */
const LARGE_OPS = new Set([...'∑∏∐∫∬∭∮∯∰⋀⋁⋂⋃⨀⨁⨂⨄⨆']);

/** The class of one character of a math run. */
export function charClass(ch: string): AtomClass {
  if (REL.has(ch)) return 'rel';
  if (BIN.has(ch)) return 'bin';
  if (OPEN.has(ch)) return 'open';
  if (CLOSE.has(ch)) return 'close';
  if (PUNCT.has(ch)) return 'punct';
  if (LARGE_OPS.has(ch)) return 'op';
  return 'ord';
}

/** The integral signs: their limits stand beside them unless `m:limLoc` says otherwise. */
export const INTEGRALS: ReadonlySet<string> = new Set([...'∫∬∭∮∯∰']);

/**
 * The space between two atoms in math units (an eighteenth of an em of the row's size), after
 * TeX's table: thin (3), medium (4), and thick (5). Script rows keep only the thin space of an
 * operator.
 */
export function atomSpace(left: AtomClass, right: AtomClass, script: boolean): number {
  if (left === 'bin' || right === 'bin') {
    if (script) return 0;
    return left === 'open' || right === 'close' || right === 'punct' ? 0 : 4;
  }
  if (left === 'rel' || right === 'rel') {
    if (script) return 0;
    return left === 'open' ||
      right === 'close' ||
      right === 'punct' ||
      (left === 'rel' && right === 'rel')
      ? 0
      : 5;
  }
  if (left === 'op' && (right === 'ord' || right === 'op')) return 3;
  if (right === 'op' && (left === 'ord' || left === 'close')) return 3;
  if (left === 'punct' && !script) return 3;
  return 0;
}

const GREEK_CAPITAL_START = 0x0391;
const GREEK_SMALL_START = 0x03b1;

/** Letter code point → math alphanumeric code point, for one style. */
const BASES: Readonly<
  Record<
    'i' | 'b' | 'bi',
    { latinCap: number; latinSmall: number; greekCap: number; greekSmall: number }
  >
> = {
  i: { latinCap: 0x1d434, latinSmall: 0x1d44e, greekCap: 0x1d6e2, greekSmall: 0x1d6fc },
  b: { latinCap: 0x1d400, latinSmall: 0x1d41a, greekCap: 0x1d6a8, greekSmall: 0x1d6c2 },
  bi: { latinCap: 0x1d468, latinSmall: 0x1d482, greekCap: 0x1d71c, greekSmall: 0x1d736 },
};

/**
 * The character that draws `cp` in a math style with a math font: a Latin or Greek letter maps
 * to its math alphanumeric form (Unicode 3.1, Mathematical Alphanumeric Symbols); other
 * characters stay. Italic leaves Greek capitals upright, as math italic does by default.
 */
export function mathAlphanumeric(cp: number, style: MathStyle): number {
  if (style === 'p') return cp;
  const base = BASES[style];
  if (cp >= 0x41 && cp <= 0x5a) return base.latinCap + (cp - 0x41);
  if (cp >= 0x61 && cp <= 0x7a) {
    // The italic small h is the Planck constant, outside the block.
    if (style === 'i' && cp === 0x68) return 0x210e;
    return base.latinSmall + (cp - 0x61);
  }
  if (cp >= GREEK_CAPITAL_START && cp <= 0x03a9 && cp !== 0x03a2)
    return style === 'i' ? cp : base.greekCap + (cp - GREEK_CAPITAL_START);
  if (cp >= GREEK_SMALL_START && cp <= 0x03c9) return base.greekSmall + (cp - GREEK_SMALL_START);
  return cp;
}

/** Whether a code point is a Latin or Greek letter (which takes the run's math style). */
export function isMathLetter(cp: number): boolean {
  return (
    (cp >= 0x41 && cp <= 0x5a) ||
    (cp >= 0x61 && cp <= 0x7a) ||
    (cp >= GREEK_CAPITAL_START && cp <= 0x03c9 && cp !== 0x03a2)
  );
}

/**
 * The spacing form of a combining accent (`m:acc/m:chr`), so it draws on its own over the base.
 * Undefined for accents that the layout draws as marks (bars and arrows).
 */
export function spacingAccent(chr: string): string | undefined {
  const map: Readonly<Record<string, string>> = {
    '̀': '`',
    '́': '´',
    '̂': 'ˆ',
    '̃': '˜',
    '̆': '˘',
    '̇': '˙',
    '̈': '¨',
    '̊': '˚',
    '̌': 'ˇ',
  };
  if (chr === '̄' || chr === '̅' || chr === '¯' || chr === '⃗' || chr === '⃖' || chr === '⃡')
    return undefined;
  return map[chr] ?? chr;
}

/**
 * The plain letter of a math alphanumeric character, with its style, for a face that has no
 * math alphanumerics: `𝑥` is an italic `x`. Undefined for any other character.
 */
export function plainLetter(
  cp: number
): { readonly cp: number; readonly style: MathStyle } | undefined {
  if (cp === 0x210e) return { cp: 0x68, style: 'i' };
  if (cp >= 0x1d400 && cp <= 0x1d6a3) {
    const block = Math.floor((cp - 0x1d400) / 52);
    const at = (cp - 0x1d400) % 52;
    const letter = at < 26 ? 0x41 + at : 0x61 + (at - 26);
    const style: MathStyle = block === 0 ? 'b' : block === 1 ? 'i' : block === 2 ? 'bi' : 'p';
    return { cp: letter, style };
  }
  if (cp >= 0x1d6a8 && cp <= 0x1d755) {
    const block = Math.floor((cp - 0x1d6a8) / 58);
    const at = (cp - 0x1d6a8) % 58;
    const style: MathStyle = block === 0 ? 'b' : block === 1 ? 'i' : 'bi';
    if (at < 25) return { cp: GREEK_CAPITAL_START + at, style };
    if (at >= 26 && at < 51) return { cp: GREEK_SMALL_START + (at - 26), style };
  }
  return undefined;
}
