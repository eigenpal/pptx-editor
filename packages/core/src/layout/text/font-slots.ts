// The font slot of each character of a run (ECMA-376 Part 1, 21.1.2.3.7 `latin`, 21.1.2.3.3
// `ea`, 21.1.2.3.1 `cs`, 21.1.2.3.10 `sym`): East Asian characters measure and paint with the
// run's `a:ea` face, complex-script characters (right-to-left, Indic, and Southeast Asian
// scripts) with its `a:cs` face, symbol characters (U+F000 to U+F0FF) with its `a:sym` face,
// and every other character with its `a:latin` face. Characters that both East Asian and
// Latin text use (curly quotes, the ellipsis, arrows) take the East Asian face in a run whose
// language is Chinese, Japanese, or Korean. Digits, punctuation, and symbols between
// complex-script characters take the complex-script face. A space takes the Latin font, also between East
// Asian words; between two complex-script characters it takes the complex-script font.
// DOM-free.

import type { RunProps } from '../../cascade/index.ts';
import { eastAsianScriptOf } from '../../cascade/script-fonts.ts';
import { isIdeographic } from './line-break.ts';

/** A run's font slot. */
export type FontSlot = 'latin' | 'ea' | 'cs' | 'sym';

/** Whether a code point is an East Asian character: CJK, kana, bopomofo, Hangul, full width. */
export function isEastAsianChar(cp: number): boolean {
  return (
    isIdeographic(cp) ||
    (cp >= 0x1100 && cp <= 0x11ff) || // Hangul Jamo
    (cp >= 0xa960 && cp <= 0xa97f) || // Hangul Jamo Extended-A
    (cp >= 0xd7b0 && cp <= 0xd7ff) || // Hangul Jamo Extended-B
    (cp >= 0xfe10 && cp <= 0xfe1f) || // vertical forms
    (cp >= 0xfe30 && cp <= 0xfe4f) // CJK compatibility forms
  );
}

/** Whether a code point is in a complex script: right-to-left, Indic, or Southeast Asian. */
export function isComplexScriptChar(cp: number): boolean {
  return (
    (cp >= 0x0590 && cp <= 0x08ff) || // Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan
    (cp >= 0x0900 && cp <= 0x109f) || // Indic scripts, Sinhala, Thai, Lao, Tibetan, Myanmar
    (cp >= 0x1200 && cp <= 0x139f) || // Ethiopic
    (cp >= 0x1780 && cp <= 0x17ff) || // Khmer
    (cp >= 0xfb1d && cp <= 0xfdff) || // Hebrew and Arabic presentation forms
    (cp >= 0xfe70 && cp <= 0xfeff)
  );
}

/** Characters both East Asian and Latin text use; East Asian text draws them full width. */
function isSharedSymbol(cp: number): boolean {
  return (
    cp === 0x00a7 ||
    cp === 0x00a8 ||
    cp === 0x00b0 ||
    cp === 0x00b1 ||
    cp === 0x00b4 ||
    cp === 0x00b6 ||
    cp === 0x00b7 ||
    cp === 0x00d7 ||
    cp === 0x00f7 ||
    (cp >= 0x2010 && cp <= 0x2027) || // dashes, curly quotes, the ellipsis
    (cp >= 0x2030 && cp <= 0x203b) || // per mille, primes, the reference mark
    (cp >= 0x2100 && cp <= 0x214f) || // letterlike symbols
    (cp >= 0x2190 && cp <= 0x23ff) || // arrows, mathematical and technical symbols
    (cp >= 0x2460 && cp <= 0x27bf) // enclosed numbers, box drawing, shapes, dingbats
  );
}

const isSpace = (cp: number): boolean => cp === 0x20 || cp === 0xa0;
const MARK = /^[\p{M}\uFE00-\uFE0F]$/u;
const NEUTRAL = /^[\p{N}\p{P}\p{S}]$/u;

/** The slot of one code point, or null for a space, which takes its neighbors' slot. */
function strongSlot(cp: number, eastAsianText: boolean, symbol: boolean): FontSlot | null {
  if (isSpace(cp)) return null;
  if (symbol && cp >= 0xf000 && cp <= 0xf0ff) return 'sym';
  if (isEastAsianChar(cp)) return 'ea';
  if (isComplexScriptChar(cp)) return 'cs';
  if (eastAsianText && isSharedSymbol(cp)) return 'ea';
  return 'latin';
}

/** One stretch of a text in one font slot: UTF-16 offsets `[start, end)`. */
export interface SlotPiece {
  readonly start: number;
  readonly end: number;
  readonly slot: FontSlot;
}

/**
 * Split `text` into stretches of one font slot. A surrogate pair stays whole; a combining
 * mark or variation selector takes the slot of the character before it.
 */
export function fontSlotPieces(
  text: string,
  props: Pick<RunProps, 'lang' | 'altLang' | 'sym'>
): readonly SlotPiece[] {
  // The `normAutofit` search lays out the same runs at each step: split each text once.
  const kept = piecesOf.get(props);
  if (kept?.text === text) return kept.pieces;
  const pieces = splitSlots(text, props);
  piecesOf.set(props, { text, pieces });
  return pieces;
}

const piecesOf = new WeakMap<object, { text: string; pieces: readonly SlotPiece[] }>();

function splitSlots(text: string, props: Pick<RunProps, 'lang' | 'altLang' | 'sym'>): SlotPiece[] {
  const eastAsianText =
    eastAsianScriptOf(props.lang) !== undefined || eastAsianScriptOf(props.altLang) !== undefined;
  const symbol = props.sym !== undefined;
  // The slot of each code point, at its first UTF-16 unit.
  const slots: (FontSlot | null)[] = new Array<FontSlot | null>(text.length).fill(null);
  const neutral: boolean[] = new Array<boolean>(text.length).fill(false);
  // A combining mark, a variation selector, a ZWJ, and the character after a ZWJ belong to
  // the cluster before them: they take its final slot, after the passes below.
  const follows: boolean[] = new Array<boolean>(text.length).fill(false);
  for (let i = 0, joined = false; i < text.length;) {
    const cp = text.codePointAt(i) ?? 0;
    const len = cp > 0xffff ? 2 : 1;
    let slot = strongSlot(cp, eastAsianText, symbol);
    const mark = cp >= 0x0300 && (cp === 0x200d || MARK.test(String.fromCodePoint(cp)));
    if (i > 0 && (joined || (slot === 'latin' && mark))) {
      follows[i] = true;
      slot = null;
    } else if (slot === 'latin' && NEUTRAL.test(String.fromCodePoint(cp))) {
      neutral[i] = true;
      slot = null;
    }
    slots[i] = slot;
    joined = cp === 0x200d;
    i += len;
  }
  // Digits, punctuation, and symbols between complex-script characters, or between such a
  // character and an edge of the run, take the complex-script slot; any other is Latin
  // (reference PDF of `rtl-text`, slide 2: the digits, full stops, and percent signs in Arabic
  // text). The nearest strong slot on each side skips spaces and other such characters.
  const strongAfter: (FontSlot | null)[] = new Array<FontSlot | null>(text.length).fill(null);
  for (let i = text.length - 1, next: FontSlot | null = null; i >= 0; i -= 1) {
    strongAfter[i] = next;
    if (!neutral[i] && !follows[i] && !isSpace(text.charCodeAt(i))) next = slots[i] ?? next;
  }
  for (let i = 0, before: FontSlot | null = null; i < text.length; i += 1) {
    if (neutral[i]) {
      const next = strongAfter[i];
      const sides = [before, next].filter((side) => side !== null);
      slots[i] = sides.length > 0 && sides.every((side) => side === 'cs') ? 'cs' : 'latin';
    } else if (!isSpace(text.charCodeAt(i))) before = slots[i] ?? before;
  }
  // A space between two complex-script characters takes their slot; any other space is Latin.
  // One pass each way finds the nearest slot before and after every run of spaces, so a long
  // run of spaces costs its length once.
  const after: (FontSlot | null)[] = new Array<FontSlot | null>(text.length).fill(null);
  for (let i = text.length - 1, next: FontSlot | null = null; i >= 0; i -= 1) {
    after[i] = next;
    if (!isSpace(text.charCodeAt(i))) next = slots[i] ?? next;
  }
  for (let i = 0, before: FontSlot | null = null; i < text.length; i += 1) {
    if (slots[i] === null && isSpace(text.charCodeAt(i)))
      // East Asian text takes the Latin space (reference PDF of `east-asian-theme-fonts`).
      slots[i] = before === 'cs' && after[i] === 'cs' ? 'cs' : 'latin';
    else if (!isSpace(text.charCodeAt(i))) before = slots[i] ?? before;
  }
  // One grapheme cluster never splits across faces (a keycap `1️⃣`, an emoji `❤️`).
  for (let i = 1, prev: FontSlot = slots[0] ?? 'latin'; i < text.length; i += 1) {
    if (follows[i]) slots[i] = prev;
    else if (slots[i] !== null) prev = slots[i]!;
  }
  const out: SlotPiece[] = [];
  let pieceStart = 0;
  let pieceSlot: FontSlot | null = null;
  for (let i = 0; i < text.length; i += 1) {
    const slot = slots[i];
    if (slot === null || slot === pieceSlot) continue; // a low surrogate, or the same slot
    if (pieceSlot !== null) out.push({ start: pieceStart, end: i, slot: pieceSlot });
    pieceStart = i;
    pieceSlot = slot;
  }
  if (pieceSlot !== null) out.push({ start: pieceStart, end: text.length, slot: pieceSlot });
  return out;
}

/** The family a run's slot names. */
export function slotFamily(props: RunProps, slot: FontSlot): string {
  if (slot === 'ea') return props.ea;
  if (slot === 'cs') return props.cs;
  if (slot === 'sym') return props.sym ?? props.latin;
  return props.latin;
}
