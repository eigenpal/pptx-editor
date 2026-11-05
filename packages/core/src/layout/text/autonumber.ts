// `a:buAutoNum` schemes (ECMA-376 ST_TextAutonumberScheme): the number format plus the
// punctuation around it, and the counters that number a text body's paragraphs.

import type { ResolvedParagraph } from '../../cascade/index.ts';
import { mapSymbolBulletText } from './symbol-encoding.ts';

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ];
  let out = '';
  let rest = Math.min(Math.max(1, n), 3999);
  for (const [v, s] of table)
    while (rest >= v) {
      out += s;
      rest -= v;
    }
  return out;
}

function alpha(n: number): string {
  let out = '';
  let rest = Math.max(1, n);
  while (rest > 0) {
    rest -= 1;
    out = String.fromCharCode(97 + (rest % 26)) + out;
    rest = Math.floor(rest / 26);
  }
  return out;
}

/**
 * Whether a bullet font has a symbol character map (`charset="2"`: Wingdings, Webdings,
 * Symbol): its characters are code points of the font, not Unicode text. @internal
 */
export function isSymbolFont(family: string): boolean {
  return /^(wingdings( [23])?|webdings|symbol)$/i.test(family);
}

/**
 * Wingdings bullets of the Bullets gallery that DOCX Editor's table
 * (`symbol-encoding.ts`, a copy kept equal to its source) does not list, each compared with the
 * glyph the face draws: the Star bullet, `v` (0x76).
 */
const GALLERY_TWINS: ReadonlyMap<string, string> = new Map([['v', '\u2756']]);

/**
 * A bullet character as a face other than its symbol font draws it: the Unicode twin of a
 * Wingdings or Symbol byte, or the character itself. `isFamilyAvailable` returns true when the
 * symbol font itself draws, and the character is kept. @internal
 */
export function symbolBulletText(
  char: string,
  font: string | null | undefined,
  isFamilyAvailable?: (family: string) => boolean
): string {
  const twin = font?.toLowerCase() === 'wingdings' ? GALLERY_TWINS.get(char) : undefined;
  if (twin !== undefined && isFamilyAvailable?.(font!) !== true) return twin;
  return mapSymbolBulletText(char, font, isFamilyAvailable);
}

/**
 * Whether `autonumberText` draws a scheme exactly (Latin letters, Arabic and Roman numbers).
 * @internal
 */
export function isExactScheme(scheme: string): boolean {
  return /^(alphaLc|alphaUc|arabic|romanLc|romanUc)(ParenBoth|ParenR|Period|Plain)$/.test(scheme);
}

/** The marker text for item `n` (1-based) of a scheme. Unknown schemes fall back to `n.`. */
export function autonumberText(scheme: string, n: number): string {
  const m = /^(alphaLc|alphaUc|arabic|romanLc|romanUc)(.*)$/.exec(scheme);
  if (!m) return `${n}.`;
  const [, kind, punct] = m;
  const core =
    kind === 'alphaLc'
      ? alpha(n)
      : kind === 'alphaUc'
        ? alpha(n).toUpperCase()
        : kind === 'romanLc'
          ? roman(n)
          : kind === 'romanUc'
            ? roman(n).toUpperCase()
            : String(n);
  switch (punct) {
    case 'ParenBoth':
      return `(${core})`;
    case 'ParenR':
      return `${core})`;
    case 'Period':
      return `${core}.`;
    case 'Plain':
      return core;
    case 'Minus':
      return `${core} -`;
    case 'Colon':
      return `${core}:`;
    default:
      return `${core}.`;
  }
}

interface Counter {
  readonly scheme: string;
  readonly startAt: number;
  readonly n: number;
}

/**
 * Autonumber counters for one text body, one per outline level. A numbered paragraph continues
 * its level's sequence when the scheme and `startAt` match the previous number on that level,
 * and restarts at `startAt` otherwise. Any paragraph with text ends the sequences of every
 * deeper level, so sub-items restart under each parent, and a paragraph without a number ends
 * the sequence on its own level. Empty paragraphs show no bullet and change nothing.
 */
export class AutonumberCounters {
  readonly #levels = new Map<number, Counter>();

  /** The number an empty numbered paragraph would show, without advancing any sequence. */
  peek(p: ResolvedParagraph): number | undefined {
    const b = p.bullet;
    if (!b || b.type !== 'autonum') return undefined;
    const prev = this.#levels.get(p.level);
    return prev && prev.scheme === b.scheme && prev.startAt === b.startAt ? prev.n + 1 : b.startAt;
  }

  /** Advance for a paragraph with text. The item number for a numbered paragraph. */
  next(p: ResolvedParagraph): number | undefined {
    for (const level of [...this.#levels.keys()]) if (level > p.level) this.#levels.delete(level);
    const b = p.bullet;
    if (!b || b.type !== 'autonum') {
      this.#levels.delete(p.level);
      return undefined;
    }
    const prev = this.#levels.get(p.level);
    const n =
      prev && prev.scheme === b.scheme && prev.startAt === b.startAt ? prev.n + 1 : b.startAt;
    this.#levels.set(p.level, { scheme: b.scheme, startAt: b.startAt, n });
    return n;
  }
}
