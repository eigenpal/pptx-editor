// Character formatting beyond the bold-to-strike toggles: superscript and subscript, Change
// Case, character spacing, and the Font dialog (Format > Font), as store ops.
//
// ECMA-376 Part 1, DrawingML `CT_TextCharacterProperties` (`a:rPr`): `baseline` (`ST_Percentage`,
// thousandths of a percent of the font size), `spc` (`ST_TextPointUnqualified`, hundredths of a
// point, -400000 to 400000), `kern` (`ST_TextNonNegativePoint`, 0 to 400000), `cap`
// (`ST_TextCapsType`), `strike` (`ST_TextStrikeType`), `u` (`ST_TextUnderlineType`), and the
// underline color (`EG_TextUnderlineFill`: `a:uFill`, or `a:uFillTx` to follow the text). Change
// Case edits the text (`a:t`) of the runs, never a field.

import { child, type Element } from '../../store/index.ts';
import { toHex, type RunProps } from '../../cascade/index.ts';
import type {
  CapsStyle,
  EditorCommand,
  FontFormatCommand,
  FontFormatRead,
  StrikeStyle,
  TextCase,
  TextPoint,
  UnderlineStyle,
} from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';
import { colorXml, isColorChoice } from './style-commands.ts';
import {
  aSpec,
  comparePoints,
  paragraphsOf,
  setRunAttr,
  setRunChild,
  slideNumberOf,
  slotsOf,
} from './text-model.ts';
import { commandWordAt } from './command-word.ts';

/** What the Superscript and Subscript buttons write (`baseline`). */
export const SUPERSCRIPT_BASELINE = 30;
export const SUBSCRIPT_BASELINE = -25;

export const TEXT_CASES: readonly TextCase[] = [
  'sentence',
  'lower',
  'upper',
  'capitalize',
  'toggle',
];

/** Every underline style, in the order the Font dialog lists them. @public */
export const UNDERLINE_STYLES: readonly UnderlineStyle[] = [
  'none',
  'words',
  'single',
  'double',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDouble',
];

/** The underline words whose `ST_TextUnderlineType` token differs; the others are the token. */
const UNDERLINE_TOKENS: Readonly<Partial<Record<UnderlineStyle, string>>> = {
  single: 'sng',
  double: 'dbl',
  wavyDouble: 'wavyDbl',
};
/** `u` for an underline style, and the style for a `u` token (unknown tokens read as single). */
const underlineToken = (style: UnderlineStyle): string => UNDERLINE_TOKENS[style] ?? style;
const underlineStyleOf = (token: string): UnderlineStyle =>
  (Object.entries(UNDERLINE_TOKENS).find(([, t]) => t === token)?.[0] as UnderlineStyle) ??
  (UNDERLINE_STYLES.includes(token as UnderlineStyle) ? (token as UnderlineStyle) : 'single');

const STRIKES: Readonly<Record<StrikeStyle, string>> = {
  none: 'noStrike',
  single: 'sngStrike',
  double: 'dblStrike',
};
const CAPS: readonly CapsStyle[] = ['none', 'small', 'all'];
/** The limits: ±1584 pt of spacing, a ±100 % offset, kerning up to 1584 pt. */
const MAX_POINTS = 1584;
const MAX_OFFSET = 100;
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
/**
 * A font family name the editor writes or offers: 1 to 64 letters, digits, marks, spaces, and
 * `-._+`. A family name can reach a CSS declaration, so every write and the font list use it.
 */
export const FONT_NAME = /^[\p{L}\p{N}\p{M} \-.+_]{1,64}$/u;

/**
 * The Latin theme font references a run can name: the theme's major (headings) and minor (body)
 * font. ECMA-376 Part 1, 20.1.4.1.17 `latin`; the text follows a later theme change. @internal
 */
export const LATIN_THEME_FONTS = { major: '+mj-lt', minor: '+mn-lt' } as const;

/**
 * A typeface the Font controls write to `a:latin`: a family name, or a Latin theme font
 * reference. Any other name that starts with `+` is refused, as it would not resolve.
 */
export function isLatinTypeface(typeface: string): boolean {
  if (typeface.startsWith('+'))
    return typeface === LATIN_THEME_FONTS.major || typeface === LATIN_THEME_FONTS.minor;
  return FONT_NAME.test(typeface);
}

/**
 * Write one family to the Latin, East Asian, and complex script faces of a run
 * (ECMA-376 Part 1, 21.1.2.3.7 `ea`, 21.1.2.3.1 `cs`), so text in every script takes the
 * family. A Latin theme font reference maps to the reference of the same slot
 * (`+mj-lt` to `+mj-ea` and `+mj-cs`). @internal
 */
export function setRunTypeface(batch: OpBatch, part: string, run: Element, typeface: string): void {
  for (const slot of ['latin', 'ea', 'cs'] as const) {
    const face = typeface.startsWith('+')
      ? typeface.replace(/-lt$/, `-${slot === 'latin' ? 'lt' : slot}`)
      : typeface;
    setRunChild(batch, part, run, [slot], aSpec(slot, { typeface: face }));
  }
}

const finite = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);

/** Check a character spacing in points, or refuse it. */
export function checkCharacterSpacing(points: unknown): number {
  if (!finite(points) || Math.abs(points) > MAX_POINTS)
    throw new CommandRefusal('errors.unsupported', 'spacing');
  return points;
}

/** The font family a command sets (Font, or the Font dialog), for the host font provider. */
export function fontFamilyOf(command: EditorCommand): string | undefined {
  const family =
    command.type === 'setFontFamily'
      ? command.family
      : command.type === 'setFontFormat'
        ? command.fontFamily
        : undefined;
  // A theme font reference names no family to load: the deck's theme fonts load with the deck.
  const name = typeof family === 'string' ? family.trim() : '';
  return name && !name.startsWith('+') ? name : undefined;
}

/** Check the Font dialog's fields, or refuse them. A command with no field is refused too. */
export function checkFontFormat(f: FontFormatCommand): void {
  const bad = (name: string): never => {
    throw new CommandRefusal('errors.unsupported', name);
  };
  const keys: (keyof FontFormatCommand)[] = [
    'fontFamily',
    'fontSizePt',
    'bold',
    'italic',
    'color',
    'underline',
    'strike',
    'baseline',
    'caps',
    'characterSpacingPt',
    'kerningPt',
    'underlineColor',
  ];
  if (!keys.some((k) => f[k] !== undefined)) bad('font');
  if (
    f.fontFamily !== undefined &&
    (typeof f.fontFamily !== 'string' || !isLatinTypeface(f.fontFamily.trim()))
  )
    bad('family');
  if (
    f.fontSizePt !== undefined &&
    !(finite(f.fontSizePt) && f.fontSizePt >= 1 && f.fontSizePt <= 4000)
  )
    bad('fontSizePt');
  if (f.bold !== undefined && typeof f.bold !== 'boolean') bad('bold');
  if (f.italic !== undefined && typeof f.italic !== 'boolean') bad('italic');
  if (f.color !== undefined && f.color !== null && !isColorChoice(f.color)) bad('color');
  if (
    f.underlineColor !== undefined &&
    f.underlineColor !== null &&
    !isColorChoice(f.underlineColor)
  )
    bad('underlineColor');
  if (f.underline !== undefined && !UNDERLINE_STYLES.includes(f.underline)) bad('underline');
  if (f.strike !== undefined && !Object.hasOwn(STRIKES, f.strike)) bad('strike');
  if (f.baseline !== undefined && !(finite(f.baseline) && Math.abs(f.baseline) <= MAX_OFFSET))
    bad('baseline');
  if (f.caps !== undefined && !CAPS.includes(f.caps)) bad('caps');
  if (f.characterSpacingPt !== undefined) checkCharacterSpacing(f.characterSpacingPt);
  if (
    f.kerningPt !== undefined &&
    f.kerningPt !== null &&
    // 0 is off (`kern="0"`), so a kerning size is one hundredth or more
    // after the rounding that `spacingValue` writes.
    !(finite(f.kerningPt) && Math.round(f.kerningPt * 100) >= 1 && f.kerningPt <= MAX_POINTS)
  )
    bad('kerningPt');
}

/**
 * `baseline` for an offset in percent. Off writes `baseline="0"`, as `b="0"` turns bold off, so
 * an offset the text inherits (a list style's superscript) is turned off too.
 */
export function baselineValue(percent: number): string {
  return String(Math.round(percent * 1000));
}

/** `spc` for a spacing in points. */
export function spacingValue(points: number): string {
  return String(Math.round(points * 100));
}

/** Write the Font dialog's fields on one run-like element (`a:r`, `a:fld`, `a:endParaRPr`, …). */
export function applyFontFormat(
  batch: OpBatch,
  part: string,
  run: Element,
  f: FontFormatCommand
): void {
  const set = (name: string, value: string | null): void =>
    setRunAttr(batch, part, run, name, value);
  if (f.fontFamily !== undefined)
    setRunChild(batch, part, run, ['latin'], aSpec('latin', { typeface: f.fontFamily.trim() }));
  if (f.fontSizePt !== undefined) set('sz', String(Math.round(f.fontSizePt * 100)));
  if (f.bold !== undefined) set('b', f.bold ? '1' : '0');
  if (f.italic !== undefined) set('i', f.italic ? '1' : '0');
  if (f.color !== undefined)
    setRunChild(
      batch,
      part,
      run,
      FILLS,
      f.color === null ? null : aSpec('solidFill', {}, [specFromXml(colorXml(f.color))])
    );
  if (f.underline !== undefined) set('u', underlineToken(f.underline));
  if (f.strike !== undefined) set('strike', STRIKES[f.strike]);
  if (f.baseline !== undefined) set('baseline', baselineValue(f.baseline));
  if (f.caps !== undefined) set('cap', f.caps);
  if (f.characterSpacingPt !== undefined) set('spc', spacingValue(f.characterSpacingPt));
  // "Kerning for fonts" off writes `kern="0"`.
  if (f.kerningPt !== undefined)
    set('kern', f.kerningPt === null ? '0' : spacingValue(f.kerningPt));
  if (f.underlineColor !== undefined)
    setRunChild(
      batch,
      part,
      run,
      ['uFillTx', 'uFill'],
      // `a:uFillTx` makes the underline follow the text, also over a list style's `a:uFill`.
      f.underlineColor === null
        ? aSpec('uFillTx')
        : aSpec('uFill', {}, [aSpec('solidFill', {}, [specFromXml(colorXml(f.underlineColor))])])
    );
}

const strikeOf = (raw: string): StrikeStyle =>
  raw === 'dblStrike' ? 'double' : raw === 'sngStrike' ? 'single' : 'none';

/** One run's values as the Font dialog shows them. */
function fontValuesOf(props: RunProps) {
  const underline = underlineStyleOf(props.underline);
  return {
    fontFamily: props.latin,
    fontSizePt: props.size / 100,
    bold: props.bold,
    italic: props.italic,
    underline,
    strike: strikeOf(props.strike),
    baseline: props.baseline,
    caps: props.caps,
    characterSpacingPt: props.spacing / 100,
    kerningPt: props.kern !== undefined && props.kern > 0 ? props.kern / 100 : null,
    underlineColor:
      props.underlineFill?.type === 'solid' ? toHex(props.underlineFill.color).toUpperCase() : null,
  };
}

/**
 * The Font dialog's view of the selected runs: each field holds the value they share, or null
 * when they disagree. Null for no run.
 */
export function fontFormatRead(runs: readonly RunProps[]): FontFormatRead | null {
  const values = runs.map(fontValuesOf);
  const first = values[0];
  if (!first) return null;
  const keys = Object.keys(first) as (keyof typeof first)[];
  const disagrees = Object.fromEntries(
    keys.map((k) => [k, values.some((v) => v[k] !== first[k])])
  ) as FontFormatRead['disagrees'];
  const read = Object.fromEntries(keys.map((k) => [k, disagrees[k] ? null : first[k]]));
  return { ...(read as Omit<FontFormatRead, 'disagrees'>), disagrees };
}

/** Whether two reads show the same thing, for adapters that keep the last read. @internal */
export function sameFontFormat(a: FontFormatRead | null, b: FontFormatRead | null): boolean {
  return a === b || (a !== null && b !== null && JSON.stringify(a) === JSON.stringify(b));
}

const LETTER = /\p{L}/u;
const WORD_PART = /[\p{L}\p{N}\p{M}]/u;
/** An apostrophe inside a word ("don't"), not one that opens a quote. */
const APOSTROPHE = /['’]/u;

/** One character in another case, or the character itself when the case would change its length. */
function recase(ch: string, upper: boolean): string {
  const out = upper ? ch.toUpperCase() : ch.toLowerCase();
  return out.length === ch.length ? out : ch;
}

/**
 * `text` with the letters in `[from, to)` changed to `mode`. Sentences and words are found in the
 * whole text, so a range that starts mid-sentence continues that sentence. The length never
 * changes, so selection offsets stay valid.
 */
export function changeCaseText(text: string, from: number, to: number, mode: TextCase): string {
  let out = '';
  let sentenceStart = true;
  let afterStop = false;
  // Whether the character before belongs to a word: a letter, digit, or mark, or an apostrophe after one.
  let inWord = false;
  for (let i = 0; i < text.length;) {
    const ch = String.fromCodePoint(text.codePointAt(i)!);
    const letter = LETTER.test(ch);
    let next = ch;
    if (letter && i >= from && i < to) {
      const wordStart = !inWord;
      next =
        mode === 'upper'
          ? recase(ch, true)
          : mode === 'lower'
            ? recase(ch, false)
            : mode === 'toggle'
              ? recase(ch, ch === ch.toLowerCase())
              : mode === 'capitalize'
                ? recase(ch, wordStart)
                : recase(ch, sentenceStart);
    }
    if (letter || /\p{N}/u.test(ch)) {
      sentenceStart = false;
      afterStop = false;
    } else if (/[.!?]/.test(ch)) afterStop = true;
    else if (/\s/u.test(ch) && afterStop) sentenceStart = true;
    out += next;
    inWord = WORD_PART.test(ch) || (inWord && APOSTROPHE.test(ch));
    i += ch.length;
  }
  return out;
}

/**
 * Change Case in one text body: the selected range, the word at or next to a collapsed caret
 * (refused with `disabled.selectWord` between words), or the whole body when the range is absent (the
 * shape is selected). Returns whether any text changed.
 */
export function changeCase(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  range: { from: TextPoint; to: TextPoint } | undefined,
  mode: TextCase
): boolean {
  if (!TEXT_CASES.includes(mode)) throw new CommandRefusal('errors.unsupported', 'case');
  const n = slideNumberOf(batch, part);
  const paragraphs = paragraphsOf(batch.element(part, txBodyId));
  // No range (the shape is selected): the whole body. At a caret: the word around it; a
  // caret outside a word has nothing to change.
  let span = range;
  if (span && comparePoints(span.from, span.to) === 0) {
    // At a word's end the word before the caret, at its start the word after (`command-word.ts`).
    const word = commandWordAt(batch, part, txBodyId, span.from);
    if (!word) throw new CommandRefusal('disabled.selectWord');
    span = word;
  }
  const [lo, hi] = !span
    ? [
        { paragraph: 0, offset: 0 },
        { paragraph: paragraphs.length - 1, offset: Number.MAX_SAFE_INTEGER },
      ]
    : comparePoints(span.from, span.to) <= 0
      ? [span.from, span.to]
      : [span.to, span.from];
  let changed = false;
  for (let pi = lo.paragraph; pi <= hi.paragraph && pi < paragraphs.length; pi += 1) {
    const slots = slotsOf(paragraphs[pi], n);
    const text = slots.map((s) => s.text).join('');
    const from = pi === lo.paragraph ? lo.offset : 0;
    const to = pi === hi.paragraph ? Math.min(hi.offset, text.length) : text.length;
    const next = changeCaseText(text, from, to, mode);
    for (const s of slots) {
      if (s.kind !== 'run' || s.end <= from || s.start >= to) continue;
      const replaced = next.slice(s.start, s.end);
      const t = child(s.el, 'a:t');
      if (!t || replaced === s.text) continue;
      batch.push({ op: 'setText', part, node: t.id, text: replaced });
      changed = true;
    }
  }
  return changed;
}
