// AutoCorrect while typing: curly quotes, symbols, fractions, dashes, capitals, and lists.
//
// Each rule reads the paragraph text before the caret and acts when a key ends a token:
// - A straight quote becomes a curly one: open at the paragraph start, after a space or an
//   opening bracket; close (and an apostrophe) elsewhere.
// - A letter c, r, or tm in parentheses becomes the ©, ®, or ™ sign
//   on the closing parenthesis; a third full stop makes an ellipsis.
// - On a space, punctuation, or Enter: `1/2`, `1/4`, and `3/4` become ½, ¼, and ¾; ` -- `
//   between words becomes an en dash; the first word of a sentence or a paragraph, the word
//   `i`, and the names of the days take a capital first letter.
// - A space after `1.`, `1)`, `*`, or `-` at the start of a paragraph with no bullet turns the
//   paragraph into a numbered or bulleted list. The caller commits that as its own undo step,
//   so Backspace or Undo right after it gives the typed text back.
//
// Typing only: a paste or an IME commit of more than one character is never corrected.

import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import {
  comparePoints,
  deleteRange,
  insertText,
  paragraphText,
  paragraphsOf,
  slideNumberOf,
  slotsOf,
  splitParagraph,
} from '../automation/commands/text-model.ts';
import { editText, typedCaret, type TextInputKind } from './text-input.ts';
import { setListKind } from '../automation/commands/list-commands.ts';
import { attr, child } from '../store/index.ts';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** Keys that end a token: AutoCorrect looks at the word before them. */
const TRIGGERS = new Set([' ', '\n', '.', ',', ';', ':', '!', '?', ')', '"', "'"]);
const OPENERS = /[\s([{“‘—–-]$/u;
const SYMBOLS: readonly (readonly [string, string])[] = [
  ['(c', '©'],
  ['(r', '®'],
  ['(tm', '™'],
];
const FRACTIONS: Readonly<Record<string, string>> = { '1/2': '½', '1/4': '¼', '3/4': '¾' };
const DAYS = new Set([
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
  'sunday',
]);
/** Keys after which a lone `i` becomes `I`. */
const I_TRIGGERS = new Set([' ', '\n', ',', ';', ':', '!', '?', "'", '\u2019']);
/** Common abbreviations whose full stop does not end a sentence. */
const ABBREVIATIONS = new Set([
  'approx',
  'cf',
  'dept',
  'dr',
  'etc',
  'fig',
  'inc',
  'jr',
  'ltd',
  'min',
  'mr',
  'mrs',
  'ms',
  'no',
  'p',
  'pp',
  'sr',
  'st',
  'vs',
]);
/** The list a paragraph start turns into, by its typed marker. */
const LIST_MARKERS: Readonly<Record<string, 'bullet' | 'arabicPeriod' | 'arabicParenR'>> = {
  '*': 'bullet',
  '-': 'bullet',
  '1.': 'arabicPeriod',
  '1)': 'arabicParenR',
};

/** One replacement in the paragraph: `[from, to)` becomes `text`. */
interface Edit {
  readonly from: number;
  readonly to: number;
  readonly text: string;
}

/** The character a typed quote becomes, by the text before it. */
export function smartQuote(before: string, typed: string): string {
  if (typed !== '"' && typed !== "'") return typed;
  const open = before === '' || OPENERS.test(before);
  if (typed === '"') return open ? '“' : '”';
  return open ? '‘' : '’';
}

/**
 * The replacements a typed `trigger` makes in `before`, the paragraph text before the caret,
 * and the text that goes in for the trigger itself. Pure, for tests.
 */
export function corrections(
  before: string,
  trigger: string
): { readonly edits: readonly Edit[]; readonly typed: string } {
  const edits: Edit[] = [];
  const typed = trigger === '\n' ? '' : smartQuote(before, trigger);
  if (trigger === ')') {
    for (const [tail, symbol] of SYMBOLS)
      if (before.toLowerCase().endsWith(tail))
        return {
          edits: [{ from: before.length - tail.length, to: before.length, text: '' }],
          typed: symbol,
        };
  }
  if (trigger === '.' && before.endsWith('..'))
    return { edits: [{ from: before.length - 2, to: before.length, text: '' }], typed: '…' };
  if (!TRIGGERS.has(trigger)) return { edits, typed };
  // The last word, by a scan back to the space before it. `/(\S+)$/u` retried from each
  // position, which made a space after a long word quadratic.
  let start = before.length;
  while (start > 0 && !/\s/u.test(before[start - 1]!)) start -= 1;
  const word = before.slice(start);
  if (Object.hasOwn(FRACTIONS, word))
    edits.push({ from: start, to: before.length, text: FRACTIONS[word] });
  if (
    word === '--' &&
    (trigger === ' ' || trigger === '\n') &&
    /\S\s+$/u.test(before.slice(0, start))
  )
    edits.push({ from: start, to: before.length, text: '–' });
  const letters = /^[\p{Ll}][\p{Ll}'’]*$/u.exec(word.replace(/^["'“‘([]+/u, ''));
  if (letters && edits.length === 0) {
    const at = before.length - letters[0].length;
    if (capitalized(letters[0], before.slice(0, start), trigger))
      edits.push({ from: at, to: at + 1, text: letters[0][0].toUpperCase() });
  }
  return { edits, typed };
}

/**
 * Whether a lower-case `word` after `prefix` (text up to the token) takes a capital when
 * `trigger` ends it. The word `i` takes one only before a space or punctuation that ends a
 * word, so `i.e.` and `(i)` stay as typed.
 */
function capitalized(word: string, prefix: string, trigger: string): boolean {
  if (word === 'i') return I_TRIGGERS.has(trigger);
  if (DAYS.has(word)) return true;
  // `trimEnd` removes what `\s` matches, without a regex retry from each space.
  const head = prefix.trimEnd();
  if (head === '') return true;
  if (!/[.!?]$/u.test(head) || head === prefix) return false;
  // An abbreviation (`e.g.`, `approx.`), a lone initial, or a number (`Step 3.`) does not end
  // a sentence.
  let from = head.length;
  while (from > 0 && !/\s/u.test(head[from - 1]!)) from -= 1;
  const last = head.slice(from);
  if (/^(?:\p{L}\.)+$/u.test(last) || /^\d+[.)]$/u.test(last)) return false;
  return !ABBREVIATIONS.has(last.slice(0, -1).toLowerCase());
}

/** What `autoCorrectEdit` did: the caret after it, and whether it started a list. */
export interface AutoCorrectResult {
  readonly at: TextPoint;
  readonly list: boolean;
}

/**
 * Type one character at a collapsed caret with AutoCorrect: the corrections before the caret,
 * then the typed character (or what it became), then a list when the paragraph start asks for
 * one. Edits go right to left, so the offsets before them stay valid.
 */
export function autoCorrectEdit(
  batch: OpBatch,
  part: string,
  slide: number,
  target: { readonly shapeId: string; readonly txBodyId: string },
  item: ShapeItem,
  caret: TextPoint,
  trigger: string
): AutoCorrectResult {
  const p = paragraphsOf(batch.element(part, target.txBodyId))[caret.paragraph];
  const before = p ? paragraphText(p, slideNumberOf(batch, part)).slice(0, caret.offset) : '';
  const { edits, typed } = corrections(before, trigger);
  // A field (a slide number, a date) is never rewritten: an edit that reaches one is skipped.
  const fields = p ? slotsOf(p, slideNumberOf(batch, part)).filter((x) => x.kind === 'field') : [];
  const clear = (e: Edit) => !fields.some((f) => f.start < e.to && f.end > e.from);
  let at = caret;
  for (const e of [...edits].filter(clear).sort((a, b) => b.from - a.from)) {
    const point = (offset: number) => ({ paragraph: caret.paragraph, offset });
    // The new text goes in at the end of the old, in its run, so it keeps that formatting.
    if (e.text) insertText(batch, part, target.txBodyId, point(e.to), e.text);
    deleteRange(batch, part, target.txBodyId, point(e.from), point(e.to));
    at = point(at.offset - (e.to - e.from) + e.text.length);
  }
  // A symbol whose deletion a field blocked goes in as the key typed it.
  const typedText = edits.every(clear) || trigger === '\n' ? typed : trigger;
  if (typedText) at = insertText(batch, part, target.txBodyId, at, typedText);
  const marker = trigger === ' ' && !item.element.text?.paragraphs[caret.paragraph]?.bullet;
  const scheme = marker && Object.hasOwn(LIST_MARKERS, before) ? LIST_MARKERS[before] : null;
  if (!scheme) return { at, list: false };
  deleteRange(batch, part, target.txBodyId, { paragraph: caret.paragraph, offset: 0 }, at);
  const list = { ...target, paragraphs: [caret.paragraph] };
  setListKind(batch, part, slide, list, scheme === 'bullet' ? 'bullet' : 'number');
  if (scheme === 'arabicParenR') {
    const pPr = child(paragraphsOf(batch.element(part, target.txBodyId))[caret.paragraph], 'a:pPr');
    const auto = pPr && child(pPr, 'a:buAutoNum');
    if (auto && attr(auto, 'type') !== scheme)
      batch.push({
        op: 'setAttribute',
        part,
        node: auto.id,
        name: { ns: '', local: 'type' },
        value: scheme,
      });
  }
  return { at: { paragraph: caret.paragraph, offset: 0 }, list: true };
}

/**
 * A typed character (at a caret or over a selection) or Enter at a caret, with AutoCorrect, or
 * null when AutoCorrect does not apply (a paste, a tab, or another kind of input): the caller
 * edits as usual then.
 */
export function autoCorrectInput(
  batch: OpBatch,
  part: string,
  slide: number,
  txBodyId: string,
  item: ShapeItem,
  sel: TextSelection,
  kind: string,
  data: string | undefined
): AutoCorrectResult | null {
  const target = { shapeId: sel.nodeId, txBodyId };
  const collapsed = comparePoints(sel.anchor, sel.focus) === 0;
  if (kind === 'paragraph') {
    if (!collapsed) return null;
    const fixed = autoCorrectEdit(batch, part, slide, target, item, sel.focus, '\n');
    return { at: splitParagraph(batch, part, txBodyId, fixed.at), list: false };
  }
  // One typed character: not a tab (a list key types it), a control, or a pasted run.
  if (kind !== 'text' || !data || [...data].length !== 1 || data < ' ' || data === '\u007F')
    return null;
  // A character typed over a selection replaces it, then corrects as at a caret.
  const at = collapsed ? sel.focus : deleteRange(batch, part, txBodyId, sel.anchor, sel.focus);
  return autoCorrectEdit(batch, part, slide, target, item, at, data);
}

/**
 * The text input function of an editor: each input with AutoCorrect when `enabled`, else (or
 * when AutoCorrect does not apply) as typed (`editText`).
 */
export function withAutoCorrect(enabled: boolean) {
  return (
    batch: OpBatch,
    part: string,
    slide: number,
    txBodyId: string,
    item: ShapeItem,
    sel: TextSelection,
    kind: TextInputKind,
    data: string | undefined
  ): AutoCorrectResult => {
    const fixed = enabled
      ? autoCorrectInput(batch, part, slide, txBodyId, item, sel, kind, data)
      : null;
    const done = fixed ?? {
      at: editText(batch, part, txBodyId, item, sel, kind, data),
      list: false,
    };
    return kind === 'text' ? { ...done, at: typedCaret(item, sel.focus.paragraph, done.at) } : done;
  };
}
