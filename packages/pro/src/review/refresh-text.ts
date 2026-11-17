/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text comparison of the computed refresh changes (refresh-changes.ts): which paragraphs of a
// shape changed, the changed range of one paragraph widened to whole words, and the formatting
// of a text body without its text. DOM-free.

import { elements, type Element } from '@pptx-editor.dev/core/store';
import { child } from './refresh-common.ts';

/** The most cells of one paragraph alignment table, as `package-diff.ts` limits its own. */
const LCS_CELLS = 250_000;

/** How the paragraphs of a shape changed. Indexes are zero-based. */
export interface ParagraphDiff {
  /** Changed or inserted paragraphs: the new index, and the old index or null for an insertion. */
  readonly changed: readonly (readonly [number, number | null])[];
  /** Old paragraphs without a new paragraph. */
  readonly removed: readonly number[];
  /** Paragraphs with the same text: the new index and the old index. */
  readonly kept: readonly (readonly [number, number])[];
}

/**
 * How the paragraphs changed from `before` to `after`: equal paragraphs in order, and the rest
 * paired by position between them. Null when the alignment would pass the cell limit.
 */
export function diffParagraphs(
  before: readonly string[],
  after: readonly string[]
): ParagraphDiff | null {
  const changed: [number, number | null][] = [];
  const removed: number[] = [];
  const kept: [number, number][] = [];
  // The common ends need no table.
  let head = 0;
  while (head < before.length && head < after.length && before[head] === after[head]) {
    kept.push([head, head]);
    head += 1;
  }
  let tail = 0;
  while (
    tail < before.length - head &&
    tail < after.length - head &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]
  )
    tail += 1;
  const n = before.length - head - tail;
  const m = after.length - head - tail;
  if (n > 0 && m > 0 && (n + 1) * (m + 1) > LCS_CELLS) return null;
  // table[i * w + j]: the longest common run of before[head + i..] and after[head + j..].
  const w = m + 1;
  const table = new Uint32Array(n > 0 && m > 0 ? (n + 1) * w : 0);
  if (table.length > 0)
    for (let i = n - 1; i >= 0; i -= 1)
      for (let j = m - 1; j >= 0; j -= 1)
        table[i * w + j] =
          before[head + i] === after[head + j]
            ? table[(i + 1) * w + j + 1]! + 1
            : Math.max(table[(i + 1) * w + j]!, table[i * w + j + 1]!);
  const at = (i: number, j: number) => (table.length > 0 ? table[i * w + j]! : 0);
  const pending: number[] = [];
  let oldFrom = 0;
  const flush = (oldTo: number) => {
    pending.forEach((nj, k) =>
      changed.push([head + nj, oldFrom + k < oldTo ? head + oldFrom + k : null])
    );
    for (let k = oldFrom + pending.length; k < oldTo; k += 1) removed.push(head + k);
    pending.length = 0;
  };
  let i = 0;
  let j = 0;
  while (i < n || j < m) {
    if (i < n && j < m && before[head + i] === after[head + j]) {
      flush(i);
      kept.push([head + j, head + i]);
      i += 1;
      j += 1;
      oldFrom = i;
    } else if (j < m && (i >= n || at(i, j + 1) >= at(i + 1, j))) {
      pending.push(j);
      j += 1;
    } else i += 1;
  }
  flush(i);
  for (let k = 0; k < tail; k += 1) kept.push([after.length - tail + k, before.length - tail + k]);
  return { changed, removed, kept };
}

/** The changed range of `after` against `before`: the part between their common ends. */
export function changedRange(before: string, after: string): { start: number; end: number } {
  let start = 0;
  while (start < before.length && start < after.length && before[start] === after[start])
    start += 1;
  let tail = 0;
  while (
    tail < before.length - start &&
    tail < after.length - start &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]
  )
    tail += 1;
  return { start, end: after.length - tail };
}

type Segmenter = {
  segment(text: string): Iterable<{ segment: string; index: number; isWordLike?: boolean }>;
};
const segmenter: Segmenter | undefined = (() => {
  const ctor = (Intl as { Segmenter?: new (l?: string, o?: object) => Segmenter }).Segmenter;
  return ctor ? new ctor(undefined, { granularity: 'word' }) : undefined;
})();

const SPACE = /^[\s\p{Z}]+$/u;
/** Punctuation and currency or math signs, which join the word they touch. */
const JOINER = /^[\p{P}\p{Sc}\p{Sm}]+$/u;

/**
 * The words of `text` as `[start, end)` ranges. A word keeps the punctuation and the currency
 * or math signs it touches, so `$4.0M` and `Hello,` are one word each. Two touching words of a
 * script without spaces stay apart, and an emoji is a word of its own. Without
 * `Intl.Segmenter`, a word is a run of characters without spaces.
 */
export function wordRanges(text: string): [number, number][] {
  const out: [number, number][] = [];
  if (!segmenter) {
    for (const m of text.matchAll(/[^\s\p{Z}]+/gu)) out.push([m.index, m.index + m[0].length]);
    return out;
  }
  let prevWord: boolean | undefined;
  for (const seg of segmenter.segment(text)) {
    if (SPACE.test(seg.segment)) {
      prevWord = undefined;
      continue;
    }
    const start = seg.index;
    const end = start + seg.segment.length;
    const word = seg.isWordLike === true;
    const joins = word || JOINER.test(seg.segment);
    const last = out[out.length - 1];
    if (last && last[1] === start && prevWord !== undefined && joins && !(prevWord && word))
      last[1] = end;
    else out.push([start, end]);
    prevWord = joins ? word : undefined;
  }
  return out;
}

const isHigh = (c: number) => c >= 0xd800 && c <= 0xdbff;
const isLow = (c: number) => c >= 0xdc00 && c <= 0xdfff;

/**
 * The range `[start, end)` of `text` widened to whole words. An empty range takes the word
 * it touches. The result never splits a surrogate pair.
 */
export function widenRange(
  text: string,
  start: number,
  end: number
): { start: number; end: number } {
  let s = start;
  let e = end;
  for (const [ws, we] of wordRanges(text)) {
    if (we < start || ws > end) continue;
    if (start === end) {
      if (ws <= start && start <= we) {
        s = Math.min(s, ws);
        e = Math.max(e, we);
      }
      continue;
    }
    if (ws <= start && start < we) s = Math.min(s, ws);
    if (ws < end && end <= we) e = Math.max(e, we);
  }
  if (s > 0 && s < text.length && isLow(text.charCodeAt(s)) && isHigh(text.charCodeAt(s - 1)))
    s -= 1;
  if (e > 0 && e < text.length && isLow(text.charCodeAt(e)) && isHigh(text.charCodeAt(e - 1)))
    e += 1;
  return { start: s, end: e };
}

/** An element without its text: names, attributes, and child elements other than `t`. */
function strip(node: Element, depth = 0): string {
  if (depth > 64) return '';
  return (
    `<${node.local} ${node.attrs.map((x) => `${x.local}=${x.value}`).join(' ')}>` +
    elements(node)
      .filter((c) => c.local !== 't')
      .map((c) => strip(c, depth + 1))
      .join('')
  );
}

/** The formatting of a shape's text: the body without paragraphs, and each paragraph. */
export interface TextFormat {
  readonly body: string;
  readonly paragraphs: readonly string[];
}

/** The formatting of a shape's text body without its text, or undefined without a body. */
export function textFormat(el: Element): TextFormat | undefined {
  const body = child(el, 'txBody');
  if (!body) return undefined;
  const paragraphs: string[] = [];
  const rest: Element = {
    ...body,
    children: body.children.filter((c) => {
      if (c.type !== 'element' || c.kind !== 'paragraph') return true;
      paragraphs.push(strip(c));
      return false;
    }),
  };
  return { body: strip(rest), paragraphs };
}

/**
 * Whether the formatting is the same for the paragraphs that `diff` pairs, and for the body.
 * Inserted and removed paragraphs do not count: their text change reports them.
 */
export function sameFormat(
  a: TextFormat | undefined,
  b: TextFormat | undefined,
  diff: ParagraphDiff | null
): boolean {
  if (!a || !b) return a === b;
  if (a.body !== b.body) return false;
  if (!diff) return a.paragraphs.join('') === b.paragraphs.join('');
  const pairs = [...diff.kept, ...diff.changed];
  return pairs.every(([nj, oi]) => oi === null || a.paragraphs[oi] === b.paragraphs[nj]);
}
