// Animation targets that name paragraphs follow their paragraphs through text edits.
//
// An effect can animate some paragraphs of a shape: `p:spTgt` holds a `p:txEl` with a `p:pRg`
// (ECMA-376 Part 1, §19.5.72 and §19.5.63), whose `st` and `end` are paragraph indexes. An edit
// that adds, removes, splits, or joins paragraphs changes the indexes. Each local transaction
// that changes the number of paragraphs of an animated shape is checked here. A paragraph
// follows its text: the changed regions between the old and the new text of the shape
// (`diffHunks`, paragraphs joined by line feeds) tell where each unchanged character went, and
// a paragraph goes to the paragraphs that hold its unchanged characters. An empty paragraph
// follows its line feed. A range takes the new indexes of its paragraphs. A range whose
// paragraphs all went takes its effect with it (`paragraphTargetRemovals`).
//
// Text alone cannot tell identical neighbors apart: deleting the first of two equal paragraphs
// and deleting the second give the same text. The editor names where the edit started
// (`TextEditAt`), and a deletion or insertion that fits in more than one place goes to the place
// nearest it. An edit with too many changes for a character diff is compared by whole
// paragraphs first, and only the changed paragraphs by characters.
//
// The fixes are ops of the same transaction, so undo restores the ranges with the text, and
// collaborators receive them with the edit.

import { NS } from '../package/namespaces.ts';
import { attr, child, elements, textOf } from '../package/query.ts';
import { diffHunks, diffResult, type TextHunk } from '../package/text-diff.ts';
import { paragraphTargetRemovals } from '../package/timing.ts';
import type { Element } from '../package/tree.ts';
import type { PptxOp } from './ops.ts';

const MAX_DEPTH = 64;

/**
 * Where a text edit started, before the edit: the slide part, the shape node (`p:sp`), and the
 * paragraph and offset of the start of the selection.
 */
export interface TextEditAt {
  readonly part: string;
  readonly node: string;
  readonly paragraph: number;
  readonly offset: number;
}

/** One copy of a shape: its node id and its paragraph texts. */
interface Copy {
  readonly node: string;
  readonly texts: string[];
}

/** A paragraph target: `p:spTgt` with `p:txEl/p:pRg`. */
interface Target {
  readonly spTgt: Element;
  readonly pRg: Element;
  readonly spid: string;
}

const isP = (el: Element, local: string): boolean => el.ns === NS.p && el.local === local;

/** The paragraph targets of a slide's animations, in every `p:timing`. */
function targetsOf(root: Element): Target[] {
  const out: Target[] = [];
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (isP(el, 'spTgt')) {
      const pRg = child(child(el, 'p:txEl'), 'p:pRg');
      const spid = attr(el, 'spid');
      if (pRg && spid !== undefined) out.push({ spTgt: el, pRg, spid });
      return;
    }
    for (const c of elements(el)) walk(c, depth + 1);
  };
  for (const c of elements(root))
    if (isP(c, 'timing') || (c.ns === NS.mc && c.local === 'AlternateContent')) walk(c, 0);
  return out;
}

/**
 * The paragraph texts of each copy of each shape that `spids` names, by `p:cNvPr/@id`, in
 * document order. A shape in `mc:AlternateContent` has a copy in each branch.
 */
function paragraphsOf(root: Element, spids: ReadonlySet<string>): Map<string, Copy[]> {
  const out = new Map<string, Copy[]>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (isP(el, 'sp')) {
      const id = attr(child(child(el, 'p:nvSpPr'), 'p:cNvPr'), 'id');
      if (id !== undefined && spids.has(id)) {
        const copies = out.get(id) ?? [];
        copies.push({
          node: el.id,
          texts: elements(child(el, 'p:txBody'))
            .filter((p) => p.ns === NS.a && p.local === 'p')
            // A line feed separates paragraphs; one inside a paragraph's text counts as a space.
            .map((p) => textOf(p).replace(/\n/g, ' ')),
        });
        out.set(id, copies);
      }
      return;
    }
    if (isP(el, 'timing')) return;
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/** A paragraph index of `p:pRg` (`xsd:unsignedInt`), or undefined. */
function index(value: string | undefined): number | undefined {
  if (value === undefined || !/^\s*\d{1,9}\s*$/.test(value)) return undefined;
  return Number(value.trim());
}

/** The new paragraphs that an old paragraph went to: the first and the last. */
interface Moved {
  readonly low: number;
  readonly high: number;
}

/**
 * Each hunk that only deletes or only inserts, moved to the place nearest `hint` (an offset in
 * `a`) among the places that give the same text. Other hunks stay.
 */
function slid(
  a: string,
  b: string,
  hunks: readonly TextHunk[],
  hint: number | undefined
): TextHunk[] {
  if (hint === undefined) return [...hunks];
  return hunks.map((h, i) => {
    const deletes = h.bStart === h.bEnd;
    if (deletes === (h.aStart === h.aEnd)) return h;
    const lo = i > 0 ? hunks[i - 1]!.aEnd : 0;
    const hi = i < hunks.length - 1 ? hunks[i + 1]!.aStart : a.length;
    const s = deletes ? a : b;
    const start = deletes ? h.aStart : h.bStart;
    const end = deletes ? h.aEnd : h.bEnd;
    // A deletion that starts at a line feed joins the paragraph after it: aim at that paragraph,
    // whose own line feed an empty paragraph follows.
    const aim = deletes && a[hint] === '\n' ? hint + 1 : hint;
    let t = 0;
    while (h.aStart + t > aim && h.aStart + t > lo && s[start + t - 1] === s[end + t - 1]) t -= 1;
    while (h.aStart + t < aim && h.aEnd + t < hi && s[start + t] === s[end + t]) t += 1;
    if (t === 0) return h;
    return { aStart: h.aStart + t, aEnd: h.aEnd + t, bStart: h.bStart + t, bEnd: h.bEnd + t };
  });
}

/** Where an edit point is in the old texts joined by line feeds. */
function charAt(was: readonly string[], at: { paragraph: number; offset: number }): number {
  let c = 0;
  for (let i = 0; i < at.paragraph && i < was.length; i += 1) c += was[i]!.length + 1;
  return c + Math.max(0, Math.min(at.offset, was[at.paragraph]?.length ?? 0));
}

/**
 * Where each old paragraph went, or undefined for a paragraph that the edit replaced: fewer
 * than half of its characters are unchanged. A paragraph that the edit split goes to every
 * part. `was` and `is` are the paragraph texts before and after the edit, and `at` is where the
 * edit started, when the editor knows it.
 */
function paragraphMap(
  was: readonly string[],
  is: readonly string[],
  at?: { paragraph: number; offset: number }
): (Moved | undefined)[] {
  const a = was.join('\n');
  const b = is.join('\n');
  const r = diffResult(a, b);
  if (r.exact) return charMap(was, is, slid(a, b, r.hunks, at && charAt(was, at)));
  // Too many changes for a character diff: compare whole paragraphs, one character each, and
  // then the characters of each changed block of paragraphs. Each paragraph is one character.
  if (was.length + is.length > 0xffff) return charMap(was, is, r.hunks);
  const ids = new Map<string, number>();
  const token = (t: string): string => {
    let id = ids.get(t);
    if (id === undefined) ids.set(t, (id = ids.size));
    return String.fromCharCode(id);
  };
  const pa = was.map(token).join('');
  const pb = is.map(token).join('');
  // An edit that starts in the second half of a paragraph keeps that paragraph.
  const len = at === undefined ? 0 : (was[at.paragraph]?.length ?? 0);
  const first =
    at === undefined ? undefined : at.paragraph + (at.offset > 0 && at.offset * 2 >= len ? 1 : 0);
  const blocks = slid(pa, pb, diffHunks(pa, pb), first);
  const out: (Moved | undefined)[] = [];
  let i = 0;
  let delta = 0;
  const end = { aStart: was.length, aEnd: was.length, bStart: is.length, bEnd: is.length };
  for (const k of [...blocks, end]) {
    for (; i < k.aStart; i += 1) out.push({ low: i + delta, high: i + delta });
    if (k === end) break;
    if (k.bStart < k.bEnd && k.aStart < k.aEnd) {
      const olds = was.slice(k.aStart, k.aEnd);
      const news = is.slice(k.bStart, k.bEnd);
      for (const m of charMap(olds, news, diffHunks(olds.join('\n'), news.join('\n'))))
        out.push(m && { low: m.low + k.bStart, high: m.high + k.bStart });
    } else for (let j = k.aStart; j < k.aEnd; j += 1) out.push(undefined);
    i = k.aEnd;
    delta = k.bEnd - k.aEnd;
  }
  return out;
}

/** Where each old paragraph went, by the changed regions `hunks` of the joined texts. */
function charMap(
  was: readonly string[],
  is: readonly string[],
  hunks: readonly TextHunk[]
): (Moved | undefined)[] {
  // The start offset of each new paragraph in the new text.
  const starts: number[] = [];
  let at = 0;
  for (const p of is) {
    starts.push(at);
    at += p.length + 1;
  }
  const paragraphAt = (q: number): number => {
    let lo = 0;
    let hi = starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (starts[mid]! <= q) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  // One pass over the old text and the hunks: the new place of each unchanged character.
  let h = 0;
  let delta = 0;
  const mapped = (c: number): number | undefined => {
    while (h < hunks.length && c >= hunks[h]!.aEnd) {
      delta += hunks[h]!.bEnd - hunks[h]!.bStart - (hunks[h]!.aEnd - hunks[h]!.aStart);
      h += 1;
    }
    const hunk = hunks[h];
    return hunk && c >= hunk.aStart ? undefined : c + delta;
  };
  const out: (Moved | undefined)[] = [];
  let start = 0;
  was.forEach((text, i) => {
    let low = Infinity;
    let high = -Infinity;
    let kept = 0;
    for (let c = start; c < start + text.length; c += 1) {
      const q = mapped(c);
      if (q === undefined) continue;
      kept += 1;
      const to = paragraphAt(q);
      low = Math.min(low, to);
      high = Math.max(high, to);
    }
    if (text.length > 0) out.push(kept * 2 >= text.length ? { low, high } : undefined);
    else {
      // An empty paragraph follows its line feed: the one after it, or, for the last
      // paragraph, the one before it. A single empty paragraph goes to the first one.
      let to: number | undefined;
      if (i < was.length - 1) {
        const q = mapped(start);
        to = q === undefined ? undefined : paragraphAt(q);
      } else if (start > 0) {
        const q = mappedBack(start - 1, hunks);
        to = q === undefined ? undefined : paragraphAt(q + 1);
      } else if (is.length > 0) to = 0;
      out.push(to === undefined ? undefined : { low: to, high: to });
    }
    start += text.length + 1;
  });
  return out;
}

/** Where a character of the old text is in the new text, by a scan of the hunks. */
function mappedBack(p: number, hunks: readonly TextHunk[]): number | undefined {
  let delta = 0;
  for (const hunk of hunks) {
    if (p < hunk.aStart) break;
    if (p < hunk.aEnd) return undefined;
    delta += hunk.bEnd - hunk.bStart - (hunk.aEnd - hunk.aStart);
  }
  return p + delta;
}

/**
 * The ops that keep the paragraph ranges of slide `part` on their paragraphs after an edit
 * changed its root from `before` to `after`. `at` is where a text edit started. Empty when
 * nothing needs to change.
 */
export function paragraphRangeFixes(
  part: string,
  before: Element,
  after: Element,
  at?: TextEditAt
): PptxOp[] {
  if (!isP(after, 'sld')) return [];
  const targets = targetsOf(after);
  if (targets.length === 0) return [];
  const spids = new Set(targets.map((t) => t.spid));
  const old = paragraphsOf(before, spids);
  const now = paragraphsOf(after, spids);
  // The range of a target as it was before the edit: the same `p:pRg` node.
  const oldRanges = new Map(targetsOf(before).map((t) => [t.pRg.id, t.pRg]));
  const maps = new Map<string, (Moved | undefined)[]>();
  const ops: PptxOp[] = [];
  const gone = new Set<Element>();
  const set = (node: Element, local: string, value: number): void => {
    if (index(attr(node, local)) === value) return;
    ops.push({
      op: 'setAttribute',
      part,
      node: node.id,
      name: { ns: '', local },
      value: String(value),
    });
  };
  for (const t of targets) {
    // The copy whose paragraphs changed: the one that the editor draws.
    const olds = old.get(t.spid) ?? [];
    const news = now.get(t.spid) ?? [];
    // The same copy before the edit, by its node: an edit of one copy in mc:AlternateContent can
    // remove the out-of-date copies in the other branches.
    const prior = (c: Copy): Copy | undefined => olds.find((o) => o.node === c.node);
    const k = news.findIndex((c) => {
      const o = prior(c);
      return o !== undefined && o.texts.length !== c.texts.length;
    });
    // A shape that the edit added or removed follows its own rules (`timingRemovals`). An edit
    // that keeps the number of paragraphs keeps their indexes.
    if (k < 0) continue;
    const was = prior(news[k]!)!.texts;
    const is = news[k]!.texts;
    // A range that the same transaction added names the new paragraphs already.
    const range = oldRanges.get(t.pRg.id);
    if (!range) continue;
    const st = index(attr(range, 'st'));
    const end = index(attr(range, 'end'));
    // A range that the file wrote out of order or out of bounds stays as it is.
    if (st === undefined || end === undefined || st > end || end >= was.length) continue;
    const key = `${t.spid} ${k}`;
    let map = maps.get(key);
    const hint = at?.part === part && at.node === news[k]!.node ? at : undefined;
    if (!map) maps.set(key, (map = paragraphMap(was, is, hint)));
    let low = Infinity;
    let high = -Infinity;
    for (let i = st; i <= end; i += 1) {
      const to = map[i];
      if (to === undefined) continue;
      low = Math.min(low, to.low);
      high = Math.max(high, to.high);
    }
    if (high < 0) {
      gone.add(t.spTgt);
      continue;
    }
    set(t.pRg, 'st', low);
    set(t.pRg, 'end', high);
  }
  if (gone.size === 0) return ops;
  const removals = paragraphTargetRemovals(after, gone);
  // A range inside a removed effect needs no change.
  const removed = new Set<string>();
  const mark = (el: Element, depth: number): void => {
    removed.add(el.id);
    if (depth < MAX_DEPTH) for (const c of elements(el)) mark(c, depth + 1);
  };
  for (const el of removals) mark(el, 0);
  return [
    ...ops.filter((o) => o.op !== 'setAttribute' || !removed.has(o.node)),
    ...removals.map((el): PptxOp => ({ op: 'removeNode', part, node: el.id })),
  ];
}
