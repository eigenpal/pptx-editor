// Animation targets that name paragraphs follow their paragraphs through text edits.
//
// An effect can animate some paragraphs of a shape: `p:spTgt` holds a `p:txEl` with a `p:pRg`
// (ECMA-376 Part 1, §19.5.72 and §19.5.63), whose `st` and `end` are paragraph indexes. An edit
// that adds, removes, splits, or joins paragraphs changes the indexes. Each local transaction
// that changes the number of paragraphs of an animated shape is checked here. A paragraph
// follows its text: the changed regions between the old and the new text of the shape
// (`diffHunks`, paragraphs joined by line feeds) tell where each unchanged character went, and
// a paragraph goes to the paragraph that holds its first unchanged character. An empty
// paragraph follows its line feed. A range takes the new indexes of its paragraphs. A range
// whose paragraphs all went takes its effect with it (`paragraphTargetRemovals`).
//
// The fixes are ops of the same transaction, so undo restores the ranges with the text, and
// collaborators receive them with the edit.

import { NS } from '../package/namespaces.ts';
import { attr, child, elements, textOf } from '../package/query.ts';
import { diffHunks, type TextHunk } from '../package/text-diff.ts';
import { paragraphTargetRemovals } from '../package/timing.ts';
import type { Element } from '../package/tree.ts';
import type { PptxOp } from './ops.ts';

const MAX_DEPTH = 64;

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
function paragraphsOf(root: Element, spids: ReadonlySet<string>): Map<string, string[][]> {
  const out = new Map<string, string[][]>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (isP(el, 'sp')) {
      const id = attr(child(child(el, 'p:nvSpPr'), 'p:cNvPr'), 'id');
      if (id !== undefined && spids.has(id)) {
        const copies = out.get(id) ?? [];
        copies.push(
          elements(child(el, 'p:txBody'))
            .filter((p) => p.ns === NS.a && p.local === 'p')
            // A line feed separates paragraphs; one inside a paragraph's text counts as a space.
            .map((p) => textOf(p).replace(/\n/g, ' '))
        );
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
 * Where each old paragraph went, or undefined for a paragraph that the edit replaced: fewer
 * than half of its characters are unchanged. A paragraph that the edit split goes to every
 * part. `was` and `is` are the paragraph texts before and after the edit.
 */
function paragraphMap(was: readonly string[], is: readonly string[]): (Moved | undefined)[] {
  const hunks = diffHunks(was.join('\n'), is.join('\n'));
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
 * changed its root from `before` to `after`. Empty when nothing needs to change.
 */
export function paragraphRangeFixes(part: string, before: Element, after: Element): PptxOp[] {
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
    const k = news.findIndex((is, i) => olds[i] !== undefined && olds[i]!.length !== is.length);
    // A shape that the edit added or removed follows its own rules (`timingRemovals`). An edit
    // that keeps the number of paragraphs keeps their indexes.
    if (k < 0 || olds.length !== news.length) continue;
    const was = olds[k]!;
    const is = news[k]!;
    // A range that the same transaction added names the new paragraphs already.
    const range = oldRanges.get(t.pRg.id);
    if (!range) continue;
    const st = index(attr(range, 'st'));
    const end = index(attr(range, 'end'));
    // A range that the file wrote out of order or out of bounds stays as it is.
    if (st === undefined || end === undefined || st > end || end >= was.length) continue;
    const key = `${t.spid} ${k}`;
    let map = maps.get(key);
    if (!map) maps.set(key, (map = paragraphMap(was, is)));
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
