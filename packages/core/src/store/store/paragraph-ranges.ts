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

/** The paragraph texts of each shape that `spids` names, by `p:cNvPr/@id`. */
function paragraphsOf(root: Element, spids: ReadonlySet<string>): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (isP(el, 'sp')) {
      const id = attr(child(child(el, 'p:nvSpPr'), 'p:cNvPr'), 'id');
      if (id !== undefined && spids.has(id) && !out.has(id))
        out.set(
          id,
          elements(child(el, 'p:txBody'))
            .filter((p) => p.ns === NS.a && p.local === 'p')
            // A line feed separates paragraphs; one inside a paragraph's text counts as a space.
            .map((p) => textOf(p).replace(/\n/g, ' '))
        );
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

/** Where a character of the old text is in the new text; undefined when it changed. */
function mapPosition(p: number, hunks: readonly TextHunk[]): number | undefined {
  let delta = 0;
  for (const h of hunks) {
    if (p < h.aStart) break;
    if (p < h.aEnd) return undefined;
    delta += h.bEnd - h.bStart - (h.aEnd - h.aStart);
  }
  return p + delta;
}

/**
 * The new index of each old paragraph, or undefined for a paragraph whose text all changed.
 * `was` and `is` are the paragraph texts before and after the edit.
 */
function paragraphMap(was: readonly string[], is: readonly string[]): (number | undefined)[] {
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
  const out: (number | undefined)[] = [];
  let start = 0;
  was.forEach((text, i) => {
    let found: number | undefined;
    for (let c = start; c < start + text.length && found === undefined; c += 1) {
      const q = mapPosition(c, hunks);
      if (q !== undefined) found = paragraphAt(q);
    }
    if (found === undefined && text.length === 0) {
      // An empty paragraph follows its line feed: the one after it, or before the last one.
      const feed = i < was.length - 1 ? start : start - 1;
      const q = feed >= 0 ? mapPosition(feed, hunks) : undefined;
      if (q !== undefined) found = i < was.length - 1 ? paragraphAt(q) : paragraphAt(q + 1);
      if (was.length === 1 && is.length > 0) found = 0;
    }
    out.push(found);
    start += text.length + 1;
  });
  return out;
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
  const maps = new Map<string, (number | undefined)[]>();
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
    const was = old.get(t.spid);
    const is = now.get(t.spid);
    // A shape that the edit added or removed follows its own rules (`timingRemovals`). An edit
    // that keeps the number of paragraphs keeps their indexes.
    if (!was || !is || was.length === is.length) continue;
    const range = oldRanges.get(t.pRg.id) ?? t.pRg;
    const st = index(attr(range, 'st'));
    const end = index(attr(range, 'end'));
    // A range that the file wrote out of order or out of bounds stays as it is.
    if (st === undefined || end === undefined || st > end || end >= was.length) continue;
    let map = maps.get(t.spid);
    if (!map) maps.set(t.spid, (map = paragraphMap(was, is)));
    let low = Infinity;
    let high = -Infinity;
    for (let i = st; i <= end; i += 1) {
      const to = map[i];
      if (to === undefined) continue;
      low = Math.min(low, to);
      high = Math.max(high, to);
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
