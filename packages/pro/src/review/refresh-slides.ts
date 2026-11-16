/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Slide matching of the computed refresh changes (refresh-changes.ts): which returned slide is
// which captured slide, and the key of each slide for the change ids. DOM-free.
//
// Slides match in passes. Each pass pairs only the slides that the passes before it left:
//
// 1. `p14:creationId`.
// 2. Equal content and layout: a writer that numbers its slides again after an insertion keeps
//    the content of the slides after it.
// 3. The same `p:sldId/@id` when the two slides are similar (`SIMILAR`).
// 4. Similar slides, when each one is the other's clear best match. A writer that numbers its
//    slides again after a deletion keeps an edited slide here.
// 5. The same `p:sldId/@id`.
// 6. The same part name, only when the two presentations share no slide id: when they share
//    one, ids are stable, and a different id is a different slide.

import {
  elements,
  type Element,
  type PptxPackage,
  type SlideRef,
} from '@pptx-editor.dev/core/store';
import { attr, child, NS_P14, partContent, rootOf } from './refresh-common.ts';

/** The least similarity (0 through 1) of two slides that a similarity pass pairs. */
const SIMILAR = 0.5;
/** The most slide pairs one similarity pass compares. */
const MAX_SIMILARITY_PAIRS = 40_000;

/** `p14:creationId/@val` of a slide part, or undefined. */
export function creationId(pkg: PptxPackage, part: string): string | undefined {
  const root = rootOf(pkg, part);
  for (const ext of elements(child(root, 'extLst')))
    for (const c of elements(ext))
      if (c.ns === NS_P14 && c.local === 'creationId') return attr(c, 'val');
  return undefined;
}

/** The features of a slide for similarity: its top-level shapes and the words of its text. */
interface Features {
  readonly shapes: ReadonlySet<string>;
  readonly words: ReadonlySet<string>;
}

function featuresOf(root: Element | undefined): Features {
  const shapes = new Set<string>();
  const words = new Set<string>();
  const tree = child(child(root, 'cSld'), 'spTree');
  for (const el of elements(tree)) {
    const nv = elements(el)[0];
    const cNvPr = nv && nv.local.startsWith('nv') ? child(nv, 'cNvPr') : undefined;
    if (cNvPr) shapes.add(`${el.local}#${attr(cNvPr, 'id')}:${attr(cNvPr, 'name')}`);
  }
  const visit = (el: Element, depth: number) => {
    if (depth > 64) return;
    for (const c of elements(el)) {
      if (c.local === 't')
        for (const node of c.children)
          if (node.type === 'text')
            for (const m of node.value.toLowerCase().matchAll(/[\p{L}\p{N}]+/gu)) words.add(m[0]);
      visit(c, depth + 1);
    }
  };
  if (tree) visit(tree, 0);
  return { shapes, words };
}

function jaccard(a: ReadonlySet<string>, b: ReadonlySet<string>): number {
  if (a.size === 0 && b.size === 0) return 0;
  let both = 0;
  for (const x of a) if (b.has(x)) both += 1;
  return both / (a.size + b.size - both);
}

/** The similarity of two slides, 0 through 1. Text counts three times as much as shapes. */
function similarity(a: Features, b: Features): number {
  const shapes = jaccard(a.shapes, b.shapes);
  if (a.words.size === 0 && b.words.size === 0) return shapes;
  return (shapes + 3 * jaccard(a.words, b.words)) / 4;
}

/** Pairs of slides: new index to old index. */
export function matchSlides(
  captured: PptxPackage,
  returned: PptxPackage,
  slidesA: readonly SlideRef[],
  slidesB: readonly SlideRef[]
): Map<number, number> {
  const pairs = new Map<number, number>();
  const usedOld = new Set<number>();
  const pair = (j: number, i: number) => {
    pairs.set(j, i);
    usedOld.add(i);
  };
  const byKey = (key: (pkg: PptxPackage, s: SlideRef) => string | undefined) => {
    const olds = new Map<string, number | null>();
    slidesA.forEach((s, i) => {
      if (usedOld.has(i)) return;
      const k = key(captured, s);
      if (k !== undefined) olds.set(k, olds.has(k) ? null : i);
    });
    slidesB.forEach((s, j) => {
      if (pairs.has(j)) return;
      const k = key(returned, s);
      const i = k === undefined ? undefined : olds.get(k);
      if (i === undefined || i === null || usedOld.has(i)) return;
      pair(j, i);
    });
  };
  byKey((pkg, s) => creationId(pkg, s.partName));
  byKey((pkg, s) => {
    const c = partContent(pkg, s.partName);
    return typeof c === 'string' ? `${s.layoutPart}|${c}` : undefined;
  });

  const featureCache = new Map<string, Features>();
  const features = (side: 'a' | 'b', s: SlideRef) => {
    const k = `${side}${s.partName}`;
    let f = featureCache.get(k);
    if (!f)
      featureCache.set(k, (f = featuresOf(rootOf(side === 'a' ? captured : returned, s.partName))));
    return f;
  };
  const sim = (i: number, j: number) =>
    similarity(features('a', slidesA[i]!), features('b', slidesB[j]!));

  // The same slide id and similar content.
  const oldById = new Map<string, number>();
  slidesA.forEach((s, i) => usedOld.has(i) || oldById.set(s.sldId, i));
  slidesB.forEach((s, j) => {
    if (pairs.has(j)) return;
    const i = oldById.get(s.sldId);
    if (i !== undefined && !usedOld.has(i) && sim(i, j) >= SIMILAR) pair(j, i);
  });

  // Similar slides that are each other's clear best match.
  const restA = slidesA.flatMap((_, i) => (usedOld.has(i) ? [] : [i]));
  const restB = slidesB.flatMap((_, j) => (pairs.has(j) ? [] : [j]));
  if (restA.length > 0 && restB.length > 0 && restA.length * restB.length <= MAX_SIMILARITY_PAIRS) {
    const bestOld = new Map<number, { j: number; score: number; tie: boolean }>();
    const bestNew = new Map<number, { i: number; score: number; tie: boolean }>();
    for (const j of restB)
      for (const i of restA) {
        const score = sim(i, j);
        if (score < SIMILAR) continue;
        const bo = bestOld.get(i);
        if (!bo || score > bo.score) bestOld.set(i, { j, score, tie: false });
        else if (score === bo.score) bo.tie = true;
        const bn = bestNew.get(j);
        if (!bn || score > bn.score) bestNew.set(j, { i, score, tie: false });
        else if (score === bn.score) bn.tie = true;
      }
    for (const [j, bn] of bestNew) {
      const bo = bestOld.get(bn.i);
      if (!bn.tie && bo && !bo.tie && bo.j === j) pair(j, bn.i);
    }
  }

  byKey((_pkg, s) => s.sldId);
  const oldIds = new Set(slidesA.map((s) => s.sldId));
  if (!slidesB.some((s) => oldIds.has(s.sldId))) byKey((_pkg, s) => s.partName);
  return pairs;
}

/** The key of each slide for the change ids: `p14:creationId`, or `id` and the slide id. */
export function slideKeys(pkg: PptxPackage, slides: readonly SlideRef[]): string[] {
  return slides.map((s) => creationId(pkg, s.partName) ?? `id${s.sldId}`);
}

/**
 * Keys that are unique in `taken`: a repeated key gets a `#n` suffix in slide order. Each key
 * goes into `taken`.
 */
export function uniqueKeys(keys: readonly string[], taken: Set<string>): string[] {
  return keys.map((key) => {
    let k = key;
    for (let n = 2; taken.has(k); n += 1) k = `${key}#${n}`;
    taken.add(k);
    return k;
  });
}
