/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The changes of one matched slide, for the computed refresh changes (refresh-changes.ts):
// slide properties, shapes, text, and notes. DOM-free.
//
// Top-level shapes match by `p:cNvPr/@id`, then by name. A shape in the first `mc:Choice` of an
// `mc:AlternateContent` counts as a top-level shape. A shape whose only difference is its text
// gives one text change for each changed or inserted paragraph, located at the new text and
// widened to whole words, and one text change for its removed paragraphs. Other differences give
// one shape change with its aspects. A changed picture, chart, or diagram part counts as a change
// of the shape that names it, also when the shape XML is the same. A location names a shape only
// when its id is unique on the slide, as the editor finds it; otherwise it names the slide. A
// slide whose part changed without a change above gets one slide change, so no difference is
// silent.

import {
  elements,
  type Element,
  type PptxPackage,
  type SlideRef,
} from '@pptx-editor.dev/core/store';
import {
  shapeParagraphTexts,
  shapesById,
  type RefreshChangeAspect,
  type RefreshLocation,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canon,
  child,
  partContent,
  rootOf,
  sameContent,
  type Collector,
} from './refresh-common.ts';
import { changedRelIds, relIdsOf } from './refresh-parts.ts';
import {
  changedRange,
  diffParagraphs,
  sameFormat,
  textFormat,
  widenRange,
  type ParagraphDiff,
} from './refresh-text.ts';

/** The most text changes one shape reports; past it, the shape reports one shape change. */
const MAX_TEXT_CHANGES = 50;

interface Shape {
  readonly el: Element;
  readonly id: number | undefined;
  readonly name: string | undefined;
}

function shapeOf(el: Element): Shape | undefined {
  const nv = elements(el)[0];
  if (!nv || !nv.local.startsWith('nv')) return undefined;
  const cNvPr = child(nv, 'cNvPr');
  const id = Number(attr(cNvPr, 'id'));
  return { el, id: Number.isSafeInteger(id) ? id : undefined, name: attr(cNvPr, 'name') };
}

/** The top-level shapes of a slide part's shape tree, with those of `mc:Choice`. */
function shapesOf(root: Element | undefined): Shape[] {
  const tree = child(child(root, 'cSld'), 'spTree');
  const out: Shape[] = [];
  for (const el of elements(tree)) {
    const list = el.local === 'AlternateContent' ? elements(child(el, 'Choice')) : [el];
    for (const c of list) {
      const s = shapeOf(c);
      if (s) out.push(s);
    }
  }
  return out;
}

/** Pairs of shapes: by id when it is unique on both sides, then by name, then by id and order. */
function matchShapes(a: readonly Shape[], b: readonly Shape[]): Map<Shape, Shape> {
  const pairs = new Map<Shape, Shape>(); // new -> old
  const used = new Set<Shape>();
  const unique = (list: readonly Shape[], key: (s: Shape) => string | undefined) => {
    const out = new Map<string, Shape | null>();
    for (const s of list) {
      const k = key(s);
      if (k !== undefined) out.set(k, out.has(k) ? null : s);
    }
    return out;
  };
  // A repeated id pairs by its occurrence, after the unique ids and the names.
  const occurrence = new Map<Shape, number>();
  for (const list of [a, b]) {
    const seen = new Map<string, number>();
    for (const s of list) {
      if (s.id === undefined) continue;
      const k = `${s.el.local}#${s.id}`;
      const n = seen.get(k) ?? 0;
      seen.set(k, n + 1);
      occurrence.set(s, n);
    }
  }
  for (const key of [
    (s: Shape) => (s.id === undefined ? undefined : `${s.el.local}#${s.id}`),
    (s: Shape) => (s.name === undefined ? undefined : `${s.el.local}:${s.name}`),
    (s: Shape) => (s.id === undefined ? undefined : `${s.el.local}#${s.id}#${occurrence.get(s)}`),
  ]) {
    const olds = unique(
      a.filter((s) => !used.has(s)),
      key
    );
    const news = unique(
      b.filter((s) => !pairs.has(s)),
      key
    );
    for (const s of b) {
      if (pairs.has(s)) continue;
      const k = key(s);
      if (k === undefined || news.get(k) !== s) continue;
      const old = olds.get(k);
      if (!old || used.has(old)) continue;
      pairs.set(s, old);
      used.add(old);
    }
  }
  return pairs;
}

const hits = (el: Element | undefined, changed: ReadonlySet<string>): string[] =>
  el && changed.size > 0 ? [...relIdsOf(el)].filter((id) => changed.has(id)) : [];

/**
 * What differs between two versions of a shape, other than its text. `relHits` holds the
 * relationship ids of the shape whose related parts changed.
 */
function shapeAspects(a: Element, b: Element, relHits: ReadonlySet<string>): RefreshChangeAspect[] {
  const out = new Set<RefreshChangeAspect>();
  const prA = elements(a).find((c) => c.local.endsWith('spPr'));
  const prB = elements(b).find((c) => c.local.endsWith('spPr'));
  const xfA = child(prA, 'xfrm') ?? child(a, 'xfrm');
  const xfB = child(prB, 'xfrm') ?? child(b, 'xfrm');
  if (canon(child(xfA, 'off')) !== canon(child(xfB, 'off'))) out.add('position');
  if (canon(child(xfA, 'ext')) !== canon(child(xfB, 'ext'))) out.add('size');
  const turn = (x: Element | undefined) =>
    `${attr(x, 'rot')}|${attr(x, 'flipH')}|${attr(x, 'flipV')}`;
  if (turn(xfA) !== turn(xfB)) out.add('rotation');
  const pick = (pr: Element | undefined, names: readonly string[]) =>
    elements(pr)
      .filter((c) => names.includes(c.local))
      .map(canon)
      .join('');
  const groups: [RefreshChangeAspect, readonly string[]][] = [
    ['geometry', ['prstGeom', 'custGeom']],
    ['fill', ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']],
    ['line', ['ln']],
    ['effects', ['effectLst', 'effectDag', 'scene3d', 'sp3d']],
  ];
  for (const [aspect, names] of groups) if (pick(prA, names) !== pick(prB, names)) out.add(aspect);
  const explained = new Set<string>();
  const related = (el: Element | undefined) => {
    const found = el ? [...relIdsOf(el)].filter((id) => relHits.has(id)) : [];
    for (const id of found) explained.add(id);
    return found.length > 0;
  };
  // A picture: its fill element, or the part its relationship names.
  if (b.local === 'pic') {
    const fillB = child(b, 'blipFill');
    if (canon(child(a, 'blipFill')) !== canon(fillB) || related(fillB)) out.add('picture');
  }
  if (b.local === 'graphicFrame') {
    const data = (el: Element) => child(child(el, 'graphic'), 'graphicData');
    const uri = attr(data(b), 'uri') ?? '';
    const kind: RefreshChangeAspect = uri.endsWith('/table')
      ? 'table'
      : uri.endsWith('/chart')
        ? 'chart'
        : uri.endsWith('/diagram')
          ? 'diagram'
          : 'other';
    // A chart or a diagram keeps its data in other parts, which `relHits` compares.
    if (canon(data(a)) !== canon(data(b)) || related(data(b))) out.add(kind);
  }
  // Another relationship of the shape, such as a hyperlink target.
  if ([...relHits].some((id) => !explained.has(id))) out.add('other');
  if (out.size === 0) out.add('other');
  return [...out];
}

const sameTexts = (a: readonly string[], b: readonly string[]) =>
  a.length === b.length && a.every((t, i) => t === b[i]);

const NO_DIFF: ParagraphDiff = { changed: [], removed: [], kept: [] };

/** One matched slide of a comparison. */
export interface SlidePair {
  readonly a: PptxPackage;
  readonly b: PptxPackage;
  readonly sa: SlideRef;
  readonly sb: SlideRef;
  /** The slide index in the returned presentation. */
  readonly index: number;
  /** The slide key for the change ids. */
  readonly key: string;
  /** The parts reached through the slide's relationships go here. */
  readonly reached: Set<string>;
}

/**
 * The changes of one matched slide. Returns true when a relationship of the slide changed that
 * no part of the slide names: the caller reports it as a package change.
 */
export function slideChanges(out: Collector, pair: SlidePair): boolean {
  const { a, b, sa, sb, index, key, reached } = pair;
  let emitted = 0;
  const emit: Collector['add'] = (change) => {
    emitted += 1;
    out.add(change);
  };
  const rootA = rootOf(a, sa.partName);
  const rootB = rootOf(b, sb.partName);
  const changed = changedRelIds(a, b, sa.partName, sb.partName, reached);
  const sameXml = canon(rootA) === canon(rootB);
  let unexplained = false;
  if (!sameXml || changed.size > 0 || sa.layoutPart !== sb.layoutPart) {
    const slideAspects = new Set<RefreshChangeAspect>();
    if (sa.layoutPart !== sb.layoutPart) slideAspects.add('layout');
    if (sa.hidden !== sb.hidden) slideAspects.add('visibility');
    const bgA = child(child(rootA, 'cSld'), 'bg');
    const bgB = child(child(rootB, 'cSld'), 'bg');
    if (canon(bgA) !== canon(bgB) || hits(bgA, changed).length + hits(bgB, changed).length > 0)
      slideAspects.add('background');
    const around = (root: Element | undefined, names: readonly string[]) =>
      elements(root)
        .filter(
          (c) =>
            names.includes(c.local) ||
            (c.local === 'AlternateContent' && canon(c).includes('transition'))
        )
        .map(canon)
        .join('');
    if (around(rootA, ['transition']) !== around(rootB, ['transition']))
      slideAspects.add('transition');
    if (around(rootA, ['timing']) !== around(rootB, ['timing'])) slideAspects.add('animation');
    const shapesA = shapesOf(rootA);
    const shapesB = shapesOf(rootB);
    if (changed.size > 0) {
      // A changed relationship that no shape and no background names.
      const named = new Set<string>();
      for (const s of [...shapesA, ...shapesB]) relIdsOf(s.el, named);
      for (const bg of [bgA, bgB]) if (bg) relIdsOf(bg, named);
      const elsewhere = new Set<string>();
      for (const root of [rootA, rootB]) if (root) relIdsOf(root, elsewhere);
      for (const id of changed)
        if (!named.has(id)) {
          if (elsewhere.has(id)) slideAspects.add('other');
          else unexplained = true;
        }
    }
    if (slideAspects.size > 0)
      emit({
        id: `slide:${key}`,
        location: { slideIndex: index },
        computed: { scope: 'slide', action: 'changed', aspects: [...slideAspects] },
      });
    shapeChanges(emit, pair, rootB, shapesA, shapesB, changed);
    if (emitted === 0 && !sameXml)
      emit({
        id: `slide:${key}`,
        location: { slideIndex: index },
        computed: { scope: 'slide', action: 'changed', aspects: ['other'] },
      });
  }
  if (sb.notesPart) {
    if (!sameContent(partContent(a, sa.notesPart), partContent(b, sb.notesPart)))
      emit({
        id: `notes:${key}`,
        location: { slideIndex: index, notes: true },
        computed: { scope: 'notes', action: 'changed', aspects: ['text'] },
      });
  } else if (sa.notesPart)
    // The notes part is gone: the location names the slide.
    emit({
      id: `notes:${key}`,
      location: { slideIndex: index },
      computed: { scope: 'notes', action: 'removed' },
    });
  return unexplained;
}

/** The shape and text changes of one matched slide. */
function shapeChanges(
  emit: Collector['add'],
  pair: SlidePair,
  rootB: Element | undefined,
  shapesA: readonly Shape[],
  shapesB: readonly Shape[],
  changed: ReadonlySet<string>
): void {
  const { index, key } = pair;
  const pairs = matchShapes(shapesA, shapesB);
  const matchedOld = new Set(pairs.values());
  for (const s of shapesA) {
    if (matchedOld.has(s)) continue;
    emit({
      id: `shape:${key}:${s.id ?? s.name ?? '?'}:removed`,
      location: { slideIndex: index },
      computed: { scope: 'shape', action: 'removed' },
    });
  }
  for (const s of shapesB) {
    const old = pairs.get(s);
    const shapeKey = `${key}:${s.id ?? s.name ?? '?'}`;
    // The editor finds a shape by an id that is unique on the slide, in groups too.
    const shapeId = s.id !== undefined && shapesById(rootB, s.id).length === 1 ? s.id : undefined;
    const at: RefreshLocation =
      shapeId === undefined ? { slideIndex: index } : { slideIndex: index, shapeId };
    if (!old) {
      emit({
        id: `shape:${shapeKey}`,
        location: at,
        computed: { scope: 'shape', action: 'added' },
      });
      continue;
    }
    const relHits = new Set([...hits(old.el, changed), ...hits(s.el, changed)]);
    if (canon(old.el) === canon(s.el) && relHits.size === 0) continue;
    const textsA = shapeParagraphTexts(old.el);
    const textsB = shapeParagraphTexts(s.el);
    const textChanged = !sameTexts(textsA, textsB);
    const diff = textChanged ? diffParagraphs(textsA, textsB) : NO_DIFF;
    const withoutText = (el: Element) =>
      canon({
        ...el,
        children: el.children.filter((c) => c.type !== 'element' || c.local !== 'txBody'),
      });
    const bodyOnly = withoutText(old.el) === withoutText(s.el);
    const formatSame = sameFormat(textFormat(old.el), textFormat(s.el), diff);
    const count = diff ? diff.changed.length + (diff.removed.length > 0 ? 1 : 0) : Infinity;
    if (
      bodyOnly &&
      formatSame &&
      relHits.size === 0 &&
      textChanged &&
      diff &&
      count <= MAX_TEXT_CHANGES &&
      shapeId !== undefined
    ) {
      for (const [p, oldIndex] of diff.changed) {
        const text = textsB[p]!;
        const range =
          oldIndex === null
            ? { start: 0, end: text.length }
            : (() => {
                const r = changedRange(textsA[oldIndex]!, text);
                return widenRange(text, r.start, r.end);
              })();
        emit({
          id: `text:${shapeKey}:${p}`,
          location:
            range.end > range.start
              ? {
                  slideIndex: index,
                  shapeId,
                  paragraphIndex: p,
                  start: range.start,
                  end: range.end,
                  text: text.slice(range.start, range.end),
                }
              : at,
          computed: { scope: 'text', action: 'changed', aspects: ['text'] },
        });
      }
      // Removed paragraphs have no text left: the change locates the shape.
      if (diff.removed.length > 0)
        emit({
          id: `text:${shapeKey}:removed`,
          location: at,
          computed: { scope: 'text', action: 'changed', aspects: ['text'] },
        });
      continue;
    }
    const aspects = bodyOnly && relHits.size === 0 ? [] : shapeAspects(old.el, s.el, relHits);
    if (textChanged) aspects.push('text');
    if (!formatSame) aspects.push('textFormat');
    const unique = [...new Set(aspects.length > 0 ? aspects : ['other' as const])];
    emit({
      id: `shape:${shapeKey}`,
      location: at,
      computed: { scope: 'shape', action: 'changed', aspects: unique },
    });
  }
}
