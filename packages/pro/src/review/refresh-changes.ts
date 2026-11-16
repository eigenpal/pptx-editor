/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The changes from one presentation to another, for a document refresh whose update has no
// `changes` (`reviewModule().review.refreshChanges`) and for `comparePresentations`. DOCX
// Editor's review module supplies the tracked revisions of a returned file; PresentationML has
// no revision markup, so this compares the two packages.
//
// Slides match by `p14:creationId`, then by `p:sldId/@id`, then by part name. Unmatched slides
// are added or removed; a matched slide out of the longest run in the old order moved. On a
// matched slide, top-level shapes match by `p:cNvPr/@id`, then by name. A shape whose only
// difference is its text gives one text change for each changed paragraph, located at the new
// text; other differences give one shape change with its aspects. Notes give one notes change.
// A changed layout, slide master, or theme gives one change with the slides that show it, not
// one change for each slide. Other changed parts give one change without a location. Ids come
// from the keys, so the same change keeps its id across cumulative results. DOM-free.

import {
  canonicalForm,
  elements,
  presentationInfo,
  readPptx,
  relTarget,
  type Element,
  type PptxPackage,
  type SlideRef,
} from '@pptx-editor.dev/core/store';
import {
  shapeParagraphTexts,
  type ComputedRefreshChange,
  type RefreshChangeAspect,
} from '@pptx-editor.dev/core/editor';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';

/** The most changes one comparison reports. */
const MAX_CHANGES = 10000;
/** The most text changes one shape reports; past it, the shape reports one shape change. */
const MAX_TEXT_CHANGES = 50;

const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

const canon = (el: Element | undefined): string => (el ? canonicalForm(el) : '');
const child = (el: Element | undefined, local: string): Element | undefined =>
  el ? elements(el).find((c) => c.local === local) : undefined;
const attr = (el: Element | undefined, local: string): string | undefined =>
  el?.attrs.find((a) => a.ns === '' && a.local === local)?.value;

/** `p14:creationId/@val` of a slide part, or undefined. */
function creationId(pkg: PptxPackage, part: string): string | undefined {
  const root = rootOf(pkg, part);
  for (const ext of elements(child(root, 'extLst')))
    for (const c of elements(ext))
      if (c.ns === NS_P14 && c.local === 'creationId') return attr(c, 'val');
  return undefined;
}

function rootOf(pkg: PptxPackage, part: string | undefined): Element | undefined {
  const p = part ? pkg.parts.get(part) : undefined;
  return p?.kind === 'xml' ? p.xml.root : undefined;
}

/** The bytes or canonical XML of a part, for equality. */
function partContent(pkg: PptxPackage, part: string | undefined): string | Uint8Array | undefined {
  const p = part ? pkg.parts.get(part) : undefined;
  if (!p) return undefined;
  return p.kind === 'xml' ? canonicalForm(p.xml.root) : p.bytes;
}

function sameContent(a: string | Uint8Array | undefined, b: string | Uint8Array | undefined) {
  if (typeof a === 'string' || typeof b === 'string') return a === b;
  if (!a || !b) return a === b;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

/** The longest increasing run of `values`, as a set of positions. */
function longestRun(values: readonly number[]): Set<number> {
  const tails: number[] = [];
  const prev = new Array<number>(values.length).fill(-1);
  for (let i = 0; i < values.length; i += 1) {
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (values[tails[mid]!]! < values[i]!) lo = mid + 1;
      else hi = mid;
    }
    if (lo > 0) prev[i] = tails[lo - 1]!;
    tails[lo] = i;
  }
  const out = new Set<number>();
  for (let i = tails[tails.length - 1] ?? -1; i >= 0; i = prev[i]!) out.add(i);
  return out;
}

interface Shape {
  readonly el: Element;
  readonly id: number | undefined;
  readonly name: string | undefined;
}

/** The top-level shapes of a slide part's shape tree. */
function shapesOf(root: Element | undefined): Shape[] {
  const tree = child(child(root, 'cSld'), 'spTree');
  const out: Shape[] = [];
  for (const el of elements(tree)) {
    const nv = elements(el)[0];
    if (!nv || !nv.local.startsWith('nv')) continue;
    const cNvPr = child(nv, 'cNvPr');
    const id = Number(attr(cNvPr, 'id'));
    out.push({ el, id: Number.isSafeInteger(id) ? id : undefined, name: attr(cNvPr, 'name') });
  }
  return out;
}

/** Pairs of shapes, by id when it is unique on both sides, then by name. */
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
  for (const key of [
    (s: Shape) => (s.id === undefined ? undefined : `${s.el.local}#${s.id}`),
    (s: Shape) => (s.name === undefined ? undefined : `${s.el.local}:${s.name}`),
  ]) {
    const olds = unique(
      a.filter((s) => !used.has(s)),
      key
    );
    for (const s of b) {
      if (pairs.has(s)) continue;
      const k = key(s);
      const old = k === undefined ? undefined : olds.get(k);
      if (!old || used.has(old)) continue;
      pairs.set(s, old);
      used.add(old);
    }
  }
  return pairs;
}

/** What differs between two versions of a shape, other than its text. */
function shapeAspects(
  a: Element,
  b: Element,
  pkgA: PptxPackage,
  pkgB: PptxPackage,
  partA: string,
  partB: string
): RefreshChangeAspect[] {
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
      .map(canonicalForm)
      .join('');
  const groups: [RefreshChangeAspect, readonly string[]][] = [
    ['geometry', ['prstGeom', 'custGeom']],
    ['fill', ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']],
    ['line', ['ln']],
    ['effects', ['effectLst', 'effectDag', 'scene3d', 'sp3d']],
  ];
  for (const [aspect, names] of groups) if (pick(prA, names) !== pick(prB, names)) out.add(aspect);
  // A picture: its fill element, or the bytes its relationship names.
  if (b.local === 'pic') {
    const blip = (el: Element) => child(child(el, 'blipFill'), 'blip');
    const embed = (el: Element) =>
      blip(el)?.attrs.find((x) => x.local === 'embed' || x.local === 'link')?.value;
    const targetA = embed(a) === undefined ? undefined : relTarget(pkgA, partA, embed(a)!);
    const targetB = embed(b) === undefined ? undefined : relTarget(pkgB, partB, embed(b)!);
    if (
      canon(child(a, 'blipFill')) !== canon(child(b, 'blipFill')) ||
      !sameContent(partContent(pkgA, targetA), partContent(pkgB, targetB))
    )
      out.add('picture');
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
    // A chart or a diagram keeps its data in other parts: compare those through the relationships.
    const related = (el: Element, pkg: PptxPackage, part: string) =>
      elements(data(el))
        .flatMap((c) => c.attrs.filter((x) => x.ns.endsWith('/relationships')))
        .map((x) => partContent(pkg, relTarget(pkg, part, x.value)));
    const ra = related(a, pkgA, partA);
    const rb = related(b, pkgB, partB);
    if (
      canon(data(a)) !== canon(data(b)) ||
      ra.length !== rb.length ||
      ra.some((x, i) => !sameContent(x, rb[i]))
    )
      out.add(kind);
  }
  if (out.size === 0) out.add('other');
  return [...out];
}

/** The changed range of `after` against `before`: the part between their common ends. */
function changedRange(before: string, after: string): { start: number; end: number } {
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

/** The changed paragraphs of a shape's text: `[new index, old text or null]`. */
function changedParagraphs(
  before: readonly string[],
  after: readonly string[]
): [number, string | null][] {
  // Equal paragraphs in order; the rest pair by position between them.
  const n = before.length;
  const m = after.length;
  const table: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i -= 1)
    for (let j = m - 1; j >= 0; j -= 1)
      table[i]![j] =
        before[i] === after[j]
          ? table[i + 1]![j + 1]! + 1
          : Math.max(table[i + 1]![j]!, table[i]![j + 1]!);
  const out: [number, string | null][] = [];
  let i = 0;
  let j = 0;
  const pending: number[] = [];
  const flush = (oldFrom: number, oldTo: number) => {
    pending.forEach((nj, k) => out.push([nj, oldFrom + k < oldTo ? before[oldFrom + k]! : null]));
    pending.length = 0;
  };
  let oldFrom = 0;
  while (i < n || j < m) {
    if (i < n && j < m && before[i] === after[j]) {
      flush(oldFrom, i);
      i += 1;
      j += 1;
      oldFrom = i;
    } else if (j < m && (i >= n || table[i]![j + 1]! >= table[i + 1]![j]!)) {
      pending.push(j);
      j += 1;
    } else i += 1;
  }
  flush(oldFrom, i);
  return out;
}

/** The text body of a shape without its paragraphs' text: formatting and body properties. */
function textFormat(el: Element): string {
  const body = child(el, 'txBody');
  if (!body) return '';
  const strip = (node: Element): string =>
    `<${node.local} ${node.attrs.map((x) => `${x.local}=${x.value}`).join(' ')}>` +
    elements(node)
      .filter((c) => c.local !== 't')
      .map(strip)
      .join('');
  return strip(body);
}

class Collector {
  readonly changes: ComputedRefreshChange[] = [];
  truncated = false;

  add(change: ComputedRefreshChange): void {
    if (this.changes.length >= MAX_CHANGES) {
      this.truncated = true;
      return;
    }
    this.changes.push(change);
  }
}

/** The changes of one matched slide. */
function slideChanges(
  out: Collector,
  a: PptxPackage,
  b: PptxPackage,
  sa: SlideRef,
  sb: SlideRef,
  index: number,
  key: string
): void {
  const rootA = rootOf(a, sa.partName);
  const rootB = rootOf(b, sb.partName);
  const slideAspects = new Set<RefreshChangeAspect>();
  if (sa.layoutPart !== sb.layoutPart) slideAspects.add('layout');
  if (canon(child(child(rootA, 'cSld'), 'bg')) !== canon(child(child(rootB, 'cSld'), 'bg')))
    slideAspects.add('background');
  const around = (root: Element | undefined, names: readonly string[]) =>
    elements(root)
      .filter(
        (c) =>
          names.includes(c.local) ||
          (c.local === 'AlternateContent' && canon(c).includes('transition'))
      )
      .map(canonicalForm)
      .join('');
  if (around(rootA, ['transition']) !== around(rootB, ['transition']))
    slideAspects.add('transition');
  if (around(rootA, ['timing']) !== around(rootB, ['timing'])) slideAspects.add('animation');
  if (slideAspects.size > 0)
    out.add({
      id: `slide:${key}`,
      location: { slideIndex: index },
      computed: { scope: 'slide', action: 'changed', aspects: [...slideAspects] },
    });
  const shapesA = shapesOf(rootA);
  const shapesB = shapesOf(rootB);
  const pairs = matchShapes(shapesA, shapesB);
  const matchedOld = new Set(pairs.values());
  for (const s of shapesA) {
    if (matchedOld.has(s)) continue;
    out.add({
      id: `shape:${key}:${s.id ?? s.name ?? '?'}:removed`,
      location: { slideIndex: index },
      computed: { scope: 'shape', action: 'removed' },
    });
  }
  for (const s of shapesB) {
    const old = pairs.get(s);
    const shapeKey = `${key}:${s.id ?? s.name ?? '?'}`;
    const at = s.id === undefined ? { slideIndex: index } : { slideIndex: index, shapeId: s.id };
    if (!old) {
      out.add({
        id: `shape:${shapeKey}`,
        location: at,
        computed: { scope: 'shape', action: 'added' },
      });
      continue;
    }
    if (canonicalForm(old.el) === canonicalForm(s.el)) continue;
    const textsA = shapeParagraphTexts(old.el);
    const textsB = shapeParagraphTexts(s.el);
    const withoutText = (el: Element) =>
      canonicalForm({
        ...el,
        children: el.children.filter((c) => c.type !== 'element' || c.local !== 'txBody'),
      });
    const bodyOnly = withoutText(old.el) === withoutText(s.el);
    const formatOnly = textFormat(old.el) === textFormat(s.el);
    const paragraphs =
      textsA.join('\n') === textsB.join('\n') ? [] : changedParagraphs(textsA, textsB);
    if (
      bodyOnly &&
      formatOnly &&
      paragraphs.length > 0 &&
      paragraphs.length <= MAX_TEXT_CHANGES &&
      s.id !== undefined
    ) {
      for (const [p, before] of paragraphs) {
        const text = textsB[p]!;
        const range = changedRange(before ?? '', text);
        const location =
          range.end > range.start
            ? {
                slideIndex: index,
                shapeId: s.id,
                paragraphIndex: p,
                start: range.start,
                end: range.end,
                text: text.slice(range.start, range.end),
              }
            : { slideIndex: index, shapeId: s.id };
        out.add({
          id: `text:${shapeKey}:${p}`,
          location,
          computed: { scope: 'text', action: 'changed', aspects: ['text'] },
        });
      }
      continue;
    }
    const aspects = bodyOnly ? [] : shapeAspects(old.el, s.el, a, b, sa.partName, sb.partName);
    if (paragraphs.length > 0) aspects.push('text');
    if (!formatOnly) aspects.push('textFormat');
    const unique = [...new Set(aspects.length > 0 ? aspects : ['other' as const])];
    out.add({
      id: `shape:${shapeKey}`,
      location: at,
      computed: { scope: 'shape', action: 'changed', aspects: unique },
    });
  }
  if (!sameContent(partContent(a, sa.notesPart), partContent(b, sb.notesPart)) && sb.notesPart)
    out.add({
      id: `notes:${key}`,
      location: { slideIndex: index, notes: true },
      computed: { scope: 'notes', action: 'changed', aspects: ['text'] },
    });
}

/**
 * The changes from `captured` to `returned`, with stable ids and their `computed` fields, as
 * `RefreshUpdate.changes` takes them. @internal
 */
export function computeRefreshChanges(
  captured: PptxPackage,
  returned: PptxPackage
): { readonly changes: ComputedRefreshChange[]; readonly truncated: boolean } {
  const out = new Collector();
  const slidesA = presentationInfo(captured).slides;
  const slidesB = presentationInfo(returned).slides;
  const keyOf = (pkg: PptxPackage, s: SlideRef) => creationId(pkg, s.partName) ?? `id${s.sldId}`;
  // Slides: creation id, then slide id, then part name.
  const pairs = new Map<number, number>(); // new index -> old index
  const usedOld = new Set<number>();
  // Equal content comes before the slide id and the part name: a writer that numbers its
  // slides again after an insertion keeps the content of the slides after it.
  const content = (pkg: PptxPackage, s: SlideRef) => {
    const c = partContent(pkg, s.partName);
    return typeof c === 'string' ? `${s.layoutPart}|${c}` : undefined;
  };
  const keys: ((pkg: PptxPackage, s: SlideRef) => string | undefined)[] = [
    (pkg, s) => creationId(pkg, s.partName),
    content,
    (_pkg, s) => s.sldId,
    (_pkg, s) => s.partName,
  ];
  for (const key of keys) {
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
      pairs.set(j, i);
      usedOld.add(i);
    });
  }
  slidesA.forEach((s, i) => {
    if (usedOld.has(i)) return;
    out.add({
      id: `slide:${keyOf(captured, s)}:removed`,
      unavailableReason: 'deleted',
      computed: { scope: 'slide', action: 'removed', previousSlideIndex: i },
    });
  });
  const order = [...pairs.entries()].sort((x, y) => x[0] - y[0]);
  const kept = longestRun(order.map(([, i]) => i));
  slidesB.forEach((s, j) => {
    const key = keyOf(returned, s);
    const i = pairs.get(j);
    if (i === undefined) {
      out.add({
        id: `slide:${key}`,
        location: { slideIndex: j },
        computed: { scope: 'slide', action: 'added' },
      });
      return;
    }
    const position = order.findIndex(([nj]) => nj === j);
    if (!kept.has(position))
      out.add({
        id: `slide:${key}:moved`,
        location: { slideIndex: j },
        computed: { scope: 'slide', action: 'moved', previousSlideIndex: i },
      });
    slideChanges(out, captured, returned, slidesA[i]!, s, j, key);
  });
  // Layouts, masters, and themes: one change each, with the slides that show it.
  const templates: ['layout' | 'master' | 'theme', (s: SlideRef) => string | undefined][] = [
    ['layout', (s) => s.layoutPart],
    ['master', (s) => s.masterPart],
    ['theme', (s) => s.themePart],
  ];
  const reported = new Set<string>();
  for (const [scope, partOf] of templates) {
    const parts = new Set(slidesB.map(partOf).filter((p): p is string => !!p));
    for (const part of parts) {
      reported.add(part);
      if (sameContent(partContent(captured, part), partContent(returned, part))) continue;
      const affectedSlides = slidesB.flatMap((s, j) => (partOf(s) === part ? [j] : []));
      out.add({ id: `${scope}:${part}`, computed: { scope, action: 'changed', affectedSlides } });
    }
  }
  // Any other changed part: one change without a location. The slide list of the presentation
  // part and the document properties change with every added or removed slide and every save.
  for (const s of [...slidesA, ...slidesB])
    for (const p of [s.partName, s.notesPart]) if (p) reported.add(p);
  const main = presentationInfo(returned).presentationPart;
  reported.add(main);
  const withoutSlides = (pkg: PptxPackage) => {
    const root = rootOf(pkg, main);
    return root
      ? canonicalForm({
          ...root,
          children: root.children.filter((c) => c.type !== 'element' || c.local !== 'sldIdLst'),
        })
      : '';
  };
  let other = withoutSlides(captured) !== withoutSlides(returned);
  for (const [name, part] of returned.parts) {
    if (other) break;
    if (reported.has(name) || /\.rels$|\[Content_Types\]\.xml$|^\/docProps\//i.test(name)) continue;
    const before = captured.parts.get(name);
    if (!before || before.kind !== part.kind) {
      other = true;
      break;
    }
    if (!sameContent(partContent(captured, name), partContent(returned, name))) {
      other = true;
      break;
    }
  }
  if (other) out.add({ id: 'package', computed: { scope: 'package', action: 'changed' } });
  return { changes: out.changes, truncated: out.truncated };
}

/** Options of `comparePresentations`. @public */
export type ComparePresentationsOptions = ProLicenseOptions;

/** The outcome of `comparePresentations`. @public */
export type ComparePresentationsResult =
  | {
      readonly ok: true;
      /** The changes, in the form `RefreshUpdate.changes` takes, with stable ids. */
      readonly changes: readonly ComputedRefreshChange[];
      /** True when the change limit stopped the comparison. */
      readonly truncated: boolean;
    }
  | {
      readonly ok: false;
      readonly code: 'invalidDocument';
      /** Diagnostic detail, in English; not for display. */
      readonly message: string;
    };

/**
 * Compare two presentations and list the changes from `original` to `revised`: slides added,
 * removed, moved, or changed; shapes added, removed, or changed, with what changed; text
 * changes, located at the revised text; notes; and layouts, slide masters, and themes with the
 * slides that show them. Runs without a DOM, on a server or in a browser. Both files go through
 * the same limits as a file the editor opens. Requires the EigenPal Pro License.
 *
 * @example
 * ```ts
 * import { comparePresentations } from '@pptx-editor.dev/pro';
 *
 * const result = comparePresentations(originalBytes, revisedBytes);
 * if (result.ok) {
 *   await refresh.applyUpdate({ submission, sequence: 1, bytes: revisedBytes, changes: result.changes });
 * }
 * ```
 *
 * @public
 */
export function comparePresentations(
  original: Uint8Array | ArrayBuffer,
  revised: Uint8Array | ArrayBuffer,
  options: ComparePresentationsOptions = {}
): ComparePresentationsResult {
  rememberLicenseKey(options.licenseKey);
  const read = (bytes: Uint8Array | ArrayBuffer) =>
    readPptx(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const a = read(original);
  if (!a.ok)
    return {
      ok: false,
      code: 'invalidDocument',
      message: 'The original is not a valid PPTX file.',
    };
  const b = read(revised);
  if (!b.ok)
    return {
      ok: false,
      code: 'invalidDocument',
      message: 'The revised file is not a valid PPTX file.',
    };
  return { ok: true, ...computeRefreshChanges(a.pkg, b.pkg) };
}
