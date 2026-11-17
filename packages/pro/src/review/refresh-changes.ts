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
// Slides match as `refresh-slides.ts` describes. Unmatched slides are added or removed; a
// matched slide out of the longest run in the old order moved. A matched slide gives the
// changes that `refresh-shapes.ts` describes. A changed layout, slide master, or theme gives one
// change with the slides that show it, not one change for each slide. A changed part that no
// slide change explains gives one package change without a location. Ids come from the keys, so
// the same change keeps its id across cumulative results, and each id is unique. DOM-free.

import {
  canonicalForm,
  presentationInfo,
  readPptx,
  type Element,
  type Node,
  type PptxPackage,
  type SlideRef,
} from '@pptx-editor.dev/core/store';
import type { ComputedRefreshChange } from '@pptx-editor.dev/core/editor';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';
import {
  attr,
  Collector,
  NS_P14,
  partContent,
  rootOf,
  sameContent,
  sameList,
} from './refresh-common.ts';
import { markReached, templateContent } from './refresh-parts.ts';
import { slideChanges } from './refresh-shapes.ts';
import { matchSlides, slideKeys, uniqueKeys } from './refresh-slides.ts';

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

/**
 * The presentation part without its slide list, for equality. Each slide id of a section list
 * (`p14:sectionLst`) becomes the captured index of its slide, and an added or removed slide
 * leaves its section: a section changes only when its name or its kept slides change.
 */
function presentationForm(
  pkg: PptxPackage,
  slideOf: (sldId: string) => string | undefined
): string {
  const root = rootOf(pkg, presentationInfo(pkg).presentationPart);
  if (!root) return '';
  const strip = (el: Element, depth: number): Element => ({
    ...el,
    children: el.children.flatMap((c): Node[] => {
      if (c.type !== 'element') return [c];
      if (depth === 0 && c.local === 'sldIdLst') return [];
      if (c.ns === NS_P14 && c.local === 'sldId') {
        const slide = slideOf(attr(c, 'id') ?? '');
        return slide === undefined
          ? []
          : [{ ...c, attrs: [{ ns: '', local: 'id', value: slide }] }];
      }
      return [depth < 64 ? strip(c, depth + 1) : c];
    }),
  });
  return canonicalForm(strip(root, 0));
}

/** Parts that never give a change of their own. */
const IGNORED_PART = /\[Content_Types\]\.xml$|^\/docProps\/|^\/_rels\/\.rels$/i;

/** The source part of a relationships part, or undefined for another part. */
function relsSource(name: string): string | undefined {
  const m = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(name);
  return m ? `${m[1]}/${m[2]}` : undefined;
}

/**
 * The changes from `captured` to `returned`, with stable ids and their `computed` fields, as
 * `RefreshUpdate.changes` takes them. `truncated` is true when the change limit stopped the
 * comparison. @internal
 */
export function computeRefreshChanges(
  captured: PptxPackage,
  returned: PptxPackage
): { readonly changes: ComputedRefreshChange[]; readonly truncated: boolean } {
  const out = new Collector();
  const slidesA = presentationInfo(captured).slides;
  const slidesB = presentationInfo(returned).slides;
  const pairs = matchSlides(captured, returned, slidesA, slidesB); // new index -> old index
  const usedOld = new Set(pairs.values());
  // A matched slide keeps the key of its captured slide, so its ids stay the same across
  // cumulative results. An added slide takes a key that no captured slide has.
  const taken = new Set<string>();
  const keysA = uniqueKeys(slideKeys(captured, slidesA), taken);
  const rawB = slideKeys(returned, slidesB);
  const keysB = slidesB.map((_, j) => {
    const i = pairs.get(j);
    return i !== undefined ? keysA[i]! : uniqueKeys([rawB[j]!], taken)[0]!;
  });
  // The parts that a slide change explains.
  const reached = new Set<string>();
  let other = false;
  slidesA.forEach((s, i) => {
    if (usedOld.has(i)) return;
    markReached(captured, s.partName, reached);
    out.add({
      id: `slide:${keysA[i]}:removed`,
      unavailableReason: 'deleted',
      computed: { scope: 'slide', action: 'removed', previousSlideIndex: i },
    });
  });
  const order = [...pairs.entries()].sort((x, y) => x[0] - y[0]);
  const position = new Map(order.map(([nj], k) => [nj, k]));
  const kept = longestRun(order.map(([, i]) => i));
  slidesB.forEach((s, j) => {
    const key = keysB[j]!;
    const i = pairs.get(j);
    if (i === undefined) {
      markReached(returned, s.partName, reached);
      out.add({
        id: `slide:${key}`,
        location: { slideIndex: j },
        computed: { scope: 'slide', action: 'added' },
      });
      return;
    }
    if (!kept.has(position.get(j)!))
      out.add({
        id: `slide:${key}:moved`,
        location: { slideIndex: j },
        computed: { scope: 'slide', action: 'moved', previousSlideIndex: i },
      });
    const pair = { a: captured, b: returned, sa: slidesA[i]!, sb: s, index: j, key, reached };
    if (slideChanges(out, pair)) other = true;
  });
  // Layouts, masters, and themes: one change each, with the slides that show it.
  const templates: ['layout' | 'master' | 'theme', (s: SlideRef) => string | undefined][] = [
    ['layout', (s) => s.layoutPart],
    ['master', (s) => s.masterPart],
    ['theme', (s) => s.themePart],
  ];
  const reported = new Set<string>(reached);
  for (const [scope, partOf] of templates) {
    const parts = new Set(slidesB.map(partOf).filter((p): p is string => !!p));
    for (const part of parts) {
      reported.add(part);
      if (sameList(templateContent(captured, part), templateContent(returned, part))) continue;
      const affectedSlides = slidesB.flatMap((s, j) => (partOf(s) === part ? [j] : []));
      out.add({ id: `${scope}:${part}`, computed: { scope, action: 'changed', affectedSlides } });
    }
  }
  // Any other changed, added, or removed part: one change without a location. The slide list
  // of the presentation part and the document properties change with every added or removed
  // slide and every save.
  for (const s of [...slidesA, ...slidesB])
    for (const p of [s.partName, s.notesPart]) if (p) reported.add(p);
  reported.add(presentationInfo(captured).presentationPart);
  reported.add(presentationInfo(returned).presentationPart);
  const oldIndex = new Map(slidesA.map((s, i) => [s.sldId, i]));
  const newIndex = new Map(slidesB.map((s, j) => [s.sldId, j]));
  const keptA = (id: string) => {
    const i = oldIndex.get(id);
    return i !== undefined && usedOld.has(i) ? String(i) : undefined;
  };
  const keptB = (id: string) => {
    const j = newIndex.get(id);
    const i = j === undefined ? undefined : pairs.get(j);
    return i === undefined ? undefined : String(i);
  };
  if (!other) other = presentationForm(captured, keptA) !== presentationForm(returned, keptB);
  const skip = (name: string) => {
    if (IGNORED_PART.test(name)) return true;
    const source = relsSource(name);
    // The relationships of a reported part belong to its own comparison.
    return reported.has(name) || (source !== undefined && reported.has(source));
  };
  for (const [name, part] of returned.parts) {
    if (other) break;
    if (skip(name)) continue;
    const before = captured.parts.get(name);
    if (
      !before ||
      before.kind !== part.kind ||
      !sameContent(partContent(captured, name), partContent(returned, name))
    )
      other = true;
  }
  for (const name of captured.parts.keys()) {
    if (other) break;
    if (!skip(name) && !returned.parts.has(name)) other = true;
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
