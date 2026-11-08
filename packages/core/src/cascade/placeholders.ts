// Placeholder lookup: a part's placeholders, and the layout and master placeholder that a
// slide placeholder inherits from (ECMA-376 Part 1, 19.3.1.36 `ph`). In the slide master view,
// a layout placeholder inherits from its master, and a master placeholder from none.

import {
  attr,
  elements,
  isTemplateRef,
  NS,
  path,
  presentationInfo,
  shapeTreeOf,
  slideRefAt,
  templateRefs,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { contentElements } from './alternate-content.ts';

/** Placeholder identity on a shape. */
export interface PlaceholderKey {
  readonly type: string;
  readonly idx?: string;
}

export interface PlaceholderEntry extends PlaceholderKey {
  readonly el: Element;
}

export function placeholderOf(el: Element): PlaceholderKey | undefined {
  const nv = elements(el).find((c) => c.local.startsWith('nv') && c.ns === NS.p);
  const ph = path(nv, 'p:nvPr', 'p:ph');
  if (!ph) return undefined;
  const idx = attr(ph, 'idx');
  return { type: attr(ph, 'type') ?? 'obj', ...(idx !== undefined ? { idx } : {}) };
}

/** The shape tree of a part and the part root, where `mc:Choice` prefixes may be declared. */
export function treeOf(
  pkg: PptxPackage,
  part: string | undefined
): { tree?: Element; root?: Element } {
  if (!part) return {};
  const tree = shapeTreeOf(pkg, part);
  const root = xmlPartOf(pkg, part)?.root;
  return { ...(tree ? { tree } : {}), ...(root ? { root } : {}) };
}

/** A part's content elements, split into placeholders and the rest. */
export interface PartElements {
  readonly tree: Element;
  readonly placeholders: readonly PlaceholderEntry[];
  /** The content elements that are not placeholders, in document order. */
  readonly others: readonly Element[];
}

/**
 * The content of a part by its root, which an edit of the part replaces: a layout or a master
 * splits into placeholders and content once, not on each resolution of each slide.
 */
const partElements = new WeakMap<Element, PartElements>();

export function partElementsOf(
  pkg: PptxPackage,
  part: string | undefined
): PartElements | undefined {
  const { tree, root } = treeOf(pkg, part);
  if (!tree) return undefined;
  const hit = root ? partElements.get(root) : undefined;
  if (hit && hit.tree === tree) return hit;
  const placeholders: PlaceholderEntry[] = [];
  const others: Element[] = [];
  for (const el of contentElements(tree, root)) {
    const key = placeholderOf(el);
    if (key) placeholders.push({ ...key, el });
    else others.push(el);
  }
  const out = { tree, placeholders, others };
  if (root) partElements.set(root, out);
  return out;
}

export function collectPlaceholders(
  pkg: PptxPackage,
  part: string | undefined
): readonly PlaceholderEntry[] {
  return partElementsOf(pkg, part)?.placeholders ?? [];
}

/** The master placeholder type a layout/slide placeholder type falls back to. */
export function masterTypeFor(type: string): string {
  if (type === 'ctrTitle') return 'title';
  if (type === 'subTitle' || type === 'obj' || type === 'tbl' || type === 'chart' || type === 'pic')
    return 'body';
  return type;
}

/** The first entry of each idx, type, and master type of a placeholder list. */
interface PlaceholderIndex {
  readonly byIdx: ReadonlyMap<string, Element>;
  readonly byType: ReadonlyMap<string, Element>;
  readonly byMasterType: ReadonlyMap<string, Element>;
}

/**
 * Lookups by key, built once per list: a slide can hold thousands of placeholders that match
 * nothing, and a scan per placeholder made that quadratic.
 */
const placeholderIndexes = new WeakMap<readonly PlaceholderEntry[], PlaceholderIndex>();

function placeholderIndex(list: readonly PlaceholderEntry[]): PlaceholderIndex {
  let index = placeholderIndexes.get(list);
  if (!index) {
    const byIdx = new Map<string, Element>();
    const byType = new Map<string, Element>();
    const byMasterType = new Map<string, Element>();
    for (const p of list) {
      if (p.idx !== undefined && !byIdx.has(p.idx)) byIdx.set(p.idx, p.el);
      if (!byType.has(p.type)) byType.set(p.type, p.el);
      const master = masterTypeFor(p.type);
      if (!byMasterType.has(master)) byMasterType.set(master, p.el);
    }
    index = { byIdx, byType, byMasterType };
    placeholderIndexes.set(list, index);
  }
  return index;
}

export function findLayoutPlaceholder(
  key: PlaceholderKey,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  const index = placeholderIndex(list);
  if (key.idx !== undefined && key.idx !== '0') {
    const byIdx = index.byIdx.get(key.idx);
    if (byIdx) return byIdx;
  }
  return (
    index.byType.get(key.type) ??
    (key.type !== 'obj' ? index.byMasterType.get(masterTypeFor(key.type)) : undefined)
  );
}

export function findMasterPlaceholder(
  type: string,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  const want = masterTypeFor(type);
  const index = placeholderIndex(list);
  return index.byType.get(want) ?? index.byMasterType.get(want);
}

/**
 * The placeholders a slide element inherits from, nearest first: its layout placeholder, then
 * the master placeholder, matched as `resolveSlide` matches them. Empty for an element that is
 * not a placeholder. An edit that starts an own value from the inherited one reads these.
 */
export function inheritedPlaceholders(pkg: PptxPackage, slidePart: string, el: Element): Element[] {
  const ph = placeholderOf(el);
  if (!ph) return [];
  // A layout placeholder in the slide master view inherits from its master; a master's from none.
  const template = templateRefs(pkg).find((t) => t.partName === slidePart);
  if (template?.template === 'master') return [];
  if (template) {
    const masterPh = findMasterPlaceholder(ph.type, collectPlaceholders(pkg, template.masterPart));
    return masterPh ? [masterPh] : [];
  }
  const ref = presentationInfo(pkg).slides.find((s) => s.partName === slidePart);
  if (!ref) return [];
  const layoutPh = findLayoutPlaceholder(ph, collectPlaceholders(pkg, ref.layoutPart));
  const masterPh = findMasterPlaceholder(
    layoutPh ? (placeholderOf(layoutPh)?.type ?? ph.type) : ph.type,
    collectPlaceholders(pkg, ref.masterPart)
  );
  return [layoutPh, masterPh].filter((e): e is Element => e !== undefined);
}

/**
 * The layout placeholder that a placeholder on slide `index` inherits from (the matching rule of
 * `findLayoutPlaceholder`), or undefined. The editor reads its prompt text from it. @internal
 */
export function layoutPlaceholderOf(
  pkg: PptxPackage,
  index: number,
  el: Element
): Element | undefined {
  const ref = slideRefAt(pkg, index);
  const key = placeholderOf(el);
  // A template's own placeholders are the prompts: they have no layout placeholder above them.
  if (!ref || !key || isTemplateRef(ref)) return undefined;
  return findLayoutPlaceholder(key, collectPlaceholders(pkg, ref.layoutPart));
}
