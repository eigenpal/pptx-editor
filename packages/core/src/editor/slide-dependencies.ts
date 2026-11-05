// What one slide's rendering reads from the package, by object identity.
//
// The package is persistent and structurally shared: an edit replaces only the parts and the
// element paths it changes, and every other part keeps its objects. So a slide whose
// dependencies are the same objects as before lays out and paints exactly as before, and its
// layout and thumbnail can be reused. This is the cache key of `DocumentSession.layout` and of
// the slide summaries' `key` (thumbnails).
//
// The list follows what `resolveSlide`, `layoutSlide`, and `paintSlide` read:
//
// - the slide, layout, and master parts and their relationship parts (placeholders, list
//   styles, color maps, backgrounds, picture targets),
// - the theme part,
// - `p:sldSz` and `p:defaultTextStyle` of the presentation part (not the whole part: its slide
//   list changes when a slide is added, moved, or deleted, and no slide reads it),
// - the table styles part,
// - every other internal target of those relationships (pictures and other media), except the
//   structural ones: sibling layouts, the master, and the notes slide,
// - the slide index (slide-number fields).
//
// XML parts count by their root element, not by the part object: undo can re-wrap a part to
// raise its id counter without changing its content.

import {
  child,
  presentationInfo,
  REL,
  relationshipsOf,
  relTargetByType,
  relsPartName,
  xmlPartOf,
  type PptxPackage,
} from '../store/index.ts';

/** Relationship types whose targets are other slides' business, not this slide's pixels. */
const STRUCTURAL = new Set<string>([
  REL.slideLayout,
  REL.slideMaster,
  REL.notesSlide,
  REL.slide,
  REL.theme,
]);

/** The content object of a part: the root element of an XML part, the bytes of a binary one. */
function contentOf(pkg: PptxPackage, partName: string | undefined): unknown {
  if (!partName) return undefined;
  const part = pkg.parts.get(partName);
  if (!part) return undefined;
  return part.kind === 'xml' ? part.xml.root : part.bytes;
}

/**
 * Everything slide `index` renders from, in a fixed order, or undefined for a slide that does
 * not exist. Two lists with the same entries (by identity) render the same slide.
 */
export function slideDependencies(pkg: PptxPackage, index: number): readonly unknown[] | undefined {
  const info = presentationInfo(pkg);
  const ref = info.slides[index];
  if (!ref) return undefined;
  const presentation = xmlPartOf(pkg, info.presentationPart)?.root;
  const out: unknown[] = [
    index,
    ref.partName,
    contentOf(pkg, ref.themePart),
    child(presentation, 'p:sldSz'),
    child(presentation, 'p:defaultTextStyle'),
    contentOf(pkg, relTargetByType(pkg, info.presentationPart, REL.tableStyles)),
  ];
  for (const source of [ref.partName, ref.layoutPart, ref.masterPart]) {
    out.push(source, contentOf(pkg, source));
    if (!source) continue;
    out.push(contentOf(pkg, relsPartName(source)));
    for (const rel of relationshipsOf(pkg, source))
      if (!rel.external && !STRUCTURAL.has(rel.type)) out.push(contentOf(pkg, rel.target));
  }
  return out;
}

/** Whether two dependency lists hold the same entries. */
export function sameDependencies(
  a: readonly unknown[] | undefined,
  b: readonly unknown[] | undefined
): boolean {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

// One number per object, for the whole page: keys of two documents never collide.
const ids = new WeakMap<object, number>();
let nextId = 0;

function idOf(value: unknown): string {
  if (value === undefined) return '-';
  if (typeof value !== 'object' || value === null) return JSON.stringify(value);
  let id = ids.get(value);
  if (id === undefined) {
    id = ++nextId;
    ids.set(value, id);
  }
  return `#${id}`;
}

/**
 * A string that changes whenever slide `index` may render differently: its dependencies, or
 * the font generation (`DocumentSession.setHostFaces` re-measures every slide).
 */
export function renderKey(
  pkg: PptxPackage,
  index: number,
  fontGeneration: number
): string | undefined {
  const deps = slideDependencies(pkg, index);
  if (!deps) return undefined;
  let key = `f${fontGeneration}`;
  for (const d of deps) key += `.${idOf(d)}`;
  return key;
}
