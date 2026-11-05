// Whether a saved presentation package is whole. The automation tests of `core` and of
// `editor-api` both check every package they save with it, so one list of rules covers both.
//
// The rules: the package reads back with the bounded reader; every internal relationship names
// a part (ECMA-376 Part 2, 9.3); every part has a content type (Part 2, 10.1.2); every `r:`
// reference of a slide, layout, master, or notes part names a relationship of its part; master
// and layout ids are unique and at least 2^31 (ECMA-376 Part 1, 19.2.1.27 and 19.3.1.42);
// slide ids are unique (19.2.1.33); each slide has a layout, a master, and a theme; and the
// shape ids (`p:cNvPr/@id`) of each slide are unique (19.3.1.12).

import {
  attr,
  child,
  elements,
  NS,
  presentationInfo,
  readPptx,
  relationshipsOf,
  contentTypeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';

/** The parts whose `r:` references must name a relationship of the same part. */
const REFERENCING = new Set(['sld', 'sldLayout', 'sldMaster', 'notes', 'notesMaster']);

/** Every `r:` attribute value of a part's tree. */
export function relationshipReferences(root: Element | undefined): string[] {
  const out: string[] = [];
  const walk = (el: Element | undefined, depth: number): void => {
    if (!el || depth > 256) return;
    for (const a of el.attrs) if (a.ns === NS.r) out.push(a.value);
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/** Every `p:cNvPr/@id` of a part's tree. */
function shapeIds(root: Element | undefined): string[] {
  const out: string[] = [];
  const walk = (el: Element | undefined, depth: number): void => {
    if (!el || depth > 256) return;
    if (el.local === 'cNvPr' && el.ns === NS.p) {
      const id = attr(el, 'id');
      if (id !== undefined) out.push(id);
    }
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

function duplicates(values: readonly string[]): string[] {
  const seen = new Set<string>();
  return [...new Set(values.filter((v) => seen.size === seen.add(v).size))];
}

/** The package and the rules it breaks, as plain sentences. An empty list is a whole package. */
export function packageIntegrity(bytes: Uint8Array): { pkg?: PptxPackage; issues: string[] } {
  const read = readPptx(bytes);
  if (!read.ok) return { issues: [`the package does not read back: ${JSON.stringify(read.error)}`] };
  const pkg = read.pkg;
  const issues: string[] = [];
  for (const name of ['/', ...pkg.parts.keys()]) {
    if (name.endsWith('.rels')) continue;
    for (const rel of relationshipsOf(pkg, name))
      if (!rel.external && !pkg.parts.has(rel.target!))
        issues.push(`${name} ${rel.id} names a missing part ${rel.target}`);
    if (name !== '/' && name !== '/[Content_Types].xml' && contentTypeOf(pkg, name) === undefined)
      issues.push(`${name} has no content type`);
    const root = xmlPartOf(pkg, name)?.root;
    if (root && REFERENCING.has(root.local)) {
      const ids = new Set(relationshipsOf(pkg, name).map((r) => r.id));
      for (const id of relationshipReferences(root))
        if (!ids.has(id)) issues.push(`${name} references a missing relationship ${id}`);
    }
  }
  const info = presentationInfo(pkg);
  const presentation = xmlPartOf(pkg, info.presentationPart)?.root;
  const masterIds: string[] = [];
  for (const m of elements(child(presentation, 'p:sldMasterIdLst'))) {
    masterIds.push(attr(m, 'id') ?? '');
    const target = relationshipsOf(pkg, info.presentationPart).find(
      (r) => r.id === attr(m, 'r:id')
    )?.target;
    for (const l of elements(child(xmlPartOf(pkg, target)?.root, 'p:sldLayoutIdLst')))
      masterIds.push(attr(l, 'id') ?? '');
  }
  for (const id of duplicates(masterIds)) issues.push(`master or layout id ${id} repeats`);
  for (const id of masterIds)
    if (!(Number(id) >= 2_147_483_648)) issues.push(`master or layout id ${id} is below 2^31`);
  for (const id of duplicates(info.slides.map((s) => s.sldId))) issues.push(`slide id ${id} repeats`);
  for (const slide of info.slides) {
    if (!(slide.layoutPart && slide.masterPart && slide.themePart))
      issues.push(`${slide.partName} lacks its layout, master, or theme`);
    for (const id of duplicates(shapeIds(xmlPartOf(pkg, slide.partName)?.root)))
      issues.push(`${slide.partName} repeats shape id ${id}`);
  }
  return { pkg, issues };
}
