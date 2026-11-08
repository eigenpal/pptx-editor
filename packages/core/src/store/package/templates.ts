// The slide masters and slide layouts of a presentation, as the slide master view shows them:
// each master in `p:sldMasterIdLst` order, then its layouts in `p:sldLayoutIdLst` order
// (ECMA-376 Part 1, 19.3.1.41 and 19.3.1.42). A template has an index past the slides
// (`TEMPLATE_BASE`), so the code that reads a slide by index reads a layout or a master the
// same way (`slideRefAt`). Every lookup follows relationships, never file-name conventions.

import { relationshipsOf, xmlPartOf, type PptxPackage } from './package.ts';
import { presentationInfo, type SlideRef } from './presentation.ts';
import { attr, child, children, path } from './query.ts';
import { REL } from './namespaces.ts';

/**
 * A slide master or a slide layout, as a slide of the slide master view. Its `partName` is the
 * master or layout part. A layout's `masterPart` is its master; a master has none, and neither
 * has a `layoutPart`.
 */
export interface TemplateRef extends SlideRef {
  readonly template: 'master' | 'layout';
  /** `p:cSld/@name`, or the theme name of a master without one. */
  readonly name: string;
}

/**
 * The first template index. Slide indices stay below it: a deck never holds this many slides
 * (the open limits allow far fewer).
 */
export const TEMPLATE_BASE = 1 << 24;

/** Whether `index` names a template, not a slide. */
export const isTemplateIndex = (index: number): boolean => index >= TEMPLATE_BASE;

/** Whether `ref` is a slide master or a slide layout. */
export const isTemplateRef = (ref: SlideRef | undefined): ref is TemplateRef =>
  !!ref && 'template' in ref;

const cache = new WeakMap<PptxPackage, readonly TemplateRef[]>();

/** The deck's masters, each followed by its layouts. Memoized per package identity. */
export function templateRefs(pkg: PptxPackage): readonly TemplateRef[] {
  const hit = cache.get(pkg);
  if (hit) return hit;
  const info = presentationInfo(pkg);
  const out: TemplateRef[] = [];
  const seen = new Set<string>();
  const target = (source: string, rId: string | undefined, type: string): string | undefined =>
    relationshipsOf(pkg, source).find((r) => r.id === rId && r.type === type && !r.external)
      ?.target;
  const base = { sldId: '', rId: '', hidden: false, nodeId: '' };
  const presentation = xmlPartOf(pkg, info.presentationPart)?.root;
  for (const m of children(path(presentation, 'p:sldMasterIdLst'), 'p:sldMasterId')) {
    const master = target(info.presentationPart, attr(m, 'r:id'), REL.slideMaster);
    const root = xmlPartOf(pkg, master)?.root;
    if (!master || root?.local !== 'sldMaster' || seen.has(master)) continue;
    seen.add(master);
    const themePart = relationshipsOf(pkg, master).find(
      (r) => r.type === REL.theme && !r.external
    )?.target;
    const themeName = attr(xmlPartOf(pkg, themePart)?.root, 'name') ?? '';
    out.push({
      ...base,
      partName: master,
      ...(themePart ? { themePart } : {}),
      template: 'master',
      name: attr(child(root, 'p:cSld'), 'name') || themeName,
    });
    for (const l of children(child(root, 'p:sldLayoutIdLst'), 'p:sldLayoutId')) {
      const layout = target(master, attr(l, 'r:id'), REL.slideLayout);
      const layoutRoot = xmlPartOf(pkg, layout)?.root;
      if (!layout || layoutRoot?.local !== 'sldLayout' || seen.has(layout)) continue;
      seen.add(layout);
      out.push({
        ...base,
        partName: layout,
        masterPart: master,
        ...(themePart ? { themePart } : {}),
        template: 'layout',
        name: attr(child(layoutRoot, 'p:cSld'), 'name') ?? '',
      });
    }
  }
  cache.set(pkg, out);
  return out;
}

/** The index of a template part in the slide master view, or -1. */
export function templateIndexOf(pkg: PptxPackage, partName: string): number {
  const at = templateRefs(pkg).findIndex((t) => t.partName === partName);
  return at < 0 ? -1 : TEMPLATE_BASE + at;
}

/**
 * The slide at `index`, or the template at `index` for a template index. Code that reads a
 * part by index (the cascade, the session, the commands) reads it here, so the slide master view
 * edits a layout or a master as a slide.
 */
export function slideRefAt(pkg: PptxPackage, index: number): SlideRef | undefined {
  return isTemplateIndex(index)
    ? templateRefs(pkg)[index - TEMPLATE_BASE]
    : presentationInfo(pkg).slides[index];
}
