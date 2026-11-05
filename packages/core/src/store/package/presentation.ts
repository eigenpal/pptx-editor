// The presentation structure: slide order, the slide → layout → master → theme chain, notes,
// slide size, and embedded font list. Read-only views over the package; every lookup follows
// relationships, never file-name conventions.

import {
  equivalentPartName,
  mainPresentationPart,
  relationshipsOf,
  relTarget,
  relTargetByType,
  repairedPresentationRoot,
  xmlPartOf,
  type PptxPackage,
} from './package.ts';
import { attr, child, children, intAttr, path } from './query.ts';
import { NS, REL } from './namespaces.ts';
import type { Element } from './tree.ts';

/** EMUs per inch, point, and CSS pixel (96 dpi). */
export const EMU_PER_INCH = 914_400;
export const EMU_PER_POINT = 12_700;
export const EMU_PER_PX = 9_525;

/** One slide in presentation order. */
export interface SlideRef {
  /** `p:sldId/@id`, stable across reorders. */
  readonly sldId: string;
  /** The `r:id` in presentation.xml.rels. */
  readonly rId: string;
  readonly partName: string;
  readonly layoutPart?: string;
  readonly masterPart?: string;
  readonly themePart?: string;
  readonly notesPart?: string;
  /** `p:sld/@show="0"`. */
  readonly hidden: boolean;
  /** The `p:sldId` node id in presentation.xml, for reorder ops. */
  readonly nodeId: string;
}

/** One embedded font family with the part of each face. */
export interface EmbeddedFontRef {
  readonly typeface: string;
  readonly faces: readonly {
    readonly style: 'regular' | 'bold' | 'italic' | 'boldItalic';
    readonly partName: string;
  }[];
}

/** The structural summary of a presentation. */
export interface PresentationInfo {
  readonly presentationPart: string;
  /** Slide size in EMU. */
  readonly width: number;
  readonly height: number;
  readonly slides: readonly SlideRef[];
  readonly embeddedFonts: readonly EmbeddedFontRef[];
  readonly notesMasterPart?: string;
}

const cache = new WeakMap<PptxPackage, PresentationInfo>();

/** Default slide size when `p:sldSz` is missing or corrupt: 10 × 7.5 in (4:3). */
const DEFAULT_SIZE = { cx: 9_144_000, cy: 6_858_000 };
/** ECMA-376 `ST_SlideSizeCoordinate` bounds. */
const MIN_SLIDE = 914_400;
const MAX_SLIDE = 51_206_400;

/** Read the presentation structure. Memoized per package identity. */
export function presentationInfo(pkg: PptxPackage): PresentationInfo {
  const hit = cache.get(pkg);
  if (hit) return hit;
  const presentationPart = mainPresentationPart(pkg) ?? '/ppt/presentation.xml';
  const root = xmlPartOf(pkg, presentationPart)?.root;
  const sz = child(root, 'p:sldSz');
  const clamp = (v: number | undefined, d: number): number =>
    v === undefined || v < MIN_SLIDE || v > MAX_SLIDE ? d : v;
  const width = clamp(intAttr(sz, 'cx'), DEFAULT_SIZE.cx);
  const height = clamp(intAttr(sz, 'cy'), DEFAULT_SIZE.cy);
  const slides: SlideRef[] = [];
  // The slide entries as the writer saves them (`repairPresentationRoot`): one entry for each
  // slide part, and a fresh id for a repeated one. A concurrent edit can leave the part
  // otherwise. The rail, the commands, and the saved file count the same slides.
  const entries = mainPresentationPart(pkg) ? (repairedPresentationRoot(pkg) ?? root) : root;
  for (const el of children(path(entries, 'p:sldIdLst'), 'p:sldId')) {
    const rId = attr(el, 'r:id');
    const target = rId ? relTarget(pkg, presentationPart, rId) : undefined;
    // OPC part names fold ASCII case (ECMA-376 Part 2, 6.2.2.3), as the writer's repair does
    // (`liveRelationshipIds`): a target that differs only in case names the part.
    const partName = target ? equivalentPartName(pkg, target) : undefined;
    if (!rId || !partName || !xmlPartOf(pkg, partName)) continue;
    const layoutPart = relTargetByType(pkg, partName, REL.slideLayout);
    const masterPart = layoutPart ? relTargetByType(pkg, layoutPart, REL.slideMaster) : undefined;
    const themePart = masterPart ? relTargetByType(pkg, masterPart, REL.theme) : undefined;
    const notesPart = relTargetByType(pkg, partName, REL.notesSlide);
    const sld = xmlPartOf(pkg, partName)?.root;
    slides.push({
      sldId: attr(el, 'id') ?? '',
      rId,
      partName,
      ...(layoutPart ? { layoutPart } : {}),
      ...(masterPart ? { masterPart } : {}),
      ...(themePart ? { themePart } : {}),
      ...(notesPart ? { notesPart } : {}),
      hidden: attr(sld, 'show') === '0' || attr(sld, 'show') === 'false',
      nodeId: el.id,
    });
  }
  const embeddedFonts: EmbeddedFontRef[] = [];
  for (const ef of children(path(root, 'p:embeddedFontLst'), 'p:embeddedFont')) {
    const typeface = attr(child(ef, 'p:font'), 'typeface');
    if (!typeface) continue;
    const faces: { style: EmbeddedFontRef['faces'][number]['style']; partName: string }[] = [];
    for (const style of ['regular', 'bold', 'italic', 'boldItalic'] as const) {
      const rId = attr(child(ef, `p:${style}`), 'r:id');
      const partName = rId ? relTarget(pkg, presentationPart, rId) : undefined;
      if (partName) faces.push({ style, partName });
    }
    embeddedFonts.push({ typeface, faces });
  }
  const notesMasterId = attr(path(root, 'p:notesMasterIdLst', 'p:notesMasterId'), 'r:id');
  const notesMasterPart = notesMasterId
    ? relTarget(pkg, presentationPart, notesMasterId)
    : undefined;
  const info: PresentationInfo = {
    presentationPart,
    width,
    height,
    slides,
    embeddedFonts,
    ...(notesMasterPart ? { notesMasterPart } : {}),
  };
  cache.set(pkg, info);
  return info;
}

/** The `p:cSld/p:spTree` of a slide, layout, or master part. */
export function shapeTreeOf(pkg: PptxPackage, partName: string | undefined): Element | undefined {
  return path(xmlPartOf(pkg, partName)?.root, 'p:cSld', 'p:spTree');
}

/** Whether a relationship target is an image the engine may decode (never external). */
export function imagePartFor(pkg: PptxPackage, source: string, rId: string): string | undefined {
  const rel = relationshipsOf(pkg, source).find((r) => r.id === rId);
  if (!rel || rel.external) return undefined;
  return rel.target;
}

export { NS };
