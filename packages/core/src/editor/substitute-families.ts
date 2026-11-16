// The font families a deck names, and the packaged substitute families that cover them. The
// session loads the substitutes of the opening slide before its first paint and the rest after
// it (`DocumentSession.addMissingSubstitutes`), so a large deck does not hold its first slide.

import { ALL_DEFAULT_FONT_FAMILIES } from '@docx-editor.dev/fonts';
import {
  attr,
  changedPartNames,
  deferredParseCount,
  isDeferredPending,
  NS,
  presentationInfo,
  type Element,
  type PackagePart,
  type PptxPackage,
  type PresentationInfo,
  type XmlPart,
} from '../store/index.ts';
import { faceStylesOf, partStyles, type FaceStyle } from './face-styles.ts';

/** Elements whose `typeface` names a family layout measures with. */
const TYPEFACE_ELEMENTS = new Set(['latin', 'sym', 'buFont']);
const familyCache = new WeakMap<XmlPart, ReadonlySet<string>>();

function partFamilies(part: XmlPart): ReadonlySet<string> {
  const hit = familyCache.get(part);
  if (hit) return hit;
  const out = new Set<string>();
  const walk = (el: Element, depth: number): void => {
    if (depth > 256) return;
    if (el.ns === NS.a && TYPEFACE_ELEMENTS.has(el.local)) {
      const face = attr(el, 'typeface');
      if (face && !face.startsWith('+') && face.length <= 64) out.add(face);
    }
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  walk(part.root, 0);
  familyCache.set(part, out);
  return out;
}

const FAMILY_PARTS = /^\/ppt\/(slides|slideLayouts|slideMasters)\//;

/** The families of the parts that name measured text, for the last package read. */
interface FamilyCount {
  /** `deferredParseCount()` when counted: a slide that parses changes no part identity. */
  readonly parsed: number;
  readonly slides: PresentationInfo['slides'];
  readonly themes: ReadonlySet<string | undefined>;
  /** The number of parts that name each family, by its exact spelling. */
  readonly counts: ReadonlyMap<string, number>;
  readonly list: readonly string[];
}

const familyCounts = new WeakMap<PptxPackage, FamilyCount>();
/** The last package counted, held weakly: a closed deck is not kept alive for it. */
let lastFamilyPackage: WeakRef<PptxPackage> | null = null;
const EMPTY_FAMILIES: ReadonlySet<string> = new Set();

/** The families `part` adds to the count, or none. */
function countedFamilies(
  part: PackagePart | undefined,
  themes: ReadonlySet<string | undefined>
): ReadonlySet<string> {
  if (part?.kind !== 'xml' || !(FAMILY_PARTS.test(part.partName) || themes.has(part.partName)))
    return EMPTY_FAMILIES;
  // A slide still opening names its families once it parses (`progressive-open.ts`).
  if (isDeferredPending(part)) return EMPTY_FAMILIES;
  return partFamilies(part.xml);
}

/**
 * The sorted list, one entry per case-folded family, or null when two spellings of one family
 * remain: a full pass then picks the spelling of the first part that names it.
 */
function familyList(counts: ReadonlyMap<string, number>): string[] | null {
  const folded = new Set<string>();
  for (const f of counts.keys()) {
    const key = f.toLowerCase();
    if (folded.has(key)) return null;
    folded.add(key);
  }
  return [...counts.keys()].sort((a, b) => a.localeCompare(b));
}

/**
 * Every family the deck names for measured text: run fonts, symbol fonts, bullet fonts, and
 * the theme's font scheme. Cached per part identity, so an edit rescans only the parts it
 * changed. After an edit that keeps the slide list, the count of the last package changes by
 * the changed parts only, so a keystroke walks no other part. The result is the one a full
 * pass gives.
 */
export function namedFamilies(pkg: PptxPackage): string[] {
  const info = presentationInfo(pkg);
  const parsed = deferredParseCount();
  const hit = familyCounts.get(pkg);
  if (hit?.parsed === parsed) return rememberFamilies(pkg, hit);
  const lastPkg = lastFamilyPackage?.deref();
  const last = lastPkg && familyCounts.get(lastPkg);
  const changed =
    lastPkg && last && last.parsed === parsed && last.slides === info.slides
      ? changedPartNames(lastPkg, pkg, 64)
      : null;
  if (last && changed) {
    const counts = new Map(last.counts);
    const add = (families: ReadonlySet<string>, by: number): void => {
      for (const f of families) {
        const n = (counts.get(f) ?? 0) + by;
        if (n > 0) counts.set(f, n);
        else counts.delete(f);
      }
    };
    for (const name of changed) {
      add(countedFamilies(lastPkg?.parts.get(name), last.themes), -1);
      add(countedFamilies(pkg.parts.get(name), last.themes), 1);
    }
    const list = familyList(counts);
    if (list) {
      return rememberFamilies(pkg, {
        parsed,
        slides: info.slides,
        themes: last.themes,
        counts,
        list,
      });
    }
  }
  // Only the themes slide masters use: a notes or handout master theme never reaches a slide.
  const themes = new Set(info.slides.map((s) => s.themePart));
  const counts = new Map<string, number>();
  const byKey = new Map<string, string>();
  for (const part of pkg.parts.values())
    for (const f of countedFamilies(part, themes)) {
      counts.set(f, (counts.get(f) ?? 0) + 1);
      if (!byKey.has(f.toLowerCase())) byKey.set(f.toLowerCase(), f);
    }
  const list = [...byKey.values()].sort((a, b) => a.localeCompare(b));
  return rememberFamilies(pkg, { parsed, slides: info.slides, themes, counts, list });
}

function rememberFamilies(pkg: PptxPackage, count: FamilyCount): string[] {
  familyCounts.set(pkg, count);
  lastFamilyPackage = new WeakRef(pkg);
  return [...count.list];
}

/** A family that `@docx-editor.dev/fonts` ships a substitute for. */
export type DefaultFamily = (typeof ALL_DEFAULT_FONT_FAMILIES)[number];

/**
 * Which packaged substitute family covers a typeface the deck names. A deck can name a packaged
 * open face itself (Carlito, Caladea, TeX Gyre Adventor): it loads the family that ships it, so
 * the face draws as itself and not as the Arial fallback.
 */
function substituteFamilyFor(typeface: string): DefaultFamily {
  const t = typeface.toLowerCase();
  if (t.startsWith('calibri') || t.startsWith('aptos') || t === 'carlito') return 'Calibri';
  if (t.startsWith('cambria') || t === 'caladea') return 'Cambria';
  if (t.startsWith('tex gyre adventor')) return 'Century Gothic';
  if (t.startsWith('times') || t.includes('serif') || t.startsWith('georgia'))
    return 'Times New Roman';
  if (t.startsWith('courier') || t.includes('mono') || t.startsWith('consolas'))
    return 'Courier New';
  if (t.startsWith('century gothic')) return 'Century Gothic';
  return 'Arial';
}

/** The substitute families a deck needs: every typeface it names, plus the Arial fallback. */
export function substituteFamiliesFor(pkg: PptxPackage): DefaultFamily[] {
  const families = new Set<DefaultFamily>(['Arial']);
  for (const f of namedFamilies(pkg)) families.add(substituteFamilyFor(f));
  return [...families];
}

/**
 * The substitute families that slide `index` needs: the families its slide, layout, and master
 * parts and its theme name. A family with no substitute of its own needs Arial, and so does a
 * slide that names no family. The Arial fallback for missing characters loads with the rest of
 * the deck (`substituteFamiliesFor`). A part that is still opening names no family.
 */
export function slideSubstituteFamilies(pkg: PptxPackage, index: number): DefaultFamily[] {
  const families = new Set<DefaultFamily>();
  const ref = presentationInfo(pkg).slides[index];
  const themes = new Set([ref?.themePart]);
  for (const name of slideParts(ref))
    for (const family of countedFamilies(name ? pkg.parts.get(name) : undefined, themes))
      families.add(substituteFamilyFor(family));
  return families.size > 0 ? [...families] : ['Arial'];
}

/** The parts that slide `ref` draws text from: its slide, layout, master, and theme. */
function slideParts(ref: PresentationInfo['slides'][number] | undefined): (string | undefined)[] {
  return ref ? [ref.partName, ref.layoutPart, ref.masterPart, ref.themePart] : [];
}

/** The part that holds the deck's table styles (ECMA-376 Part 1, 20.1.4.2.27 `tblStyleLst`). */
const TABLE_STYLES = '/ppt/tableStyles.xml';

/**
 * The faces that the text of slide `index` asks for (`face-styles.ts`), the same for each part
 * of its slide, layout, master, and theme, and for the deck's table styles.
 */
export function slideFaceStyles(pkg: PptxPackage, index: number): FaceStyle[] {
  const names = [...slideParts(presentationInfo(pkg).slides[index]), TABLE_STYLES];
  return faceStylesOf(
    names.flatMap((name) => {
      const part = name ? pkg.parts.get(name) : undefined;
      return part?.kind === 'xml' && !isDeferredPending(part) ? [partStyles(part.xml)] : [];
    })
  );
}

/** A packaged substitute family and the faces of it to load. */
export interface SubstituteFaces {
  readonly family: DefaultFamily;
  readonly faces: readonly FaceStyle[];
}

/** The substitute families that slide `index` needs, each with the faces its text asks for. */
export function slideSubstituteFaces(pkg: PptxPackage, index: number): SubstituteFaces[] {
  const faces = slideFaceStyles(pkg, index);
  return slideSubstituteFamilies(pkg, index).map((family) => ({ family, faces }));
}
