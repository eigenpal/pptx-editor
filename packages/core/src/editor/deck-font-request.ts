// The font request a deck makes of the host's font origins (PPTX only; the composition itself
// is DOCX Editor's, `layout/font-resolver.ts`). The families are the theme fonts of every theme
// a slide master uses, with the script fonts of the languages the deck names, then every
// typeface the slides, layouts, masters, and the presentation's default text style name. The default
// family is the minor Latin font of the first master's theme: body text measures with it.

import {
  attr,
  isDeferredPending,
  NS,
  presentationInfo,
  type Element,
  type PptxPackage,
  type XmlPart,
  xmlPartOf,
} from '../store/index.ts';
import { collectionScriptFont, parseTheme, type FontCollection } from '../cascade/index.ts';
import {
  DEFAULT_FONT,
  MAX_RESOLVER_FAMILIES,
  type FontResolutionRequest,
} from '../layout/font-composition.ts';
import { FONT_NAME } from '../automation/commands/font-format.ts';

/**
 * Elements whose `typeface` names a family (ECMA-376 Part 1, 21.1.2.3.7 `latin`, 21.1.2.3.3
 * `ea`, 21.1.2.3.1 `cs`, 21.1.2.3.10 `sym`, 21.1.2.4.6 `buFont`).
 */
const TYPEFACE_ELEMENTS = new Set(['latin', 'ea', 'cs', 'sym', 'buFont']);
const FAMILY_PARTS = /^\/ppt\/(slides|slideLayouts|slideMasters)\//;
const cache = new WeakMap<XmlPart, PartFonts>();

/** What a part names: its typefaces in document order, and its run languages. */
interface PartFonts {
  readonly faces: readonly string[];
  /** `lang` and `altLang` of its run properties, which pick theme script fonts. */
  readonly languages: ReadonlySet<string>;
}

/** The most run languages read from one part. */
const MAX_PART_LANGUAGES = 32;
const RUN_PROPERTIES = new Set(['rPr', 'defRPr', 'endParaRPr']);

/** Typefaces a part names, in document order. Theme references (`+mn-lt`) are not names. */
function partFonts(part: XmlPart): PartFonts {
  const hit = cache.get(part);
  if (hit) return hit;
  const faces: string[] = [];
  const languages = new Set<string>();
  const walk = (el: Element, depth: number): void => {
    if (depth > 256) return;
    if (el.ns === NS.a && TYPEFACE_ELEMENTS.has(el.local)) {
      const face = attr(el, 'typeface');
      if (face && !face.startsWith('+')) faces.push(face);
    }
    if (el.ns === NS.a && RUN_PROPERTIES.has(el.local) && languages.size < MAX_PART_LANGUAGES)
      for (const name of ['lang', 'altLang']) {
        const lang = attr(el, name);
        if (lang) languages.add(lang);
      }
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  walk(part.root, 0);
  const out = { faces, languages };
  cache.set(part, out);
  return out;
}

function themeOf(pkg: PptxPackage, partName: string | undefined): FontCollection[] {
  const part = xmlPartOf(pkg, partName);
  if (!part) return [];
  const theme = parseTheme(part);
  return [theme.majorFont, theme.minorFont];
}

/**
 * The families a deck names, theme fonts first, without duplicates (case-insensitive), each
 * checked with the font picker's name rule, at most `MAX_RESOLVER_FAMILIES`.
 */
export function deckFontFamilies(pkg: PptxPackage): string[] {
  const byKey = new Map<string, string>();
  const add = (family: string | undefined): void => {
    if (!family || byKey.size >= MAX_RESOLVER_FAMILIES || !FONT_NAME.test(family)) return;
    if (!byKey.has(family.toLowerCase())) byKey.set(family.toLowerCase(), family);
  };
  // The presentation part holds `p:defaultTextStyle`, whose `lang` applies to every run that
  // names none.
  const presentationPart = presentationInfo(pkg).presentationPart;
  const parts = [...pkg.parts.values()].flatMap((part) =>
    // A slide still opening is asked for when the open settles (`progressive-open.ts`).
    part.kind === 'xml' &&
    !isDeferredPending(part) &&
    (FAMILY_PARTS.test(part.partName) || part.partName === presentationPart)
      ? [partFonts(part.xml)]
      : []
  );
  const languages = new Set(parts.flatMap((p) => [...p.languages]));
  const themes = [...new Set(presentationInfo(pkg).slides.map((s) => s.themePart))];
  for (const name of themes)
    for (const collection of themeOf(pkg, name)) {
      add(collection.latin);
      add(collection.ea);
      add(collection.cs);
      // The script fonts of the languages the runs name (ECMA-376 Part 1, 20.1.4.1.16 `font`),
      // so East Asian text in a theme with an empty `a:ea` asks for its face.
      for (const lang of languages) {
        if (!collection.ea) add(collectionScriptFont(collection, 'ea', lang));
        if (!collection.cs) add(collectionScriptFont(collection, 'cs', lang));
      }
    }
  for (const p of parts) for (const face of p.faces) add(face);
  return [...byKey.values()];
}

/**
 * The family a run naming no font measures with: the minor Latin font of the first slide
 * master's theme (ECMA-376 Part 1, 20.1.4.1.24 `minorFont`). `DEFAULT_FONT.family` when the
 * theme names none.
 */
export function deckDefaultFamily(pkg: PptxPackage): string {
  const themePart = presentationInfo(pkg).slides.find((s) => s.themePart)?.themePart;
  const part = xmlPartOf(pkg, themePart);
  const latin = part ? parseTheme(part).minorFont.latin : undefined;
  return latin && FONT_NAME.test(latin) ? latin : DEFAULT_FONT.family;
}

/** The request a deck makes of the host's font origins. */
export function deckFontRequest(pkg: PptxPackage, signal?: AbortSignal): FontResolutionRequest {
  return {
    families: deckFontFamilies(pkg),
    defaultFamily: deckDefaultFamily(pkg),
    ...(signal ? { signal } : {}),
  };
}
