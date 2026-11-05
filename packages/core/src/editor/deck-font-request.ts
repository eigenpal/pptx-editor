// The font request a deck makes of the host's font origins (PPTX only; the composition itself
// is DOCX Editor's, `layout/font-resolver.ts`). The families are the theme fonts of every theme
// a slide master uses, then every typeface the slides, layouts, and masters name. The default
// family is the minor Latin font of the first master's theme: body text measures with it.

import {
  attr,
  NS,
  presentationInfo,
  type Element,
  type PptxPackage,
  type XmlPart,
  xmlPartOf,
} from '../store/index.ts';
import { parseTheme, type FontCollection } from '../cascade/index.ts';
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
const cache = new WeakMap<XmlPart, readonly string[]>();

/** Typefaces a part names, in document order. Theme references (`+mn-lt`) are not names. */
function partTypefaces(part: XmlPart): readonly string[] {
  const hit = cache.get(part);
  if (hit) return hit;
  const out: string[] = [];
  const walk = (el: Element, depth: number): void => {
    if (depth > 256) return;
    if (el.ns === NS.a && TYPEFACE_ELEMENTS.has(el.local)) {
      const face = attr(el, 'typeface');
      if (face && !face.startsWith('+')) out.push(face);
    }
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  walk(part.root, 0);
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
  const themes = [...new Set(presentationInfo(pkg).slides.map((s) => s.themePart))];
  for (const name of themes)
    for (const collection of themeOf(pkg, name)) {
      add(collection.latin);
      add(collection.ea);
      add(collection.cs);
    }
  for (const part of pkg.parts.values())
    if (part.kind === 'xml' && FAMILY_PARTS.test(part.partName))
      for (const face of partTypefaces(part.xml)) add(face);
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
