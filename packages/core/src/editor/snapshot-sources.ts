// Snapshot fields that keep their reference until their content changes, so a selector on
// them does not re-render on every edit. Moved out of `pptx-editor.ts` (the 1000-line cap).

import type { EditorSnapshot } from '../contracts/index.ts';
import type { PptxPackage } from '../store/index.ts';
import { namedFamilies, type DocumentSession } from './document-session.ts';
import { notesParagraphs, notesText } from '../automation/commands/notes-commands.ts';
import { FONT_NAME, STANDARD_FONT_FAMILIES } from './selection-state.ts';

/** The notes text of a slide, recomputed only when that slide's notes part changes. */
export function createNotesSource(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined,
  slide: number
) => string {
  let cache: { slide: number; content: unknown; text: string } | null = null;
  return (session, pkg, slide) => {
    if (!session || !pkg) return '';
    const part = session.info.slides[slide]?.notesPart;
    const stored = part ? pkg.parts.get(part) : undefined;
    const content = stored?.kind === 'xml' ? stored.xml.root : stored;
    if (cache && cache.slide === slide && cache.content === content && content) return cache.text;
    const text = notesText(pkg, slide);
    cache = { slide, content, text };
    return text;
  };
}

const NO_NOTES: NonNullable<EditorSnapshot['notesFormatted']> = [];

/** The formatted notes of a slide, recomputed only when that slide's notes part changes. */
export function createNotesFormattedSource(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined,
  slide: number
) => NonNullable<EditorSnapshot['notesFormatted']> {
  let cache: { slide: number; content: unknown; value: EditorSnapshot['notesFormatted'] } | null =
    null;
  return (session, pkg, slide) => {
    if (!session || !pkg) return NO_NOTES;
    const part = session.info.slides[slide]?.notesPart;
    const stored = part ? pkg.parts.get(part) : undefined;
    const content = stored?.kind === 'xml' ? stored.xml.root : stored;
    if (!content) return NO_NOTES;
    if (cache && cache.slide === slide && cache.content === content) return cache.value!;
    const value = notesParagraphs(pkg, slide);
    cache = { slide, content, value };
    return value;
  };
}

/** The substituted font families, the same array until the list changes. */
export function createSubstitutionSource(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined
) => EditorSnapshot['substitutedFonts'] {
  let last: EditorSnapshot['substitutedFonts'] = [];
  let of: { pkg: PptxPackage | undefined; fonts: number } = { pkg: undefined, fonts: -1 };
  return (session, pkg) => {
    const fonts = session?.fontGeneration ?? -1;
    if (of.pkg === pkg && of.fonts === fonts) return last;
    of = { pkg, fonts };
    const next = session?.substitutedFamilies() ?? [];
    if (JSON.stringify(next) !== JSON.stringify(last)) last = next;
    return last;
  };
}

/** The font picker's family list, and the host families it adds. */
export interface FontFamilySource {
  /** The list for `session`: the same array until its content can change. */
  list(session: DocumentSession | null): readonly string[];
  /** Replace the host's families. False when `next` names the same families, in order. */
  set(next: readonly string[]): boolean;
}

/**
 * The families the font picker offers: standard, configured, embedded, and the deck's own, one
 * per case-folded name, sorted. The same array until the session, its revision, or the
 * configured families change; the cache keys on the session too, because a revision number
 * repeats in the next document.
 */
export function createFontFamilySource(
  initial: readonly string[],
  supported: () => readonly string[] = () => []
): FontFamilySource {
  let configured = [...initial];
  let cache: {
    session: unknown;
    revision: number;
    fonts: number;
    list: readonly string[];
  } | null = null;
  return {
    set(next) {
      if (next.length === configured.length && next.every((f, i) => f === configured[i]))
        return false;
      configured = [...next];
      cache = null;
      return true;
    },
    list(session) {
      const revision = session?.store.revision ?? -1;
      const fonts = session?.fontGeneration ?? -1;
      if (cache?.session === session && cache.revision === revision && cache.fonts === fonts)
        return cache.list;
      const byFold = new Map<string, string>();
      const add = (f: string | undefined): void => {
        if (f && FONT_NAME.test(f) && !byFold.has(f.toLowerCase())) byFold.set(f.toLowerCase(), f);
      };
      for (const f of STANDARD_FONT_FAMILIES) add(f);
      for (const f of configured) add(f);
      // Families the host's font origins can serve (`supportedFamilies`), as DOCX Editor's
      // `getAvailableFonts()` lists them.
      for (const f of supported()) add(f);
      if (session) {
        for (const f of session.fontReport.faces) if (f.origin === 'embedded') add(f.family);
        for (const f of namedFamilies(session.store.package)) add(f);
      }
      const list = [...byFold.values()].sort((a, b) => a.localeCompare(b));
      cache = { session, revision, fonts, list };
      return list;
    },
  };
}

/** A number per object, the same each time for the same object: a cache key for parts. */
export function createVersionSource(): (part: object) => number {
  const versions = new WeakMap<object, number>();
  let counter = 0;
  return (part) => {
    let v = versions.get(part);
    if (v === undefined) {
      v = ++counter;
      versions.set(part, v);
    }
    return v;
  };
}
