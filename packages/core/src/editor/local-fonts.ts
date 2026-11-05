// Read fonts installed on the user's computer through the Local Font Access API, for the
// families a deck needs. Browser-only, feature-detected, and user-initiated: the browser asks
// for permission, nothing is uploaded, and nothing persists beyond the page. The faces come
// back as `FontSource`s (`createFontSource`), ready for `setFonts` with `withFontOrigin`.

import type { FontSource } from '../contracts/index.ts';
import { createFontSource } from './load-fonts.ts';

/** The subset of the Local Font Access API this module uses. */
interface LocalFontData {
  readonly family: string;
  readonly style: string;
  blob(): Promise<Blob>;
}

type QueryLocalFonts = () => Promise<readonly LocalFontData[]>;

/** Whether the browser offers the Local Font Access API. @public */
export function canQueryLocalFonts(): boolean {
  return (
    typeof globalThis !== 'undefined' &&
    typeof (globalThis as { queryLocalFonts?: unknown }).queryLocalFonts === 'function'
  );
}

/** Why reading local fonts produced nothing. @public */
export type LocalFontsOutcome =
  | { readonly ok: true; readonly sources: readonly FontSource[] }
  | { readonly ok: false; readonly reason: 'unsupported' | 'denied' | 'none' };

const MAX_FONT_BYTES = 64 * 1024 * 1024;

function styleOf(style: string): { bold: boolean; italic: boolean; regular: boolean } | null {
  const s = style.toLowerCase();
  const italic = /italic|oblique/.test(s);
  const weight = s.replace(/italic|oblique/g, '').trim();
  if (
    weight === '' ||
    weight === 'regular' ||
    weight === 'normal' ||
    weight === 'roman' ||
    weight === 'book'
  )
    return { bold: false, italic, regular: true };
  if (weight === 'bold') return { bold: true, italic, regular: true };
  // Light, semibold, black, … are separate families in the deck's font model; skip them.
  return null;
}

/**
 * Read the regular, bold, italic, and bold italic faces of `families` from the computer.
 * Must be called from a user gesture (a click). Font collections (`.ttc`) are skipped.
 *
 * @public
 */
export async function queryLocalFontSources(
  families: readonly string[]
): Promise<LocalFontsOutcome> {
  if (!canQueryLocalFonts()) return { ok: false, reason: 'unsupported' };
  const wanted = new Map(families.map((f) => [f.toLowerCase(), f]));
  let available: readonly LocalFontData[];
  try {
    available = await (
      globalThis as unknown as { queryLocalFonts: QueryLocalFonts }
    ).queryLocalFonts();
  } catch {
    return { ok: false, reason: 'denied' };
  }
  const sources: FontSource[] = [];
  const seen = new Set<string>();
  for (const font of available) {
    const family = wanted.get(font.family.toLowerCase());
    const style = styleOf(font.style);
    if (!family || !style) continue;
    const key = `${family}|${style.bold}|${style.italic}`;
    if (seen.has(key)) continue;
    try {
      const blob = await font.blob();
      if (blob.size > MAX_FONT_BYTES) continue;
      const bytes = new Uint8Array(await blob.arrayBuffer());
      // `ttcf`: a font collection, which FontFace cannot load.
      if (bytes[0] === 0x74 && bytes[1] === 0x74 && bytes[2] === 0x63 && bytes[3] === 0x66)
        continue;
      const made = createFontSource(bytes, {
        family,
        weight: style.bold ? 700 : 400,
        style: style.italic ? 'italic' : 'normal',
      });
      // A face that fails the byte or structure checks is skipped, as an unreadable one.
      if (!('source' in made)) continue;
      seen.add(key);
      sources.push(made.source);
    } catch {
      // One unreadable face does not block the others.
    }
  }
  return sources.length ? { ok: true, sources } : { ok: false, reason: 'none' };
}
