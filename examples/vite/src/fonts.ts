// The demo's fonts, as the DOCX Editor demo sets them up: the packaged substitutes first (the
// package's own bytes, no third party), then the pinned Google Fonts catalog from
// `@docx-editor.dev/fonts/google` for every family they do not cover. The engine never fetches
// fonts by itself; this host opts in. A requested family is only a lookup key into a closed
// catalog, every face is pinned to an immutable commit with a SHA-256 hash, and nothing is
// fetched until a deck or a pick names a catalogued family.

import { packagedFonts } from '@docx-editor.dev/fonts';
import {
  GOOGLE_FONT_FAMILIES,
  GOOGLE_METRIC_SUBSTITUTES,
  googleFonts,
} from '@docx-editor.dev/fonts/google';
import { useFonts } from '@pptx-editor.dev/react';

/** Stand-ins (Carlito for Calibri, …) are implementation faces, not choices. */
const standIns = new Set(Object.values(GOOGLE_METRIC_SUBSTITUTES).map((f) => f.toLowerCase()));

/** The catalog families the font picker offers. */
export const CATALOG_FAMILIES: readonly string[] = GOOGLE_FONT_FAMILIES.filter(
  (f) => !standIns.has(f.toLowerCase())
);

/** The demo's `fonts` prop: one stable value for the editor's life. */
export function useDemoFonts(): ReturnType<typeof useFonts> {
  return useFonts(packagedFonts(), googleFonts());
}
