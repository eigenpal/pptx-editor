// The demo's fonts: the Vue twin of examples/vite/src/fonts.ts, as the DOCX Editor demos set
// them up. The packaged substitutes come first (the package's own bytes, no third party), then
// the pinned Google Fonts catalog from `@docx-editor.dev/fonts/google` for every family they do
// not cover. The engine never fetches fonts by itself; this host opts in.

import { packagedFonts } from '@docx-editor.dev/fonts';
import {
  GOOGLE_FONT_FAMILIES,
  GOOGLE_METRIC_SUBSTITUTES,
  googleFonts,
} from '@docx-editor.dev/fonts/google';
import { useFonts } from '@pptx-editor.dev/vue';

/** Stand-ins (Carlito for Calibri, …) are implementation faces, not choices. */
const standIns = new Set(Object.values(GOOGLE_METRIC_SUBSTITUTES).map((f) => f.toLowerCase()));

/** The catalog families the font picker offers. */
export const CATALOG_FAMILIES: readonly string[] = GOOGLE_FONT_FAMILIES.filter(
  (f) => !standIns.has(f.toLowerCase())
);

/** The demo's `fonts` prop: one stable value for the editor's life. Call it in `setup`. */
export function useDemoFonts(): ReturnType<typeof useFonts> {
  return useFonts(packagedFonts(), googleFonts());
}
