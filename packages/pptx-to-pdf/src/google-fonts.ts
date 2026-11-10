/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { googleFonts } from '@docx-editor.dev/fonts/google';
import { deckFontRequest, embeddedFontSources } from '@pptx-editor.dev/core/editor';
import { readPptx } from '@pptx-editor.dev/core/store';
import {
  composeFontOrigins,
  sha256FontBytes,
  boundedStructuralFontValidator,
  type FontSource,
  type FontOrigin,
} from '@pptx-editor.dev/core/layout';

/** Explicit network font source. Catalog entries are pinned and verified before shaping. */
export async function googleExportFonts(
  source: Uint8Array,
  signal?: AbortSignal,
  provided: readonly FontSource[] = [],
  origin: FontOrigin = googleFonts({ cjk: { loading: 'wait' } })
): Promise<FontSource[]> {
  const read = readPptx(source);
  if (!read.ok || signal?.aborted) return [];
  const config = await composeFontOrigins([origin], deckFontRequest(read.pkg, signal));
  const existing = [...provided, ...embeddedFontSources(read.pkg).sources];
  const fonts: FontSource[] = [];
  for (const face of config?.sources ?? []) {
    if (face.availability === 'forbidden') continue;
    if (
      existing.some(
        (f) =>
          f.family.toLowerCase() === face.request.family.toLowerCase() &&
          f.bold === face.request.weight >= 600 &&
          f.italic === (face.request.style === 'italic')
      )
    )
      continue;
    if (
      face.faceIndex !== 0 ||
      sha256FontBytes(face.bytes) !== face.hash ||
      !boundedStructuralFontValidator(face.bytes, face.faceIndex).valid
    )
      throw new Error('Google font failed admission');
    fonts.push({
      family: face.request.family,
      bold: face.request.weight >= 600,
      italic: face.request.style === 'italic',
      bytes: face.bytes,
      origin: 'google',
    });
  }
  return fonts;
}
