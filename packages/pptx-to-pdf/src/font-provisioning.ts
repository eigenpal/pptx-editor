/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { admitFontConfiguration, deckFontRequest } from '@pptx-editor.dev/core/editor';
import { readPptx } from '@pptx-editor.dev/core/store';
import { composeFontOrigins, type FontOrigin } from '@pptx-editor.dev/core/layout';
import type { PdfFontsSource } from './types.ts';

/** Compose and admit all font layers before the immutable export layout opens. */
export async function exportFonts(
  source: Uint8Array,
  origins: PdfFontsSource | undefined,
  signal?: AbortSignal
) {
  const read = readPptx(source);
  if (!read.ok || signal?.aborted) return admitFontConfiguration([]).faces;
  const list: readonly FontOrigin[] = Array.isArray(origins) ? origins : [origins as FontOrigin];
  const failures: unknown[] = [];
  let config = await composeFontOrigins(list, deckFontRequest(read.pkg, signal), {
    onOriginFailure: ({ cause }) => {
      failures.push(cause);
    },
  });
  // Composition merges earlier faces into every deferred result and bounds the chain.
  while (config?.deferred) {
    const { deferred, ...now } = config;
    config = (await deferred) ?? now;
  }
  if (signal?.aborted) throw signal.reason;
  if (failures.length) throw new AggregateError(failures, 'Font resolution failed');
  const admitted = admitFontConfiguration(config ? [config] : []);
  if (admitted.refused.length) throw new AggregateError(admitted.refused, 'Font admission failed');
  return admitted.faces;
}
