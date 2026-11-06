/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/server.ts` (commit 5993bd223945).
//
// A runtime over PPTX bytes. Opening is the editor's own open: the bounded OPC/XML reader, one
// store, and the HarfBuzz measurer the editor lays text out with, so a text write re-fits its
// text body by its autofit rules as an edit in the editor does.

import { createStoreAutomationHost } from '@pptx-editor.dev/core/automation';
import {
  openPresentation,
  renderImageWith,
  type ImageRasterizer,
} from '@pptx-editor.dev/core/editor';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { fail, failArgument } from './errors.ts';
import { createRuntime, type PptxEditorServerRuntime } from './runtime.ts';

/**
 * How `PptxEditor.createServer` opens a presentation.
 *
 * Opening PPTX bytes is a bounded parse: zip entry, size, and ratio caps, part name checks, and
 * XML with no DTD. Malformed input refuses with `InvalidArgument`; a zip over a cap refuses with
 * `ResourceLimitExceeded`. No message carries file content.
 *
 * @public
 */
export interface CreateServerOptions {
  /**
   * Faces to measure text with, tried before the presentation's embedded fonts and the
   * metric-compatible substitutes.
   */
  readonly fonts?: FontSource | readonly FontSource[];
  /** Load the metric-compatible substitute fonts. Default `true`. */
  readonly substitutes?: boolean;
  /**
   * Turns the SVG of a slide or shape picture into PNG bytes, so `getImageAsBase64` works on a
   * server. The engine paints the picture with the editor's own renderer into a standalone SVG
   * document, with its pictures and fonts inline, and the rasterizer only draws that SVG, for
   * example with an SVG rasterizer library or a headless browser. It must answer a PNG of
   * exactly `width` by `height` pixels; any other answer refuses with `GeneralException`.
   * Without it, `capabilities.images` is false and `getImageAsBase64` refuses with
   * `NotSupported`.
   */
  readonly rasterize?: ImageRasterizer;
}

export async function createServer(
  bytes: Uint8Array,
  options: CreateServerOptions = {}
): Promise<PptxEditorServerRuntime> {
  if (!(bytes instanceof Uint8Array) || bytes.length === 0)
    failArgument('createServer', 'the bytes of a PPTX file as a non-empty Uint8Array');
  if (options === null || typeof options !== 'object')
    failArgument('createServer', 'an options object');
  const rasterize = options.rasterize;
  if (rasterize !== undefined && typeof rasterize !== 'function')
    failArgument('createServer', 'a rasterize function');
  const fonts =
    options.fonts === undefined
      ? undefined
      : Array.isArray(options.fonts)
        ? (options.fonts as readonly FontSource[])
        : [options.fonts as FontSource];
  // The runtime keeps no reference to the caller's array.
  const opened = await openPresentation(bytes.slice(), {
    ...(fonts ? { fonts } : {}),
    ...(options.substitutes === false ? { substitutes: false } : {}),
  });
  if (!opened.ok) {
    const error = opened.error;
    const limit =
      (error.reason === 'zip' &&
        (error.detail === 'too-many-entries' || error.detail === 'too-large')) ||
      error.reason === 'too-many-parts';
    fail({ code: limit ? 'ResourceLimitExceeded' : 'InvalidArgument', target: 'createServer' });
  }
  const session = opened.session;
  const host = createStoreAutomationHost(session.store, {
    measurer: session.measurer,
    ...(rasterize
      ? {
          renderImage: (target) =>
            renderImageWith(
              session,
              target.slideIndex,
              {
                ...(target.shapeNodeId === undefined ? {} : { shapeNodeId: target.shapeNodeId }),
                ...(target.width === undefined ? {} : { width: target.width }),
                ...(target.height === undefined ? {} : { height: target.height }),
              },
              rasterize
            ),
        }
      : {}),
  });
  return createRuntime({
    host: {
      ...host,
      dispose() {
        host.dispose();
        session.dispose();
      },
    },
    save: true,
  });
}
