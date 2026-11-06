/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Slide.getImageAsBase64` and `Shape.getImageAsBase64`: a PNG picture as base64.
//
// The sync checks the target and the size in its batch (`prepareImage`), then asks the host to
// draw (`renderImage`) and waits for the picture before it returns. Only a host that draws
// pictures (a live editor in a browser, or a server runtime with a rasterizer) has the `images`
// capability; any other host refuses in the batch with `NotSupported`.

import type { AutomationHandle, AutomationValue } from '@pptx-editor.dev/core/automation';
import { hostFailure } from '../runtime/batch.ts';
import { diagnosticOf, failArgument, PptxEditorError } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ContextInternals } from '../runtime/internals.ts';
import { ShapeGetImageFormatType } from './enums.ts';

/**
 * Options for `Slide.getImageAsBase64`. The picture keeps the slide's aspect ratio and fits into
 * the box: with both sizes it fits into both, with one size it takes that size, and with neither
 * it has the slide's own size at 96 pixels per inch.
 *
 * @public
 */
export interface SlideGetImageOptions {
  /** The height of the box, in pixels. */
  height?: number;
  /** The width of the box, in pixels. */
  width?: number;
}

/**
 * Options for `Shape.getImageAsBase64`. The picture keeps the shape's aspect ratio and fits into
 * the box, as for a slide.
 *
 * @public
 */
export interface ShapeGetImageOptions {
  /** The image format. Only PNG. */
  format?: ShapeGetImageFormatType | `${ShapeGetImageFormatType}` | undefined;
  /** The height of the box, in pixels. */
  height?: number;
  /** The width of the box, in pixels. */
  width?: number;
}

/** The box a picture fits into, as the host takes it. */
export interface ImageBox {
  readonly width?: number;
  readonly height?: number;
}

/** A size of the box: absent, or a positive finite number. Else `InvalidArgument`. */
function side(value: unknown, target: string): number | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0)
    failArgument(target, 'a number of pixels greater than 0');
  return value;
}

/** The box of picture options, checked before the sync sends anything. `format` is a shape option only. */
export function imageBox(options: unknown, target: string, format: boolean): ImageBox {
  if (options === undefined || options === null) return {};
  if (typeof options !== 'object' || Array.isArray(options))
    failArgument(target, 'an options object with width and height');
  const o = options as Record<string, unknown>;
  if (
    format &&
    o['format'] !== undefined &&
    o['format'] !== null &&
    o['format'] !== ShapeGetImageFormatType.png
  )
    failArgument(target, "the format 'Png'");
  const width = side(o['width'], target);
  const height = side(o['height'], target);
  return {
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
  };
}

/** Ask the host to draw the picture of a checked handle, as base64. */
async function draw(
  internals: ContextInternals,
  handle: AutomationHandle,
  box: ImageBox,
  target: string
): Promise<string> {
  const host = internals.host;
  if (typeof host.renderImage !== 'function') {
    const diagnostic = diagnosticOf('images');
    throw new PptxEditorError({
      code: 'NotSupported',
      target,
      ...(diagnostic === undefined ? {} : { diagnostic }),
    });
  }
  const answer = await host.renderImage({ target: handle, ...box });
  if (!answer.ok) throw hostFailure(answer.error, { target });
  return answer.base64;
}

/**
 * The settle of a `prepareImage` read: start the drawing, and fill the result when it is done.
 * The sync waits for it.
 */
export function settleImage(
  internals: ContextInternals,
  box: ImageBox,
  target: string,
  fill: (base64: string) => void
): (value: AutomationValue) => void {
  return (value) => {
    const handle = hydrated(value, 'handle', target).handle;
    internals.defer(draw(internals, handle, box, target).then(fill));
  };
}
