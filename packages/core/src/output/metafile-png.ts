// EMF and WMF pictures as PNG. An image element cannot draw a metafile, so the editor, the image
// export, and the PDF writer convert one with `emf-converter`, which replays the records onto
// the host's canvas: the browser canvas in the editor and the fidelity harness; on Node.js,
// `@napi-rs/canvas` when it is installed, else the converter's own rasterizer. It needs no DOM.
// It is imported lazily: only a deck with a metafile loads it.
//
// The approach is from DOCX Editor `packages/core/src/editor/browser-image-decode-port.ts`
// (`convertPlainMetafile`) at commit bac199e196492ecc3fb5fea249942ffd3657ea66.

import { DEFAULT_IMAGE_RESOURCE_LIMITS, imageSize, metafileFormat } from '@eigenpal/ooxml-core';

/** The longest side of a converted metafile, in pixels: well under every image limit. @internal */
export const MAX_METAFILE_RASTER_EDGE = 4096;
/** How long one conversion may take before the picture stays a placeholder, in ms. */
const MAX_METAFILE_MS = 10_000;
/** Pixels per metafile pixel: sharp on high-density screens and when a slide zooms in. */
const DPI_SCALE = 2;
const PNG_PREFIX = 'data:image/png;base64,';

/**
 * The PNG of an EMF or WMF picture, at most `MAX_METAFILE_RASTER_EDGE` pixels on a side. Null
 * when `bytes` are not a metafile or do not convert within the limits; undefined when the time
 * limit or `signal` stopped the conversion, so a later one can succeed. Text in the
 * metafile draws with `fonts` (TrueType or OpenType files): pass the faces that the deck's
 * text uses, so every host draws the same glyphs. Pass the same array again to reuse the fonts
 * the converter has read. `signal` stops the conversion. Never throws. @internal
 */
export async function metafilePng(
  bytes: Uint8Array,
  fonts: readonly Uint8Array[] = [],
  signal?: AbortSignal
): Promise<Uint8Array | null | undefined> {
  if (!metafileFormat(bytes) || bytes.length > DEFAULT_IMAGE_RESOURCE_LIMITS.maxEncodedBytes)
    return null;
  try {
    const { convertMetafileToDataUrl } = await import('emf-converter');
    const url = await convertMetafileToDataUrl(bytes.slice(), {
      maxWidth: MAX_METAFILE_RASTER_EDGE,
      maxHeight: MAX_METAFILE_RASTER_EDGE,
      maxCanvasDimension: MAX_METAFILE_RASTER_EDGE,
      dpiScale: DPI_SCALE,
      timeoutMs: MAX_METAFILE_MS,
      // Gray smoothing: no colored fringes that depend on a screen's subpixel order.
      ...(fonts.length ? { fonts: fonts as Uint8Array[], fontSmoothing: 'gray' as const } : {}),
      ...(signal ? { signal } : {}),
    });
    if (!url?.startsWith(PNG_PREFIX)) return null;
    const binary = atob(url.slice(PNG_PREFIX.length));
    const png = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) png[i] = binary.charCodeAt(i);
    // The output is checked like any picture before it draws.
    const size = imageSize(png);
    const fits = (n: number): boolean => n > 0 && n <= MAX_METAFILE_RASTER_EDGE;
    return size && png[0] === 0x89 && fits(size.width) && fits(size.height) ? png : null;
  } catch (error) {
    // Stopped: past its time (`TimeoutError`) or aborted. Anything else: a damaged metafile, or a
    // host where the converter fails.
    const name = (error as { name?: unknown } | null)?.name;
    return name === 'TimeoutError' || name === 'AbortError' || signal?.aborted ? undefined : null;
  }
}
