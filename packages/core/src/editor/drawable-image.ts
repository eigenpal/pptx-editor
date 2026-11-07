// Which picture bytes the browser gets to draw. A picture part is untrusted input: its name can
// lie about its format, its bytes can be empty or damaged, and its header can claim a size
// that would exhaust memory. A picture that fails these checks paints the placeholder.

import { DEFAULT_IMAGE_RESOURCE_LIMITS } from '@eigenpal/ooxml-core';
import { imageSize } from '../store/index.ts';

const LIMITS = DEFAULT_IMAGE_RESOURCE_LIMITS;

/** Whether `bytes` start with `text` in ASCII at `at`. */
const startsWith = (bytes: Uint8Array, text: string, at = 0): boolean =>
  bytes.length >= at + text.length && [...text].every((c, i) => bytes[at + i] === c.charCodeAt(0));

/** Whether the first non-space character of `bytes` is `<`, as an SVG document starts. */
function startsAsMarkup(bytes: Uint8Array): boolean {
  let i = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (i < bytes.length && i < 1024 && [0x20, 0x09, 0x0a, 0x0d].includes(bytes[i]!)) i += 1;
  return bytes[i] === 0x3c;
}

/**
 * The MIME type to draw `bytes` with, from their signature, or undefined when they are not a
 * picture the browser can draw within the limits. `named` is the type the part name gives;
 * only SVG, which has no binary signature, depends on it.
 */
export function drawableImageMime(
  bytes: Uint8Array,
  named: string | undefined
): string | undefined {
  if (bytes.length === 0 || bytes.length > LIMITS.maxEncodedBytes) return undefined;
  if (named === 'image/svg+xml') return startsAsMarkup(bytes) ? named : undefined;
  const size = imageSize(bytes);
  if (!size || size.width <= 0 || size.height <= 0) return undefined;
  if (size.width > LIMITS.maxDimension || size.height > LIMITS.maxDimension) return undefined;
  if (size.width * size.height > LIMITS.maxPixels) return undefined;
  const b = bytes;
  if (startsWith(b, 'RIFF')) return 'image/webp';
  if (b[0] === 0x89) return 'image/png';
  if (b[0] === 0xff) return 'image/jpeg';
  if (b[0] === 0x47) return 'image/gif';
  return 'image/bmp';
}
