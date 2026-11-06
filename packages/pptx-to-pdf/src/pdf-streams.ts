/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Copied from docx-editor/packages/docx-to-pdf/src/context.ts (EigenPal): the stream and hex
// helpers the font writer uses.
import { deflateSync } from 'node:zlib';
import type { PDFContext, PDFRawStream } from 'pdf-lib';

type StreamDict = NonNullable<Parameters<PDFContext['stream']>[1]>;

/**
 * A `FlateDecode` stream compressed by the runtime's native zlib. pdf-lib's `flateStream` runs
 * the same algorithm in JavaScript. The decoded content is the same; the compressed bytes can
 * differ between runtimes. A string is taken byte per UTF-16 unit, as pdf-lib takes it.
 */
export function flateStream(
  context: PDFContext,
  contents: string | Uint8Array,
  dict: StreamDict = {}
): PDFRawStream {
  const bytes = typeof contents === 'string' ? Buffer.from(contents, 'latin1') : contents;
  const deflated = deflateSync(bytes);
  // A small result is a view of zlib's 16 KiB output chunk. pdf-lib keeps every stream until
  // the save, so keep only the bytes rather than one chunk per page.
  const owned =
    deflated.buffer.byteLength > deflated.byteLength * 2 ? new Uint8Array(deflated) : deflated;
  return context.stream(owned, { ...dict, Filter: 'FlateDecode' });
}

const HEX_BYTES = Array.from({ length: 256 }, (_, byte) => byte.toString(16).padStart(2, '0'));

export function hex(value: number): string {
  // A UTF-16 code unit or a character code: always four digits, no wider.
  if (value >= 0 && value <= 0xffff && Number.isInteger(value))
    return HEX_BYTES[value >> 8]! + HEX_BYTES[value & 0xff]!;
  return value.toString(16).padStart(4, '0');
}

export function unicodeHex(text: string): string {
  const out: string[] = [];
  for (let i = 0; i < text.length; i++) out.push(hex(text.charCodeAt(i)));
  return out.join('');
}
