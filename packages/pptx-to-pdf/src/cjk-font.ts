/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The East Asian fallback face of PDF export: Noto Sans CJK JP from the optional
// `@docx-editor.dev/fonts-cjk`. `isPackageNotFound` and `importCjkFace` are copied from
// docx-editor/packages/docx-to-pdf/src/font-provisioning.ts (EigenPal, commit 12a57e37c).
// When the package is installed and a deck holds East Asian text, the export measures and
// draws that text with the face wherever the deck's own East Asian font is missing (the
// measurer's `EAST_ASIAN_FALLBACK_FAMILIES`), so no character is left out.

import { readFile } from 'node:fs/promises';
import { URL as NodeURL } from 'node:url';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import { readPptx, type Element, type PptxPackage } from '@pptx-editor.dev/core/store';

/** The optional package that carries the CJK face. */
const CJK_PACKAGE = '@docx-editor.dev/fonts-cjk';
/** The family of the packaged face. */
export const CJK_FAMILY = 'Noto Sans CJK JP';

/** Where the CJK face file is, or `null` when `@docx-editor.dev/fonts-cjk` is not installed. */
export type CjkFaceLocator = () => Promise<string | NodeURL | null>;

/**
 * Whether `error` says that `specifier` itself is not installed, as Node's ESM loader, its
 * CommonJS loader and Bun each phrase it. A package that is installed but broken (a missing
 * `dist/` file, a missing dependency of its own) names another module, so it does not match:
 * that failure propagates instead of reading as "not installed".
 */
export function isPackageNotFound(error: unknown, specifier: string): boolean {
  if (typeof error !== 'object' || error === null) return false;
  const { code, message } = error as { code?: unknown; message?: unknown };
  if (code !== 'ERR_MODULE_NOT_FOUND' && code !== 'MODULE_NOT_FOUND') return false;
  if (typeof message !== 'string') return false;
  return [`Cannot find package '${specifier}'`, `Cannot find module '${specifier}'`].some(
    (prefix) => message.startsWith(prefix)
  );
}

/** Locate the CJK face in `@docx-editor.dev/fonts-cjk`, or `null` when it is not installed. */
export async function importCjkFace(): Promise<NodeURL | null> {
  let module: { NOTO_SANS_CJK_JP_URL?: unknown };
  try {
    // A literal specifier, so a file tracer that follows it ships the package with a
    // deployment. `webpackIgnore` leaves the import to Node at run time in the bundlers that
    // read it, so such a bundle builds without the optional package.
    module = await import(/* webpackIgnore: true */ '@docx-editor.dev/fonts-cjk');
  } catch (error) {
    if (isPackageNotFound(error, CJK_PACKAGE)) return null;
    throw error;
  }
  const file = module.NOTO_SANS_CJK_JP_URL as { href?: unknown } | undefined;
  if (typeof file?.href !== 'string')
    throw new TypeError(`${CJK_PACKAGE} does not export NOTO_SANS_CJK_JP_URL`);
  return new NodeURL(file.href);
}

/** Characters only an East Asian face carries: Hangul, kana, ideographs, full-width forms. */
const CJK_TEXT = /[ᄀ-ᇿ⺀-⿟　-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-￯]|[\u{20000}-\u{3ffff}]/u;

/** The parts whose text a slide shows: slides, layouts, and masters. */
const TEXT_PARTS = /^\/ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/;

/** Whether a deck's slides, layouts, or masters hold East Asian text (`a:t`). */
export function hasEastAsianText(pkg: PptxPackage): boolean {
  const walk = (el: Element, depth: number): boolean => {
    if (depth > 256) return false;
    for (const c of el.children) {
      if (c.type === 'text') {
        if (el.local === 't' && CJK_TEXT.test(c.value)) return true;
      } else if (c.type === 'element' && walk(c, depth + 1)) return true;
    }
    return false;
  };
  for (const part of pkg.parts.values())
    if (part.kind === 'xml' && TEXT_PARTS.test(part.partName) && walk(part.xml.root, 0))
      return true;
  return false;
}

/** The face bytes, read once per process for each file. */
const faces = new Map<string, Promise<Uint8Array>>();

/**
 * The CJK fallback face for a deck: a font source when `@docx-editor.dev/fonts-cjk` is installed
 * and the deck holds East Asian text, else undefined. A deck that does not open returns
 * undefined; the open reports why.
 */
export async function cjkFallbackFont(
  source: Uint8Array,
  locate: CjkFaceLocator = importCjkFace
): Promise<FontSource | undefined> {
  const file = await locate();
  if (file === null) return undefined;
  const read = readPptx(source);
  if (!read.ok || !hasEastAsianText(read.pkg)) return undefined;
  const key = String(file);
  let bytes = faces.get(key);
  if (!bytes) {
    bytes = readFile(file).then((b) => new Uint8Array(b.buffer, b.byteOffset, b.byteLength));
    // A failed read is not kept, so a later export can try again.
    bytes.catch(() => faces.delete(key));
    faces.set(key, bytes);
  }
  return { family: CJK_FAMILY, bold: false, italic: false, bytes: await bytes, origin: 'cjk' };
}
