/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The deck's core properties (`docProps/core.xml`, ECMA-376 Part 2, 11) for the PDF information
// dictionary, as DOCX Editor's `docx-to-pdf` writes title, author, subject, and keywords.
// Every value comes from the file, so each is read through the bounded parser, stripped of
// control characters, and capped in length.

import {
  binaryPartOf,
  elements,
  parseXmlPart,
  relTargetByType,
  textOf,
  xmlPartOf,
  type PptxPackage,
  type XmlPart,
} from '@pptx-editor.dev/core/store';

const CORE_PROPERTIES =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const MAX_LENGTH = 1024;

/** The PDF document information this package writes. */
export interface DeckMetadata {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
}

function corePart(pkg: PptxPackage): XmlPart | undefined {
  const name = relTargetByType(pkg, '/', CORE_PROPERTIES) ?? '/docProps/core.xml';
  const modeled = xmlPartOf(pkg, name);
  if (modeled) return modeled;
  const bytes = binaryPartOf(pkg, name);
  if (!bytes) return undefined;
  const parsed = parseXmlPart(name, new TextDecoder().decode(bytes));
  return parsed.ok ? parsed.part : undefined;
}

function clean(text: string): string | undefined {
  let out = '';
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    out += code < 0x20 || code === 0x7f ? ' ' : ch;
    if (out.length >= MAX_LENGTH) break;
  }
  const trimmed = out.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/** Title, author, subject, and keywords from the deck's core properties. */
export function deckMetadata(pkg: PptxPackage): DeckMetadata {
  const root = corePart(pkg)?.root;
  const read = (local: string): string | undefined => {
    for (const el of elements(root)) if (el.local === local) return clean(textOf(el));
    return undefined;
  };
  const title = read('title');
  const author = read('creator');
  const subject = read('subject');
  const keywords = read('keywords');
  return {
    ...(title ? { title } : {}),
    ...(author ? { author } : {}),
    ...(subject ? { subject } : {}),
    ...(keywords ? { keywords } : {}),
  };
}
