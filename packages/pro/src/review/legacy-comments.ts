/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Legacy PresentationML comments (ECMA-376 Part 1, 19.4): `ppt/commentAuthors.xml` holds
// `p:cmAuthor id name`, and each slide links to a `p:cmLst` part of `p:cm authorId dt idx`
// with `p:text`. A reply names its parent in `p15:threadingInfo/p15:parentCm` (a `p15`
// extension). Modern threaded comments (`p188:`) are not read yet.

import {
  attr,
  binaryPartOf,
  elements,
  parseXmlPart,
  mainPresentationPart,
  presentationInfo,
  relTargetByType,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
  type XmlPart,
} from '@pptx-editor.dev/core/store';
import type { ReviewItem } from '@pptx-editor.dev/core/editor';

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const COMMENTS = `${OD}/comments`;
const COMMENT_AUTHORS = `${OD}/commentAuthors`;

/**
 * The tree of a comment part. The store keeps XML parts it does not model as bytes, so they
 * round-trip unchanged; read them here, read-only, through the same bounded parser.
 */
function commentPart(pkg: PptxPackage, partName: string | undefined): XmlPart | undefined {
  const modeled = xmlPartOf(pkg, partName);
  if (modeled || !partName) return modeled;
  const bytes = binaryPartOf(pkg, partName);
  if (!bytes) return undefined;
  const parsed = parseXmlPart(partName, new TextDecoder().decode(bytes));
  return parsed.ok ? parsed.part : undefined;
}

/** Author names by `p:cmAuthor/@id`, in a null-prototype record (ids come from the file). */
function authorsOf(pkg: PptxPackage): Record<string, string> {
  const out: Record<string, string> = Object.create(null);
  const presentation = mainPresentationPart(pkg);
  if (!presentation) return out;
  const part = commentPart(pkg, relTargetByType(pkg, presentation, COMMENT_AUTHORS));
  for (const author of elements(part?.root)) {
    if (author.local !== 'cmAuthor') continue;
    const id = attr(author, 'id');
    const name = attr(author, 'name');
    if (id !== undefined && name !== undefined) out[id] = name;
  }
  return out;
}

/** The first descendant with this local name, depth-first. */
function descendant(el: Element | undefined, local: string): Element | undefined {
  for (const c of elements(el)) {
    if (c.local === local) return c;
    const hit = descendant(c, local);
    if (hit) return hit;
  }
  return undefined;
}

const commentId = (slidePart: string, authorId: string, idx: string): string =>
  `${slidePart}#cm-${authorId}-${idx}`;

/** Every legacy comment of every slide, in slide order and then file order. */
export function legacyComments(pkg: PptxPackage): ReviewItem[] {
  const authors = authorsOf(pkg);
  const out: ReviewItem[] = [];
  for (const slide of presentationInfo(pkg).slides) {
    const part = commentPart(pkg, relTargetByType(pkg, slide.partName, COMMENTS));
    for (const cm of elements(part?.root)) {
      if (cm.local !== 'cm') continue;
      const authorId = attr(cm, 'authorId') ?? '';
      const idx = attr(cm, 'idx') ?? '';
      const parent = descendant(cm, 'parentCm');
      const author = authors[authorId];
      const date = attr(cm, 'dt');
      out.push({
        id: commentId(slide.partName, authorId, idx),
        kind: 'comment',
        slidePart: slide.partName,
        text: textOf(descendant(cm, 'text')),
        ...(author !== undefined ? { author } : {}),
        ...(date !== undefined ? { date } : {}),
        ...(parent
          ? {
              parentId: commentId(
                slide.partName,
                attr(parent, 'authorId') ?? '',
                attr(parent, 'idx') ?? ''
              ),
            }
          : {}),
      });
    }
  }
  return out;
}
