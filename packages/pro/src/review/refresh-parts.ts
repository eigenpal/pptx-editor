/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The related parts of the computed refresh changes (refresh-changes.ts): the parts that a
// slide, a layout, or a chart reaches through its relationships, compared without their part
// names, and the relationship ids that an element refers to. DOM-free.

import {
  elements,
  relationshipById,
  relationshipsOf,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { partContent, sameList } from './refresh-common.ts';

/** How deep a relationship chain goes: a slide, its chart, the chart's workbook. */
const MAX_REL_DEPTH = 3;

/** Relationship types that another comparison reports, or that the slide XML never names. */
const OWN_REPORT =
  /\/(slide|slideLayout|slideMaster|theme|notesSlide|notesMaster|handoutMaster|comments|commentAuthors)$/;

/** Whether a relationship type has a report of its own (`OWN_REPORT`). */
export const ownReport = (type: string): boolean => OWN_REPORT.test(type);

/**
 * What one relationship of `source` reaches: its type, an external target, or the content of
 * the target part and of the parts that part reaches in turn. Part names do not count, so a
 * writer that names its media again reports no change. Each part reached goes into `reached`.
 */
export function relatedContent(
  pkg: PptxPackage,
  source: string,
  id: string,
  reached: Set<string>
): (string | Uint8Array | undefined)[] {
  const out: (string | Uint8Array | undefined)[] = [];
  const seen = new Set<string>();
  const visit = (src: string, rid: string, depth: number) => {
    const rel = relationshipById(pkg, src, rid);
    if (!rel) {
      out.push(`${rid}|none`);
      return;
    }
    if (rel.external) {
      out.push(`${rid}|${rel.type}|external|${rel.externalTarget ?? rel.relativeTarget ?? ''}`);
      return;
    }
    const target = rel.target!;
    if (depth > 0 && ownReport(rel.type)) {
      // A chart or a picture that names a slide or a layout: the name is enough.
      out.push(`${rid}|${rel.type}|${target}`);
      return;
    }
    out.push(`${rid}|${rel.type}`);
    if (seen.has(target)) return;
    seen.add(target);
    reached.add(target);
    out.push(partContent(pkg, target));
    if (depth >= MAX_REL_DEPTH) return;
    for (const r of relationshipsOf(pkg, target)) visit(target, r.id, depth + 1);
  };
  visit(source, id, 0);
  return out;
}

/** The relationship ids that `el` and its descendants name, in document order. */
export function relIdsOf(el: Element, out = new Set<string>(), depth = 0): Set<string> {
  if (depth > 64) return out;
  for (const a of el.attrs) if (a.ns.endsWith('/relationships')) out.add(a.value);
  for (const c of elements(el)) relIdsOf(c, out, depth + 1);
  return out;
}

/** The relationships of `source` that the slide comparison owns: all but `OWN_REPORT` ones. */
export function contentRelIds(pkg: PptxPackage, source: string | undefined): string[] {
  return source
    ? relationshipsOf(pkg, source)
        .filter((r) => r.external || !ownReport(r.type))
        .map((r) => r.id)
    : [];
}

/**
 * The relationship ids of a part whose reach differs between two packages: an id on one side
 * only, or a different target, type, or content. Each part reached goes into `reached`.
 */
export function changedRelIds(
  a: PptxPackage,
  b: PptxPackage,
  partA: string,
  partB: string,
  reached: Set<string>
): Set<string> {
  const out = new Set<string>();
  const idsA = contentRelIds(a, partA);
  const idsB = contentRelIds(b, partB);
  for (const id of new Set([...idsA, ...idsB])) {
    const ca = idsA.includes(id) ? relatedContent(a, partA, id, reached) : [];
    const cb = idsB.includes(id) ? relatedContent(b, partB, id, reached) : [];
    if (!sameList(ca, cb)) out.add(id);
  }
  return out;
}

/** The parts that `source` reaches through its own relationships, into `reached`. */
export function markReached(pkg: PptxPackage, source: string, reached: Set<string>): void {
  for (const id of contentRelIds(pkg, source)) relatedContent(pkg, source, id, reached);
}

/**
 * The content of a layout, a slide master, or a theme with the parts it reaches, for
 * equality. The parts that `OWN_REPORT` names count by name only.
 */
export function templateContent(
  pkg: PptxPackage,
  part: string
): (string | Uint8Array | undefined)[] {
  const out: (string | Uint8Array | undefined)[] = [partContent(pkg, part)];
  const reached = new Set<string>();
  for (const r of relationshipsOf(pkg, part)) {
    if (!r.external && ownReport(r.type)) out.push(`${r.id}|${r.type}|${r.target}`);
    else out.push(...relatedContent(pkg, part, r.id, reached));
  }
  return out;
}
