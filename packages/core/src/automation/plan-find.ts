// Find and Replace All for the object model (`TextRange.search` and `Presentation.replaceAll`,
// SDK extensions).
//
// The search and the replacement are the editor's own (`commands/find.ts`): every shape on each
// slide (also in groups), every drawn table cell, and the speaker notes, in slide order. The
// replacement goes into the `a:t` of the `a:r` (`CT_RegularTextRun`, ECMA-376 Part 1, 21.1.2.3.8
// and 21.1.2.3.11) that holds the first matched character, so it keeps that run's `a:rPr`
// (`CT_TextCharacterProperties`, Part 1, 21.1.2.3.9). A text field (`a:fld`, `CT_TextField`,
// Part 1, 21.1.2.2.4) holds generated text: it stays atomic, and a match never covers it.

import type { FindOptions } from '../contracts/neutral.ts';
import { isValidXmlText } from '../store/index.ts';
import { refitAll } from './commands/autofit.ts';
import { findMatches, matchesInBody, MAX_FIND_TEXT, replaceMatches } from './commands/find.ts';
import type { AutomationTextSpan } from './operations.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import { bodyTextOf, readTarget, type TextPlanContext } from './plan-text.ts';
import { badArgument, edit, numberIn, slideOf, targetOf } from './targets.ts';

/**
 * The search text, or `invalid-argument`. A match stays inside one paragraph, so a query with a
 * paragraph break never matches. The object model writes a line break as `\v`; the find text
 * model reads it as `\n`. U+FFFC stands in for a field in the search text: a query with it
 * could match one.
 */
function findQuery(query: string): string {
  if (
    typeof query !== 'string' ||
    !query ||
    query.length > MAX_FIND_TEXT ||
    /[\r\n\uFFFC]/.test(query)
  )
    badArgument('query');
  if (!isValidXmlText(query.replace(/\v/g, ' '))) badArgument('query');
  return query.replace(/\v/g, '\n');
}

/**
 * The find options, or `invalid-argument`. An option name the search does not know is refused,
 * not dropped: a search that ignored it would match text the caller did not ask for.
 */
function findOptions(options: FindOptions | undefined): FindOptions {
  if (options !== undefined && (typeof options !== 'object' || options === null))
    badArgument('find-options');
  for (const [name, value] of Object.entries(options ?? {}))
    if (
      (name !== 'matchCase' && name !== 'wholeWord') ||
      (value !== undefined && typeof value !== 'boolean')
    )
      badArgument('find-options');
  return { matchCase: options?.matchCase, wholeWord: options?.wholeWord };
}

/**
 * The matches of `query` in a shape's text, or in a span of it, in the offsets of the whole
 * text. A match that is not wholly inside the span is left out. A shape that cannot hold text
 * (a picture, a group, a table frame) refuses, as `readText` does.
 */
export function findText(
  ctx: TextPlanContext,
  ref: AutomationRef,
  query: string,
  options: FindOptions | undefined,
  span?: AutomationTextSpan
): AutomationValue {
  const needle = findQuery(query);
  const find = findOptions(options);
  // Not nullable: a shape that cannot hold text refuses, and never answers null.
  const { t } = readTarget(ctx, ref, undefined, false)!;
  let from = 0;
  let to = Number.POSITIVE_INFINITY;
  if (span !== undefined) {
    const total = bodyTextOf(t.txBody, t.slideNumber).length;
    const start = numberIn(span.start, 'start', 0, total);
    const length = numberIn(span.length, 'length', 0, total - start);
    if (!Number.isInteger(start) || !Number.isInteger(length)) badArgument('span');
    from = start;
    to = start + length;
  }
  const spans = matchesInBody(t.txBody, t.slideNumber, needle, find)
    .filter(([start, end]) => start >= from && end <= to)
    .map(([start, end]) => ({ start, length: end - start }));
  return { kind: 'textSpans', spans };
}

/**
 * Find and replace text across the presentation, or on one slide (its shapes, table cells, and
 * notes). Answers the number of replaced matches.
 */
export function replaceText(
  ctx: TextPlanContext,
  query: string,
  replacement: string,
  options: FindOptions | undefined,
  slide?: AutomationRef
): AutomationValue {
  const needle = findQuery(query);
  if (typeof replacement !== 'string' || replacement.length > MAX_FIND_TEXT)
    badArgument('replacement');
  if (!isValidXmlText(replacement.replace(/\v/g, ' '))) badArgument('replacement');
  const find = findOptions(options);
  const only =
    slide === undefined
      ? undefined
      : slideOf(ctx.batch.pkg, targetOf(ctx, slide, 'slide').sldId).index;
  const matches = findMatches(ctx.batch.pkg, needle, find).filter(
    (m) => only === undefined || m.slide === only
  );
  if (!matches.length) return { kind: 'number', value: 0 };
  edit(() => replaceMatches(ctx.batch, matches, replacement));
  // Autofit text re-fits and a table frame takes its laid-out size, as after the editor's
  // Replace All (`editor/commands.ts`), from one resolution of each slide. Notes have no autofit.
  const measurer = ctx.measurer;
  if (measurer) {
    const bySlide = new Map<number, { part: string; ids: Set<string> }>();
    for (const m of matches) {
      if (m.where !== 'slide') continue;
      const kind = ctx.batch.element(m.part, m.nodeId).kind;
      if (kind !== 'sp' && kind !== 'graphicFrame') continue;
      const entry = bySlide.get(m.slide) ?? { part: m.part, ids: new Set<string>() };
      entry.ids.add(m.nodeId);
      bySlide.set(m.slide, entry);
    }
    for (const [slide, { part, ids }] of bySlide)
      edit(() => refitAll(ctx.batch, slide, part, ids, measurer));
  }
  return { kind: 'number', value: matches.length };
}
