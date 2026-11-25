// Replace All for the object model (`Presentation.replaceAll`, an SDK extension).
//
// The search and the replacement are the editor's own (`commands/find.ts`): every shape on each
// slide (also in groups), every drawn table cell, and the speaker notes, in slide order. The
// replacement goes into the `a:t` of the `a:r` (`CT_RegularTextRun`, ECMA-376 Part 1, 21.1.2.3.8
// and 21.1.2.3.11) that holds the first matched character, so it keeps that run's `a:rPr`
// (`CT_TextCharacterProperties`, Part 1, 21.1.2.3.9). A text field (`a:fld`, `CT_TextField`,
// Part 1, 21.1.2.2.4) holds generated text: it stays atomic, and a match never covers it.

import type { FindOptions } from '../contracts/neutral.ts';
import { isValidXmlText } from '../store/index.ts';
import { refit } from './commands/autofit.ts';
import { findMatches, MAX_FIND_TEXT, replaceMatches } from './commands/find.ts';
import type { AutomationValue } from './protocol.ts';
import type { TextPlanContext } from './plan-text.ts';
import { badArgument, edit } from './targets.ts';

/** Find and replace text across the presentation. Answers the number of replaced matches. */
export function replaceText(
  ctx: TextPlanContext,
  query: string,
  replacement: string,
  options: FindOptions | undefined
): AutomationValue {
  // A match stays inside one paragraph, so a query with a paragraph break never matches. The
  // object model writes a line break as `\v`; the find text model reads it as `\n`.
  if (typeof query !== 'string' || !query || query.length > MAX_FIND_TEXT || /[\r\n]/.test(query))
    badArgument('query');
  if (!isValidXmlText(query.replace(/\v/g, ' '))) badArgument('query');
  if (typeof replacement !== 'string' || replacement.length > MAX_FIND_TEXT)
    badArgument('replacement');
  if (!isValidXmlText(replacement.replace(/\v/g, ' '))) badArgument('replacement');
  if (options !== undefined && (typeof options !== 'object' || options === null))
    badArgument('find-options');
  for (const name of ['matchCase', 'wholeWord'] as const)
    if (options?.[name] !== undefined && typeof options[name] !== 'boolean')
      badArgument('find-options');
  const find: FindOptions = { matchCase: options?.matchCase, wholeWord: options?.wholeWord };
  const matches = findMatches(ctx.batch.pkg, query.replace(/\v/g, '\n'), find);
  if (!matches.length) return { kind: 'number', value: 0 };
  edit(() => replaceMatches(ctx.batch, matches, replacement));
  // Autofit text re-fits and a table frame takes its laid-out size, as after the editor's
  // Replace All (`editor/commands.ts`). Notes have no autofit.
  const measurer = ctx.measurer;
  if (measurer) {
    const shapes = new Set<string>();
    for (const m of matches) {
      const key = `${m.part}\u0000${m.nodeId}`;
      if (m.where !== 'slide' || shapes.has(key)) continue;
      shapes.add(key);
      const kind = ctx.batch.element(m.part, m.nodeId).kind;
      if (kind === 'sp' || kind === 'graphicFrame')
        edit(() => refit(ctx.batch, m.slide, m.part, m.nodeId, measurer));
    }
  }
  return { kind: 'number', value: matches.length };
}
