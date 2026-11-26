/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The arguments of `TextRange.search`, `Presentation.replaceAll`, and `Slide.replaceAll`, SDK
// extensions over the editor's find (`core/src/automation/commands/find.ts`). Each argument is
// checked at the call, so a refusal names it; the engine checks the operation again
// (`plan-find.ts`).

import type { AutomationOperation } from '@pptx-editor.dev/core/automation';
import type { FindOptions } from '@pptx-editor.dev/core/contracts';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import { text } from './model-object.ts';

/**
 * How `search` and `replaceAll` compare text. Both options are off by default.
 *
 * @public
 */
export interface SearchOptions {
  /** Match upper and lower case exactly. */
  matchCase?: boolean;
  /**
   * Match whole words only: no letter, digit, combining mark, or underscore touches either end
   * of a match.
   */
  matchWholeWord?: boolean;
}

/** The longest search text and replacement text that the calls take (`MAX_FIND_TEXT`). */
const MAX_REPLACE_TEXT = 1000;

const SEARCH_OPTIONS: readonly string[] = ['matchCase', 'matchWholeWord'];

/**
 * Whether XML 1.0 can hold the text as a slide shows it: no control character but tab, line
 * feed, carriage return, and the line break U+000B, no U+FFFE or U+FFFF, no unpaired surrogate.
 */
function xmlText(value: string): boolean {
  return !/[\u0000-\u0008\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(
    value
  );
}

/** Search options as the engine takes them, or `InvalidArgument` naming `target`. */
function searchOptions(
  options: SearchOptions | undefined,
  target: string
): FindOptions | undefined {
  if (options === undefined) return undefined;
  if (typeof options !== 'object' || options === null)
    failArgument(target, 'a SearchOptions object');
  // A name this API does not have is refused, not dropped: a search that ignored it would
  // find or replace matches the caller did not ask for.
  for (const name of Object.keys(options))
    if (!SEARCH_OPTIONS.includes(name))
      failArgument(`${target}.${name}`, oneOf(SEARCH_OPTIONS), nearest(name, SEARCH_OPTIONS));
  const { matchCase, matchWholeWord } = options;
  for (const [name, value] of [
    ['matchCase', matchCase],
    ['matchWholeWord', matchWholeWord],
  ] as const)
    if (value !== undefined && typeof value !== 'boolean')
      failArgument(`${target}.${name}`, 'true or false');
  return {
    ...(matchCase === undefined ? {} : { matchCase }),
    ...(matchWholeWord === undefined ? {} : { wholeWord: matchWholeWord }),
  };
}

/** The search text of a call, or `InvalidArgument` naming it under `target`. */
function searchQuery(searchText: unknown, target: string): string {
  const query = text(searchText, `${target}.searchText`);
  // U+FFFC is how the search sees a field: a query with it could match one.
  if (
    query.length === 0 ||
    query.length > MAX_REPLACE_TEXT ||
    /[\r\n\uFFFC]/.test(query) ||
    !xmlText(query)
  )
    failArgument(
      `${target}.searchText`,
      `search text of 1 to ${MAX_REPLACE_TEXT} characters, with no paragraph break, no U+FFFC, and no control characters`
    );
  return query;
}

/**
 * The query and options of a `search` call, without its shape and span, or `InvalidArgument`
 * naming the argument under `target`.
 */
export function searchArguments(
  searchText: unknown,
  options: SearchOptions | undefined,
  target: string
): { query: string; options?: FindOptions } {
  const query = searchQuery(searchText, target);
  const find = searchOptions(options, `${target}.options`);
  return { query, ...(find ? { options: find } : {}) };
}

/**
 * The `replaceText` operation of a `replaceAll` call, without its scope, or `InvalidArgument`
 * naming the argument under `target`.
 */
export function replaceOperation(
  searchText: unknown,
  replacementText: unknown,
  options: SearchOptions | undefined,
  target: string
): Extract<AutomationOperation, { op: 'replaceText' }> {
  const query = searchQuery(searchText, target);
  const replacement = text(replacementText, `${target}.replacementText`);
  if (replacement.length > MAX_REPLACE_TEXT || !xmlText(replacement))
    failArgument(
      `${target}.replacementText`,
      `replacement text of at most ${MAX_REPLACE_TEXT} characters, with no control characters`
    );
  const find = searchOptions(options, `${target}.options`);
  return { op: 'replaceText', query, replacement, ...(find ? { options: find } : {}) };
}
