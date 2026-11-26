/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The matches of `TextRange.search`, an SDK extension: the `RangeCollection` shape of the DOCX
// Editor API (`items`, `getFirst`, `getFirstOrNullObject`, `getLast`) with the item access of
// this object model (`getCount`, `getItemAt`). One `findText` read answers each call, over the
// text as the sync finds it (`core/src/automation/plan-find.ts`).

import type { AutomationOperation, AutomationTextSpan } from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ModelObject, type RefOf } from './model-object.ts';
import { TextRange } from './text.ts';

/**
 * The text ranges that a search found, in text order. Load `items`, or read one match with
 * `getItemAt`, `getFirst`, or `getLast`. Each match is a range of the searched shape's text.
 *
 * An SDK extension: Office.js has no such type.
 *
 * @public
 */
export class TextRangeCollection extends ModelObject {
  readonly #shape: ObjectPath;
  readonly #plan: (ref: RefOf) => AutomationOperation;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    plan: (ref: RefOf) => AutomationOperation
  ) {
    super(context, path);
    this.#shape = shape;
    this.#plan = plan;
  }

  /** The loaded matches. */
  get items(): TextRange[] {
    return this.loadedProperty<TextRange[]>('items');
  }

  /** The number of matches, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue('read', 'getCount', this.#plan, (value) =>
      fill(hydrated(value, 'textSpans', label).spans.length)
    );
    return result;
  }

  /** The match at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): TextRange {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.#one(`getItemAt(${index})`, (spans) => spans[index], false);
  }

  /** The first match; `ItemNotFound` at the sync when there is none. */
  getFirst(): TextRange {
    return this.#one('getFirst()', (spans) => spans[0], false);
  }

  /** The first match, or a null object when there is none. */
  getFirstOrNullObject(): TextRange {
    return this.#one('getFirstOrNullObject()', (spans) => spans[0], true);
  }

  /** The last match; `ItemNotFound` at the sync when there is none. */
  getLast(): TextRange {
    return this.#one('getLast()', (spans) => spans[spans.length - 1], false);
  }

  #one(
    name: string,
    pick: (spans: readonly AutomationTextSpan[]) => AutomationTextSpan | undefined,
    nullable: boolean
  ): TextRange {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const { range, settle } = TextRange.later(this.context, path, nullable);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      this.#plan,
      (value) => {
        const span = pick(hydrated(value, 'textSpans', label).spans);
        if (!span) {
          if (nullable) return settle(null);
          fail({ code: 'ItemNotFound', target: label });
        }
        settle({ shape: this.#shape.handle(), span });
      },
      path,
      false
    );
    return range;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    // A match is a shape and a span. Load the text or the font of a match on the match itself.
    for (const name of [...request.select, ...request.itemSelect])
      if (name !== 'items')
        failArgument(
          `${this.path.label}.${name}`,
          'items; load the properties of a match on the match',
          nearest(name, ['items'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue('read', 'items', this.#plan, (value) => {
      const spans = hydrated(value, 'textSpans', label).spans;
      const shape = this.#shape;
      this.setLoadedProperty(
        'items',
        spans.map(
          (span, i) =>
            new TextRange(
              this.context,
              ObjectPath.of(`${label}[${i}]`, shape.handle()),
              shape,
              span
            )
        )
      );
    });
  }
}
