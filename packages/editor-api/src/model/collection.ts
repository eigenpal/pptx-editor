/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/model/item-collection.ts` (commit 5993bd223945).
//
// A collection lists handles; `items` are proxies over them. `load('items/name')` reads `name`
// on each item in the same `sync()`: the listing answers first, and the item reads go out as a
// second batch pinned to the same revision. A deeper path (`items/textFrame/textRange/text`)
// reads in that same second batch, because a navigation object is addressed through its item, or
// by result reference to the read that finds it. Only a collection under the items (`items/shapes/items/name`) adds a round.
//
// `getItem`, `getItemAt`, and `getItemOrNullObject` answer a proxy at once, over a pending path
// that the lookup in the same sync resolves. A call on it in that sync names the lookup's answer
// by result reference, so `slides.getItemAt(0).shapes.addTextBox('Hi')` is one atomic batch.

import type {
  AutomationHandle,
  AutomationItemKey,
  AutomationOperation,
  AutomationValue,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import type { ClientObject } from '../runtime/client-object.ts';
import { diagnosticOf, fail, failArgument, isPptxEditorError, nearest } from '../runtime/errors.ts';
import { hydrated, hydratedHandleOrNull } from '../runtime/hydrate.ts';
import { LOAD_REQUEST } from '../runtime/internals.ts';
import { itemRequest, type ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import { ModelObject, skippingAbsent, type RefOf } from './model-object.ts';

export abstract class ItemCollection<T extends ClientObject> extends ModelObject {
  /** @internal The read that lists this collection's members. */
  protected abstract listing(ref: RefOf): AutomationOperation;
  /** @internal The proxy for a member the host named. */
  protected abstract itemOf(path: ObjectPath, nullable: boolean): T;
  /** @internal The members a listing answered. */
  protected handlesOf(value: AutomationValue, label: string): readonly AutomationHandle[] {
    return hydrated(value, 'handles', label).handles;
  }

  /**
   * @internal The read that finds one member, or undefined when the host has no lookup for this
   * collection: an index then picks from the listing, and an id is `NotImplemented`.
   */
  protected abstract lookup(
    ref: RefOf,
    key: AutomationItemKey,
    orNull: boolean
  ): AutomationOperation | undefined;

  /** The loaded members. */
  get items(): T[] {
    return this.loadedProperty<T[]>('items');
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    // On a collection, every name and path is about the items.
    const forItems = itemRequest(request);
    this.takeNestedPaths();
    if (forItems) this.#checkItemPaths(request, forItems);
    const skip = request.skip ?? 0;
    const top = request.top;
    const label = `${this.path.label}.items`;
    // Through a collection's items, a listing of a shape that has no members (a group's shapes
    // on a shape that is not a group) answers null and leaves the items not loaded.
    const orNull = this.loadingOrNull;
    this.enqueue(
      'read',
      'items',
      (ref) => (orNull ? skippingAbsent(this.listing(ref)) : this.listing(ref)),
      (value) => {
        if (value.kind === 'null') {
          this.markSkipped(diagnosticOf('not-group'));
          return;
        }
        const handles = this.handlesOf(value, label);
        const end = top === undefined ? handles.length : Math.min(handles.length, skip + top);
        const loaded = handles
          .slice(skip, end)
          .map((handle, i) => this.itemOf(this.memberPath(handle, skip + i), false));
        this.setLoadedProperty('items', loaded);
        if (forItems)
          this.internals.queue.continuing(() => {
            for (const item of loaded) item[LOAD_REQUEST](forItems);
          });
      },
      undefined,
      true,
      forItems !== undefined
    );
  }

  /**
   * Check the item paths of a load now, so a wrong name throws at the call with the nearest
   * name, as on any other object: a stand-in item plans the load, and what it queued is dropped.
   * A first segment that is a typo of `items` (`itmes/name`) suggests `items`.
   */
  #checkItemPaths(request: ResolvedLoadOptions, forItems: ResolvedLoadOptions): void {
    const restore = this.internals.queue.checkpoint();
    try {
      this.itemOf(ObjectPath.pending(`${this.path.label}.items`), false)[LOAD_REQUEST](forItems);
    } catch (error) {
      if (!isPptxEditorError(error) || error.code !== 'InvalidArgument') throw error;
      for (const name of [...request.select, ...request.nested.keys()])
        if (nearest(name, ['items']) === 'items')
          failArgument(
            `${this.path.label}.${name}`,
            "'items/<property>', or a property of the items",
            'items'
          );
      throw error;
    } finally {
      restore();
    }
  }

  /** @internal The path of a member the host named. */
  protected memberPath(handle: AutomationHandle, index: number): ObjectPath {
    return ObjectPath.of(`${this.path.label}.items[${index}]`, handle);
  }

  /** The number of members, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const { result, fill } = ClientResult.create<number>(`${this.path.label}.getCount`);
    const label = `${this.path.label}.getCount`;
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.listing(ref),
      (value) => fill(this.handlesOf(value, label).length)
    );
    return result;
  }

  /** @internal A member by key, as a proxy over the lookup queued now. */
  protected find(key: AutomationItemKey, name: string, orNull: boolean): T {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const item = this.itemOf(path, orNull);
    const label = path.label;
    if (this.hasLookup) {
      this.enqueue(
        'read',
        name,
        (ref) => this.lookup(ref, key, orNull)!,
        (value) => {
          const handle = hydratedHandleOrNull(value, label);
          if (handle) path.resolveTo(handle);
          else path.resolveNull();
        },
        path
      );
      return item;
    }
    if (!('index' in key)) fail({ code: 'NotImplemented', target: label });
    // No lookup: pick from the listing, which a call on the item waits for.
    this.enqueue(
      'read',
      name,
      (ref) => this.listing(ref),
      (value) => {
        const handle = this.handlesOf(value, label)[key.index];
        if (handle) path.resolveTo(handle);
        else if (orNull) path.resolveNull();
        else fail({ code: 'ItemNotFound', target: label });
      },
      path,
      false
    );
    return item;
  }

  /** @internal Whether `lookup` answers for this collection. */
  protected get hasLookup(): boolean {
    return true;
  }

  /** @internal Check an index argument. */
  protected indexArg(index: unknown): number {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return index;
  }

  /** @internal Check a key argument. */
  protected keyArg(key: unknown, name: string): string {
    if (typeof key !== 'string' || key.length === 0)
      failArgument(`${this.path.label}.${name}`, 'a non-empty string');
    return key;
  }
}
