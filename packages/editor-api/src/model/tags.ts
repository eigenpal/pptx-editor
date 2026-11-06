/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Tag` and `TagCollection`: the customer data tags of a presentation, slide, or shape
// (`p:custDataLst/p:tags`, a `p:tagLst` part; ECMA-376 Part 1, 19.3.3.3). Keys are stored in
// upper case, as the object model gives them.

import type { AutomationTag } from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ModelObject, text } from './model-object.ts';

/** Fills a tag a read found: its key and value. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<Tag, (key: string, value: string) => void>();

function fillTag(tag: Tag, key: string, value: string): void {
  FILLERS.get(tag)?.(key, value);
}

/**
 * One tag: a key and a value.
 *
 * @public
 */
export class Tag extends ModelObject {
  readonly #owner: ObjectPath;
  #key: string;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    owner: ObjectPath,
    key: string,
    nullable = false
  ) {
    super(context, path, { nullable });
    this.#owner = owner;
    this.#key = key;
    if (key) this.setLoadedProperty('key', key);
    FILLERS.set(this, (k, v) => {
      this.#key = k;
      this.setLoadedProperty('key', k);
      this.setLoadedProperty('value', v);
    });
  }

  /** The key, in upper case. */
  get key(): string {
    return this.loadedProperty('key');
  }
  /** The value. */
  get value(): string {
    return this.loadedProperty('value');
  }
  set value(value: string) {
    const v = text(value, this.at('value'));
    // The tag's path resolves to its owner, so it names the owner once its lookup answered.
    this.command('value', (ref) => ({
      op: 'setTag',
      owner: ref(this.path),
      key: this.#key,
      value: v,
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['key', 'value']);
    const label = `${this.path.label}.load`;
    if (!names.includes('value')) return;
    this.enqueue(
      'read',
      'load',
      (ref) => ({ op: 'readTags', owner: ref(this.#owner) }),
      (value) => {
        if (value.kind === 'null') return;
        const tag = hydrated(value, 'tags', label).tags.find((t) => t.key === this.#key);
        if (tag) this.setLoadedProperty('value', tag.value);
      }
    );
  }
}

/** A tag key in upper case, or `InvalidArgument`. */
function keyOf(value: unknown, target: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 255)
    failArgument(target, 'a key of 1 to 255 characters');
  return value.toUpperCase();
}

/**
 * The tags of a presentation, slide, or shape.
 *
 * @public
 */
export class TagCollection extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** The loaded tags. */
  get items(): Tag[] {
    return this.loadedProperty<Tag[]>('items');
  }

  /** Add a tag, or set the value of the tag with this key. */
  add(key: string, value: string): void {
    const k = keyOf(key, `${this.path.label}.add`);
    const v = text(value, this.at('value'));
    this.command('add', (ref) => ({ op: 'setTag', owner: ref(this.path), key: k, value: v }));
  }

  /** Delete the tag with this key. `ItemNotFound` at the sync when there is none. */
  delete(key: string): void {
    const k = keyOf(key, `${this.path.label}.delete`);
    this.command('delete', (ref) => ({ op: 'deleteTag', owner: ref(this.path), key: k }));
  }

  /** The number of tags, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => ({ op: 'readTags', owner: ref(this.path) }),
      (value) => fill(hydrated(value, 'tags', label).tags.length)
    );
    return result;
  }

  /** The tag with this key; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Tag {
    const k = keyOf(key, `${this.path.label}.getItem`);
    return this.#find('getItem()', (tags) => tags.find((t) => t.key === k), k, false);
  }

  /** The tag at a zero-based position. */
  getItemAt(index: number): Tag {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.#find(`getItemAt(${index})`, (tags) => tags[index], undefined, false);
  }

  /** The tag with this key, or a null object. */
  getItemOrNullObject(key: string): Tag {
    const k = keyOf(key, `${this.path.label}.getItemOrNullObject`);
    return this.#find('getItemOrNullObject()', (tags) => tags.find((t) => t.key === k), k, true);
  }

  /**
   * A tag over a pending path that a read of the tags resolves to the owner, or null. Its key is
   * known now for a lookup by key, and after the read for a lookup by position.
   */
  #find(
    name: string,
    pick: (tags: readonly AutomationTag[]) => AutomationTag | undefined,
    key: string | undefined,
    orNull: boolean
  ): Tag {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const tag = new Tag(this.context, path, this.path, key ?? '', orNull);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      (ref) => ({ op: 'readTags', owner: ref(this.path) }),
      (value) => {
        const found = pick(hydrated(value, 'tags', label).tags);
        if (!found) {
          if (orNull) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: label });
        }
        fillTag(tag, found.key, found.value);
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return tag;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    for (const name of [...request.select, ...request.itemSelect])
      if (!['items', 'key', 'value'].includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          'a property of tags: items, key, or value',
          nearest(name, ['items', 'key', 'value'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => ({ op: 'readTags', owner: ref(this.path) }),
      (value) => {
        const tags = hydrated(value, 'tags', label).tags;
        const end = request.top === undefined ? tags.length : (request.skip ?? 0) + request.top;
        this.setLoadedProperty(
          'items',
          tags.slice(request.skip ?? 0, end).map((t, i) => {
            const path = ObjectPath.of(`${label}[${i}]`, this.path.handle());
            const tag = new Tag(this.context, path, this.path, t.key);
            fillTag(tag, t.key, t.value);
            return tag;
          })
        );
      }
    );
  }
}
