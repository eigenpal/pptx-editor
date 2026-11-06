/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `CustomXmlPart`, `CustomXmlPartCollection`, and `CustomXmlPartScopedCollection`: the
// custom XML parts of a presentation, slide, layout, master, or shape (ECMA-376 Part 1, 22.5).
// The id is the part's `ds:itemID` GUID. XML the consumer writes goes through the host's bounded
// XML reader: no DTD, no entities, and size caps.

import type {
  AutomationCustomXmlPartRead,
  AutomationOperation,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ModelObject, type RefOf } from './model-object.ts';

type Read = AutomationCustomXmlPartRead;

/** Fills a part a read found. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<CustomXmlPart, (read: Read) => void>();

function fillPart(part: CustomXmlPart, read: Read): void {
  FILLERS.get(part)?.(read);
}

/** An item id in its canonical form (braced, upper case), or `InvalidArgument`. */
function idOf(value: unknown, target: string): string {
  const m =
    typeof value === 'string'
      ? /^\{?([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}?$/.exec(
          value.trim()
        )
      : null;
  if (!m) failArgument(target, "a GUID such as '{01234567-89AB-CDEF-0123-456789ABCDEF}'");
  return `{${m[1]!.toUpperCase()}}`;
}

/** XML text: a non-empty string. The host checks that it is well formed. */
function xmlOf(value: unknown, target: string): string {
  if (typeof value !== 'string' || value.length === 0) failArgument(target, 'non-empty XML text');
  return value;
}

/** A fresh item id: a random version 4 GUID, braced and upper case. */
function newItemId(): string {
  return `{${globalThis.crypto.randomUUID().toUpperCase()}}`;
}

/**
 * One custom XML part.
 *
 * @public
 */
export class CustomXmlPart extends ModelObject {
  readonly #owner: ObjectPath;
  #id: string;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    owner: ObjectPath,
    id: string,
    nullable = false
  ) {
    super(context, path, { nullable });
    this.#owner = owner;
    this.#id = id;
    if (id) this.setLoadedProperty('id', id);
    FILLERS.set(this, (read) => {
      this.#id = read.id;
      this.setLoadedProperty('id', read.id);
      this.setLoadedProperty('namespaceUri', read.namespaceUri);
    });
  }

  /** The id: the part's `ds:itemID` GUID, braced and upper case. */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The namespace URI of the root element, or an empty string. */
  get namespaceUri(): string {
    return this.loadedProperty('namespaceUri');
  }

  /** The XML text of the part, after the next `sync()`. */
  getXml(): ClientResult<string> {
    const label = `${this.path.label}.getXml`;
    const { result, fill } = ClientResult.create<string>(label);
    this.enqueue(
      'read',
      'getXml',
      (ref) => ({ op: 'readCustomXml', owner: ref(this.path), id: this.#id }),
      (value) => fill(hydrated(value, 'text', label).text)
    );
    return result;
  }

  /** Replace the XML text. Text that is not well-formed XML refuses with `InvalidArgument`. */
  setXml(xml: string): void {
    const text = xmlOf(xml, this.at('setXml'));
    this.command('setXml', (ref) => ({
      op: 'setCustomXml',
      owner: ref(this.path),
      id: this.#id,
      xml: text,
    }));
  }

  /** Delete the part from its owner. A part another owner still names stays for that owner. */
  delete(): void {
    this.command('delete', (ref) => ({
      op: 'deleteCustomXmlPart',
      owner: ref(this.path),
      id: this.#id,
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'namespaceUri']);
    if (!names.includes('namespaceUri')) return;
    const label = `${this.path.label}.load`;
    this.enqueue(
      'read',
      'load',
      (ref) => ({ op: 'listCustomXmlParts', owner: ref(this.#owner) }),
      (value) => {
        if (value.kind === 'null') return;
        const found = hydrated(value, 'customXmlParts', label).parts.find((p) => p.id === this.#id);
        if (found) fillPart(this, found);
      }
    );
  }
}

/** What both collections share: the reads of a list, filtered by namespace or not. */
abstract class CustomXmlPartList extends ModelObject {
  readonly #namespaceUri: string | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, namespaceUri?: string) {
    super(context, path);
    this.#namespaceUri = namespaceUri;
  }

  /** The loaded parts. */
  get items(): CustomXmlPart[] {
    return this.loadedProperty<CustomXmlPart[]>('items');
  }

  /** The number of parts, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.#listing(ref),
      (value) => fill(hydrated(value, 'customXmlParts', label).parts.length)
    );
    return result;
  }

  /** The part with this id; `ItemNotFound` at the sync when there is none. */
  getItem(id: string): CustomXmlPart {
    const key = idOf(id, `${this.path.label}.getItem`);
    return this.find('getItem()', (all) => all.find((p) => p.id === key), key, false);
  }

  /** The part at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): CustomXmlPart {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.find(`getItemAt(${index})`, (all) => all[index], undefined, false);
  }

  /** The part with this id, or a null object. */
  getItemOrNullObject(id: string): CustomXmlPart {
    const key = idOf(id, `${this.path.label}.getItemOrNullObject`);
    return this.find('getItemOrNullObject()', (all) => all.find((p) => p.id === key), key, true);
  }

  #listing(ref: RefOf): AutomationOperation {
    return {
      op: 'listCustomXmlParts',
      owner: ref(this.path),
      ...(this.#namespaceUri === undefined ? {} : { namespaceUri: this.#namespaceUri }),
    };
  }

  /** A part over a pending path that a read resolves to the owner, or null. */
  protected find(
    name: string,
    pick: (all: readonly Read[]) => Read | undefined,
    id: string | undefined,
    orNull: boolean
  ): CustomXmlPart {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const part = new CustomXmlPart(this.context, path, this.path, id ?? '', orNull);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      (ref) => this.#listing(ref),
      (value) => {
        const found = pick(hydrated(value, 'customXmlParts', label).parts);
        if (!found) {
          if (orNull) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: label });
        }
        fillPart(part, found);
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return part;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    for (const name of [...request.select, ...request.itemSelect])
      if (!['items', 'id', 'namespaceUri'].includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          'a property of custom XML parts: items, id, or namespaceUri',
          nearest(name, ['items', 'id', 'namespaceUri'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => this.#listing(ref),
      (value) => {
        const all = hydrated(value, 'customXmlParts', label).parts;
        const end = request.top === undefined ? all.length : (request.skip ?? 0) + request.top;
        this.setLoadedProperty(
          'items',
          all.slice(request.skip ?? 0, end).map((read, i) => {
            const path = ObjectPath.of(`${label}[${i}]`, this.path.handle());
            const part = new CustomXmlPart(this.context, path, this.path, read.id);
            fillPart(part, read);
            return part;
          })
        );
      }
    );
  }
}

/**
 * The custom XML parts of a presentation, slide, layout, master, or shape.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const parts = context.presentation.customXmlParts;
 *   const part = parts.add('<review xmlns="urn:example:review"><status>open</status></review>');
 *   part.load('id');
 *   await context.sync();
 *   console.log(part.id);
 * });
 * ```
 *
 * @public
 */
export class CustomXmlPartCollection extends CustomXmlPartList {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /**
   * Add a part with this XML text. Text that is not well-formed XML, or that declares a DTD or
   * an entity, refuses with `InvalidArgument` at the sync.
   */
  add(xml: string): CustomXmlPart {
    const text = xmlOf(xml, `${this.path.label}.add`);
    const id = newItemId();
    const label = `${this.path.label}.add`;
    this.enqueue(
      'write',
      'add',
      (ref) => ({ op: 'addCustomXmlPart', owner: ref(this.path), xml: text, id }),
      (value) => void hydrated(value, 'customXmlParts', label)
    );
    return new CustomXmlPart(
      this.context,
      ObjectPath.derived(`${this.path.label}.add()`, this.path),
      this.path,
      id
    );
  }

  /** The parts whose root element has this namespace URI. */
  getByNamespace(namespaceUri: string): CustomXmlPartScopedCollection {
    if (typeof namespaceUri !== 'string')
      failArgument(`${this.path.label}.getByNamespace`, 'a namespace URI string');
    return new CustomXmlPartScopedCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.getByNamespace()`, this.path),
      namespaceUri
    );
  }
}

/**
 * The custom XML parts of one namespace, from `CustomXmlPartCollection.getByNamespace`.
 *
 * @public
 */
export class CustomXmlPartScopedCollection extends CustomXmlPartList {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, namespaceUri: string) {
    super(context, path, namespaceUri);
  }

  /** The only part; `ItemNotFound` at the sync unless there is exactly one. */
  getOnlyItem(): CustomXmlPart {
    return this.find(
      'getOnlyItem()',
      (all) => (all.length === 1 ? all[0] : undefined),
      undefined,
      false
    );
  }

  /** The only part, or a null object unless there is exactly one. */
  getOnlyItemOrNullObject(): CustomXmlPart {
    return this.find(
      'getOnlyItemOrNullObject()',
      (all) => (all.length === 1 ? all[0] : undefined),
      undefined,
      true
    );
  }
}
