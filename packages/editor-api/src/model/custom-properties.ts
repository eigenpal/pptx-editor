/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `CustomProperty` and `CustomPropertyCollection`: the custom document properties of
// the presentation (`docProps/custom.xml`, ECMA-376 Part 1, 22.3). Keys compare without case. A
// property of a variant type other than String, Number, Boolean, or Date is not listed, and
// `deleteAll` keeps it. An `add` with its key replaces it, as in Office.js: a key names one
// property.

import type {
  AutomationCustomPropertyRead,
  AutomationCustomPropertyWrite,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import type { DocumentPropertyType } from './enums.ts';
import { ModelObject } from './model-object.ts';

/** The value types a custom property holds. */
type CustomValue = boolean | Date | number | string;

/** Fills a property a read found. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<CustomProperty, (read: AutomationCustomPropertyRead) => void>();

function fillProperty(property: CustomProperty, read: AutomationCustomPropertyRead): void {
  FILLERS.get(property)?.(read);
}

/** The value a read gives, as the object model shows it: a Date for a Date property. */
function shown(read: AutomationCustomPropertyRead): CustomValue {
  return read.type === 'Date' ? new Date(read.value as string) : read.value;
}

/** A value as the protocol writes it, or `InvalidArgument` naming `target`. */
function written(value: unknown, target: string): AutomationCustomPropertyWrite {
  if (typeof value === 'string') return { type: 'String', value };
  if (typeof value === 'number' && Number.isFinite(value)) return { type: 'Number', value };
  if (typeof value === 'boolean') return { type: 'Boolean', value };
  if (value instanceof Date && Number.isFinite(value.getTime()))
    return { type: 'Date', value: value.toISOString() };
  return failArgument(target, 'a string, a finite number, a boolean, or a valid Date');
}

/** A key, or `InvalidArgument`. Up to 255 characters. */
function keyOf(value: unknown, target: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 255)
    failArgument(target, 'a key of 1 to 255 characters');
  return value;
}

const sameKey = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

/**
 * One custom document property: a key, a type, and a value.
 *
 * @public
 */
export class CustomProperty extends ModelObject {
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
    FILLERS.set(this, (read) => {
      this.#key = read.key;
      this.setLoadedProperty('key', read.key);
      this.setLoadedProperty('type', read.type);
      this.setLoadedProperty('value', shown(read));
    });
  }

  /** The key. Lookups compare keys without case. */
  get key(): string {
    return this.loadedProperty('key');
  }
  /** The type of the value. */
  get type(): DocumentPropertyType | 'Boolean' | 'Date' | 'Number' | 'String' {
    return this.loadedProperty('type');
  }
  /** The value. A write of another JavaScript type changes the type too. */
  get value(): boolean | Date | number | string {
    return this.loadedProperty('value');
  }
  set value(value: boolean | Date | number | string) {
    const v = written(value, this.at('value'));
    // The property's path resolves to the presentation, so it names it once its lookup answered.
    this.command('value', (ref) => ({
      op: 'setCustomProperty',
      owner: ref(this.path),
      key: this.#key,
      value: v,
    }));
  }

  /** Delete the property. */
  delete(): void {
    this.command('delete', (ref) => ({
      op: 'deleteCustomProperty',
      owner: ref(this.path),
      key: this.#key,
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['key', 'type', 'value']);
    if (!names.includes('type') && !names.includes('value')) return;
    const label = `${this.path.label}.load`;
    this.enqueue(
      'read',
      'load',
      (ref) => ({ op: 'readCustomProperties', owner: ref(this.#owner) }),
      (value) => {
        if (value.kind === 'null') return;
        const found = hydrated(value, 'customProperties', label).properties.find((p) =>
          sameKey(p.key, this.#key)
        );
        if (found) fillProperty(this, found);
      }
    );
  }
}

/**
 * The custom document properties of the presentation.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const custom = context.presentation.properties.customProperties;
 *   custom.add('Client', 'Northwind');
 *   custom.add('Reviewed', new Date());
 *   await context.sync();
 * });
 * ```
 *
 * @public
 */
export class CustomPropertyCollection extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** The loaded properties. */
  get items(): CustomProperty[] {
    return this.loadedProperty<CustomProperty[]>('items');
  }

  /** Add a property, or set the value and type of the property with this key. */
  add(key: string, value: boolean | Date | number | string): CustomProperty {
    const k = keyOf(key, `${this.path.label}.add`);
    const v = written(value, `${this.path.label}.add`);
    this.command('add', (ref) => ({
      op: 'setCustomProperty',
      owner: ref(this.path),
      key: k,
      value: v,
    }));
    const property = new CustomProperty(
      this.context,
      ObjectPath.derived(`${this.path.label}.add()`, this.path),
      this.path,
      k
    );
    return property;
  }

  /** Delete every listed property. A property of another variant type stays. */
  deleteAll(): void {
    this.command('deleteAll', (ref) => ({ op: 'deleteCustomProperties', owner: ref(this.path) }));
  }

  /** The number of properties, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => ({ op: 'readCustomProperties', owner: ref(this.path) }),
      (value) => fill(hydrated(value, 'customProperties', label).properties.length)
    );
    return result;
  }

  /** The property with this key; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): CustomProperty {
    const k = keyOf(key, `${this.path.label}.getItem`);
    return this.#find('getItem()', (all) => all.find((p) => sameKey(p.key, k)), k, false);
  }

  /** The property at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): CustomProperty {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.#find(`getItemAt(${index})`, (all) => all[index], undefined, false);
  }

  /** The property with this key, or a null object. */
  getItemOrNullObject(key: string): CustomProperty {
    const k = keyOf(key, `${this.path.label}.getItemOrNullObject`);
    return this.#find(
      'getItemOrNullObject()',
      (all) => all.find((p) => sameKey(p.key, k)),
      k,
      true
    );
  }

  /** A property over a pending path that a read resolves to the presentation, or null. */
  #find(
    name: string,
    pick: (
      all: readonly AutomationCustomPropertyRead[]
    ) => AutomationCustomPropertyRead | undefined,
    key: string | undefined,
    orNull: boolean
  ): CustomProperty {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const property = new CustomProperty(this.context, path, this.path, key ?? '', orNull);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      (ref) => ({ op: 'readCustomProperties', owner: ref(this.path) }),
      (value) => {
        const found = pick(hydrated(value, 'customProperties', label).properties);
        if (!found) {
          if (orNull) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: label });
        }
        fillProperty(property, found);
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return property;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    for (const name of [...request.select, ...request.itemSelect])
      if (!['items', 'key', 'type', 'value'].includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          'a property of custom properties: items, key, type, or value',
          nearest(name, ['items', 'key', 'type', 'value'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => ({ op: 'readCustomProperties', owner: ref(this.path) }),
      (value) => {
        const all = hydrated(value, 'customProperties', label).properties;
        const end = request.top === undefined ? all.length : (request.skip ?? 0) + request.top;
        this.setLoadedProperty(
          'items',
          all.slice(request.skip ?? 0, end).map((read, i) => {
            const path = ObjectPath.of(`${label}[${i}]`, this.path.handle());
            const property = new CustomProperty(this.context, path, this.path, read.key);
            fillProperty(property, read);
            return property;
          })
        );
      }
    );
  }
}
