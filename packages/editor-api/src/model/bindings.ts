/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Binding` and `BindingCollection`: named links from code to shapes.
//
// A binding is not file content: ECMA-376 has no element for it. The runtime keeps its bindings
// in memory for its lifetime, keyed to shapes as shape objects are, and never writes them to
// the package, so a save, a reopen, or another runtime over the same bytes has none. A binding
// outlives its shape: it stays listed, and `getShape` refuses with `ItemNotFound`.

import type { AutomationBindingRead } from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { BindingType } from './enums.ts';
import { ModelObject, pathOf } from './model-object.ts';
import { Shape } from './shapes.ts';

/** A binding id, or `InvalidArgument`: a string of 1 to 255 characters. */
function idOf(value: unknown, target: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 255)
    failArgument(target, 'a string of 1 to 255 characters');
  return value;
}

/** A binding type, or `InvalidArgument`. Only `Shape`. */
function typeOf(value: unknown, target: string): void {
  if (value !== BindingType.shape) failArgument(target, "'Shape'");
}

/** Fills a binding a lookup by position found. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<Binding, (id: string) => void>();

/**
 * A binding: an id that names one shape for the lifetime of the runtime.
 *
 * @public
 */
export class Binding extends ModelObject {
  #id: string;

  /**
   * @internal The path resolves to the presentation; the id names the binding. An empty id: a
   * lookup by position that has not answered, and nothing is loaded until it does.
   */
  constructor(context: RequestContext, path: ObjectPath, id: string, nullable = false) {
    super(context, path, { nullable });
    this.#id = id;
    const fill = (found: string): void => {
      this.#id = found;
      this.setLoadedProperty('id', found);
      this.setLoadedProperty('type', BindingType.shape);
    };
    if (id) fill(id);
    FILLERS.set(this, fill);
  }

  /** The id of the binding. */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** What the binding names. Always `Shape`. */
  get type(): 'Shape' | BindingType {
    return this.loadedProperty('type');
  }

  /** Delete the binding. The shape stays. */
  delete(): void {
    // `ref` first: for a lookup in this batch it waits for the read, which gives the id.
    this.command('delete', (ref) => ({ op: 'deleteBinding', owner: ref(this.path), id: this.#id }));
  }

  /** The shape the binding names. A binding whose shape is gone refuses with `ItemNotFound`. */
  getShape(): Shape {
    return this.navigation(
      'getShape()',
      (path) => new Shape(this.context, path),
      (ref) => ({ op: 'getBindingShape', owner: ref(this.path), id: this.#id })
    );
  }

  /** @internal The id and the type come with the object; a load has nothing more to read. */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, ['id', 'type']);
  }
}

/**
 * The bindings of the runtime, in the order they were added. They live in memory and are never
 * saved with the presentation.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
 *   context.presentation.bindings.add(shape, 'Shape', 'title');
 *   await context.sync();
 * });
 * ```
 *
 * @public
 */
export class BindingCollection extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** The loaded bindings. */
  get items(): Binding[] {
    return this.loadedProperty<Binding[]>('items');
  }

  /** Bind a shape under `id`. A binding that has that id moves to the shape. */
  add(shape: Shape, bindingType: BindingType, id: string): Binding;
  /** Bind a shape under `id`. A binding that has that id moves to the shape. */
  add(shape: Shape, bindingType: 'Shape', id: string): Binding;
  add(shape: Shape, bindingType: BindingType | 'Shape', id: string): Binding {
    const target = `${this.path.label}.add`;
    if (!(shape instanceof Shape)) failArgument(target, 'a Shape');
    typeOf(bindingType, target);
    const key = idOf(id, target);
    const shapePath = pathOf(shape);
    this.command('add', (ref) => ({
      op: 'addBinding',
      owner: ref(this.path),
      shape: ref(shapePath),
      id: key,
    }));
    return new Binding(this.context, ObjectPath.derived(`${target}()`, this.path), key);
  }

  /**
   * Bind the one selected shape under `id`. Needs a runtime over a live editor; any selection
   * other than one shape refuses with `InvalidArgument`.
   */
  addFromSelection(bindingType: BindingType, id: string): Binding;
  /**
   * Bind the one selected shape under `id`. Needs a runtime over a live editor; any selection
   * other than one shape refuses with `InvalidArgument`.
   */
  addFromSelection(bindingType: 'Shape', id: string): Binding;
  addFromSelection(bindingType: BindingType | 'Shape', id: string): Binding {
    const target = `${this.path.label}.addFromSelection`;
    typeOf(bindingType, target);
    const key = idOf(id, target);
    this.command('addFromSelection', (ref) => ({
      op: 'addBindingFromSelection',
      owner: ref(this.path),
      id: key,
    }));
    return new Binding(this.context, ObjectPath.derived(`${target}()`, this.path), key);
  }

  /** The number of bindings, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => ({ op: 'listBindings', owner: ref(this.path) }),
      (value) => fill(hydrated(value, 'bindings', label).bindings.length)
    );
    return result;
  }

  /** The binding with this id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Binding {
    const id = idOf(key, `${this.path.label}.getItem`);
    return this.#find('getItem()', (all) => all.find((b) => b.id === id), id, false);
  }

  /** The binding at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): Binding {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.#find(`getItemAt(${index})`, (all) => all[index], undefined, false);
  }

  /** The binding with this id, or a null object. */
  getItemOrNullObject(key: string): Binding {
    const id = idOf(key, `${this.path.label}.getItemOrNullObject`);
    return this.#find('getItemOrNullObject()', (all) => all.find((b) => b.id === id), id, true);
  }

  /** A binding over a pending path that a listing resolves to the presentation, or null. */
  #find(
    name: string,
    pick: (all: readonly AutomationBindingRead[]) => AutomationBindingRead | undefined,
    id: string | undefined,
    orNull: boolean
  ): Binding {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const binding = new Binding(this.context, path, id ?? '', orNull);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      (ref) => ({ op: 'listBindings', owner: ref(this.path) }),
      (value) => {
        const read = pick(hydrated(value, 'bindings', label).bindings);
        if (!read) {
          if (orNull) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: label });
        }
        FILLERS.get(binding)?.(read.id);
        path.resolveTo(this.path.handle());
      },
      path,
      false
    );
    return binding;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    for (const name of [...request.select, ...request.itemSelect])
      if (!['items', 'id', 'type'].includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          'a property of bindings: items, id, or type',
          nearest(name, ['items', 'id', 'type'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => ({ op: 'listBindings', owner: ref(this.path) }),
      (value) => {
        const all = hydrated(value, 'bindings', label).bindings;
        const end = request.top === undefined ? all.length : (request.skip ?? 0) + request.top;
        this.setLoadedProperty(
          'items',
          all
            .slice(request.skip ?? 0, end)
            .map(
              (read, i) =>
                new Binding(
                  this.context,
                  ObjectPath.of(`${label}[${i}]`, this.path.handle()),
                  read.id
                )
            )
        );
      }
    );
  }
}
