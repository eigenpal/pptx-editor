/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Adjustments` over the adjust values of a preset geometry
// (`a:prstGeom/a:avLst/a:gd`, ECMA-376 Part 1, 20.1.9.5). The index order is the order of the
// preset's own `a:avLst`; a value the shape does not store reads as the preset's default.

import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { finite, ModelObject } from './model-object.ts';

/** A zero-based index, or `InvalidArgument`. */
function indexOf(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0)
    failArgument(target, 'a whole number of 0 or more');
  return value;
}

/**
 * The adjust values of a shape's preset geometry: the handles that change, for example, the
 * corner radius of a rounded rectangle. A length is a fraction of the shape size (0.5 is half);
 * an angle is in degrees. A write is bounded to the range of the preset's adjust handle.
 *
 * A shape with no preset geometry (a freeform, a group, a table) refuses with `NotSupported`.
 *
 * @public
 */
export class Adjustments extends ModelObject {
  readonly #shape: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: ObjectPath) {
    super(context, path);
    this.#shape = shape;
  }

  /** The number of adjust values that the preset defines. */
  get count(): number {
    return this.loadedProperty('count');
  }

  /** The adjust value at a zero-based index, after the next `sync()`. */
  get(index: number): ClientResult<number> {
    const at = indexOf(index, `${this.path.label}.get`);
    const label = `${this.path.label}.get(${at})`;
    const { result, fill } = ClientResult.create<number>(label);
    const shape = this.#shape;
    this.enqueue(
      'read',
      `get(${at})`,
      (ref) => ({ op: 'readAdjustments', shape: ref(shape) }),
      (value) => {
        const values = hydrated(value, 'adjustments', label).adjustments.values;
        if (at >= values.length) fail({ code: 'ItemNotFound', target: label });
        fill(values[at]!);
      }
    );
    return result;
  }

  /** Set the adjust value at a zero-based index. */
  set(index: number, value: number): void {
    const at = indexOf(index, `${this.path.label}.set`);
    const v = finite(value, `${this.path.label}.set`);
    const shape = this.#shape;
    this.command('set', (ref) => ({
      op: 'setAdjustment',
      shape: ref(shape),
      index: at,
      value: v,
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const shape = this.#shape;
    this.loadScalars(
      this.selection(request, ['count']),
      'adjustments',
      (ref) => ({ op: 'readAdjustments', shape: ref(shape) }),
      (v) => ({ count: v.adjustments.values.length })
    );
  }
}
