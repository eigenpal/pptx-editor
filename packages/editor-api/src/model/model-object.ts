/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/model/model-object.ts` (commit 5993bd223945).
//
// The helpers every object of the model plans through: a scalar load that sets only the names
// the consumer selected, a command whose answer is `applied`, and the path of another object.

import type {
  AutomationOperation,
  AutomationRef,
  AutomationValue,
} from '@pptx-editor.dev/core/automation';
import { ClientObject } from '../runtime/client-object.ts';
import {
  diagnosticOf,
  failArgument,
  nearest,
  type PptxEditorDiagnostic,
} from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import { LOAD_REQUEST, PATH } from '../runtime/internals.ts';
import {
  DEFAULT_LOAD,
  loadsAllScalars,
  type ResolvedLoadOptions,
} from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';

/** Turns a path into what an operation names: a handle or a result reference. */
export type RefOf = (path: ObjectPath) => AutomationRef;

/** The path of any proxy, for an operation that names a second object. */
export function pathOf(object: ClientObject): ObjectPath {
  if (!(object instanceof ClientObject)) failArgument(undefined, 'an object of this API');
  return object[PATH];
}

/** Navigation properties a load may name: loading one loads that object. */
export type Navigation = Readonly<Record<string, () => ClientObject>>;

export abstract class ModelObject extends ClientObject {
  /**
   * The scalar properties a load asked for. An empty selection means every scalar property. A
   * navigation property name loads that object instead. Any other name is `InvalidArgument`.
   */
  protected selection(
    request: ResolvedLoadOptions,
    available: readonly string[],
    navigation: Navigation = {}
  ): readonly string[] {
    const navigates = (name: string): boolean =>
      Object.prototype.hasOwnProperty.call(navigation, name);
    // Check every name before anything is queued. A refusal names the nearest property and
    // every property this object loads.
    const all = [...available, ...Object.keys(navigation)];
    const unknown = (name: string, accepted: readonly string[]): never =>
      failArgument(
        `${this.path.label}.${name}`,
        `a property of this object: ${listOf(accepted)}`,
        nearest(name, accepted)
      );
    for (const name of request.select)
      if (!available.includes(name) && !navigates(name)) unknown(name, all);
    for (const name of request.nested.keys())
      if (!navigates(name)) unknown(name, Object.keys(navigation));
    this.takeNestedPaths();
    const names: string[] = loadsAllScalars(request) ? [...available] : [];
    // Below a collection's items, the objects a path reaches skip a shape with no text too.
    const below = (next: ResolvedLoadOptions): ResolvedLoadOptions =>
      request.orNull ? { ...next, orNull: true } : next;
    // Office.js `toJSON()` shows a loaded navigation object under its name.
    for (const name of request.select) {
      if (available.includes(name)) {
        if (!names.includes(name)) names.push(name);
      } else this.setLoadedProperty(name, navigation[name]!()[LOAD_REQUEST](below(DEFAULT_LOAD)));
    }
    for (const [name, deeper] of request.nested)
      this.setLoadedProperty(name, navigation[name]!()[LOAD_REQUEST](below(deeper)));
    return names;
  }

  /**
   * Queue one read whose answer of `kind` fills `names` through `project`. A null answer (the
   * object turned out not to exist in this batch) fills nothing.
   */
  protected loadScalars<K extends AutomationValue['kind']>(
    names: readonly string[],
    kind: K,
    plan: (ref: RefOf) => AutomationOperation,
    project: (value: Extract<AutomationValue, { kind: K }>) => Readonly<Record<string, unknown>>
  ): void {
    if (names.length === 0) return;
    const label = `${this.path.label}.load`;
    const orNull = this.loadingOrNull;
    let planned: AutomationOperation | undefined;
    const send = (ref: RefOf): AutomationOperation => (planned = skippingAbsent(plan(ref)));
    this.enqueue('read', 'load', orNull ? send : plan, (value) => {
      if (value.kind === 'null') {
        this.markSkipped(orNull && planned ? skipReason(planned) : undefined);
        return;
      }
      const all = project(hydrated(value, kind, label));
      for (const name of names) this.setLoadedProperty(name, all[name]);
    });
  }

  /** @internal The consumer-facing path of a member, for the `target` of a refusal. */
  protected at(name: string): string {
    return `${this.path.label}.${name}`;
  }

  /** Queue a command that answers `applied`. Nothing is written until `sync()`. */
  protected command(name: string, plan: (ref: RefOf) => AutomationOperation): void {
    const label = `${this.path.label}.${name}`;
    this.enqueue('write', name, plan, (value) => void hydrated(value, 'applied', label));
  }

  /**
   * A navigation object (`slide.layout`, `shape.parentGroup`) over a pending path, with the read
   * that answers its handle queued now. A call on it in the same sync names that read's answer.
   */
  protected navigation<T>(
    name: string,
    make: (path: ObjectPath) => T,
    plan: (ref: RefOf) => AutomationOperation,
    nullable = false
  ): T {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const object = make(path);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      plan,
      (value) => {
        if (value.kind === 'null' && nullable) return path.resolveNull();
        path.resolveTo(hydrated(value, 'handle', label).handle);
      },
      path
    );
    return object;
  }

  /** @internal */
  protected onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, []);
  }
}

/** Every name, quoted, as `'a', 'b', or 'c'`. */
function listOf(names: readonly string[]): string {
  const quoted = names.map((n) => `'${n}'`);
  if (quoted.length <= 2) return quoted.join(' or ');
  return `${quoted.slice(0, -1).join(', ')}, or ${quoted.at(-1)}`;
}

/** The text reads that answer null for a shape that cannot hold text, when asked to. */
const TEXT_READS = new Set(['readTextFrame', 'readText', 'readFont', 'readParagraph']);

/** Whether a read through a collection skips a shape that does not have what it reads. */
function skippable(op: AutomationOperation): boolean {
  if (TEXT_READS.has(op.op)) return !('cell' in op && op.cell !== undefined);
  return (op.op === 'readShape' && op.group === true) || op.op === 'listShapes';
}

/**
 * A read of a shape's own text, or of a shape as a group, that answers null instead of
 * refusing for a shape that has no text frame or is not a group.
 */
export function skippingAbsent(op: AutomationOperation): AutomationOperation {
  return skippable(op) ? ({ ...op, orNull: true } as AutomationOperation) : op;
}

/** Why a read that `skippingAbsent` made nullable answered null. */
function skipReason(op: AutomationOperation): PptxEditorDiagnostic | undefined {
  if (!skippable(op)) return undefined;
  if (TEXT_READS.has(op.op)) return diagnosticOf('no-text');
  return op.op === 'readShape' || op.op === 'listShapes' ? diagnosticOf('not-group') : undefined;
}

/** A finite number, or `InvalidArgument` naming `target`. */
export function finite(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) failArgument(target, 'a finite number');
  return value;
}

/** A string, or `InvalidArgument` naming `target`. */
export function text(value: unknown, target: string): string {
  if (typeof value !== 'string') failArgument(target, 'a string');
  return value;
}

/** A boolean, or `InvalidArgument` naming `target`. */
export function flag(value: unknown, target: string): boolean {
  if (typeof value !== 'boolean') failArgument(target, 'true or false');
  return value;
}
