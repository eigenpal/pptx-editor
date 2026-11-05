// Office.js `Binding` and `BindingCollection`: named links from an automation client to shapes.
//
// A binding is not file content. ECMA-376 has no element for it, so it never enters the package:
// the host keeps its bindings in memory for its lifetime, by id, each naming a shape by its
// owner and `p:cNvPr/@id`, as a shape handle does. A batch edits a working copy, and the host
// keeps that copy only when the batch commits, so a refused batch changes no binding.
//
// A binding outlives its shape: after the shape is removed, the binding is still listed, and
// `getBindingShape` refuses with `item-not-found`. Shape ids and slide ids are used again (a new
// shape takes the highest id plus one; a new last slide takes the id of a removed one), so a
// binding does not trust the id alone. The host checks every binding after each change of the
// document, by anyone (`sweepBindings`): a binding whose shape is gone is orphaned, and stays
// orphaned until the very shape comes back, which only an undo or a redo does. A new shape that
// takes the old id is another element, so it does not bind.

import type { AutomationOwner } from './handles.ts';
import type { AutomationSelectionPort } from './plan.ts';
import type { AutomationBindingRead, AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  findShape,
  ownerPart,
  shapeAt,
  targetOf,
  type PlanContext,
} from './targets.ts';
import { shapeTreeOf, type Element, type PptxPackage } from '../store/index.ts';

/** The shape a binding names. */
export interface AutomationBindingTarget {
  readonly owner: AutomationOwner;
  readonly shapeId: string;
  /**
   * Every element the bound shape has been in a state the host saw, so an undo that restores
   * one of them is the same shape. Shared by the copies of the target; adding to it is harmless.
   */
  readonly seen: WeakSet<Element>;
  /** The shape was gone at a check: only an element in `seen` binds again. */
  readonly orphaned: boolean;
}

/** The bindings of one host, by id, in the order they were added. */
export type AutomationBindings = Map<string, AutomationBindingTarget>;

/** The most bindings one host keeps. */
export const MAX_BINDINGS = 10_000;

/** What a binding plan needs besides the batch: the working copy of the host's bindings. */
export interface BindingPlanContext extends PlanContext {
  readonly bindings: AutomationBindings;
  readonly selection?: AutomationSelectionPort;
}

/** A binding id: a string of 1 to 255 characters. */
function bindingId(value: unknown): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 255)
    throw new AutomationFailure('invalid-argument', 'invalid binding id', 'id');
  return value;
}

/** The binding with `id`, or `item-not-found`. */
function bindingOf(ctx: BindingPlanContext, id: unknown): AutomationBindingTarget {
  const found = ctx.bindings.get(bindingId(id));
  if (!found) throw new AutomationFailure('item-not-found', 'no binding has that id', 'binding');
  return found;
}

/** Add or replace a binding. A new id past the cap refuses. */
function bind(ctx: BindingPlanContext, id: string, target: AutomationBindingTarget): void {
  if (!ctx.bindings.has(id) && ctx.bindings.size >= MAX_BINDINGS)
    throw new AutomationFailure('invalid-argument', 'too many bindings', 'binding-limit');
  ctx.bindings.set(id, target);
}

/** The element of the shape a binding names in `pkg`, or undefined when it is not that shape. */
function boundElement(pkg: PptxPackage, target: AutomationBindingTarget): Element | undefined {
  let part: string;
  try {
    part = ownerPart(pkg, target.owner);
  } catch {
    return undefined;
  }
  const el = findShape(shapeTreeOf(pkg, part), target.shapeId)?.el;
  if (!el || (target.orphaned && !target.seen.has(el))) return undefined;
  return el;
}

/**
 * Check every binding against `pkg`, a state of the document: a binding whose shape is gone is
 * orphaned, and one whose very shape is back binds again. Changes replace the entries of
 * `bindings`, so a copy that a refused batch drops leaves the host's bindings as they were.
 */
export function sweepBindings(bindings: AutomationBindings, pkg: PptxPackage): void {
  for (const [id, target] of bindings) {
    const el = boundElement(pkg, target);
    if (el) {
      target.seen.add(el);
      if (target.orphaned) bindings.set(id, { ...target, orphaned: false });
    } else if (!target.orphaned) bindings.set(id, { ...target, orphaned: true });
  }
}

/** A new binding target for the shape at `owner` and `shapeId`, which `pkg` holds. */
function targetIn(
  pkg: PptxPackage,
  owner: AutomationOwner,
  shapeId: string
): AutomationBindingTarget {
  const target: AutomationBindingTarget = {
    owner,
    shapeId,
    seen: new WeakSet(),
    orphaned: false,
  };
  const el = boundElement(pkg, target);
  if (el) target.seen.add(el);
  return target;
}

export function listBindings(ctx: BindingPlanContext, owner: AutomationRef): AutomationValue {
  targetOf(ctx, owner, 'presentation');
  const bindings: AutomationBindingRead[] = [];
  for (const id of ctx.bindings.keys()) bindings.push({ id, type: 'shape' });
  return { kind: 'bindings', bindings };
}

/** Bind a shape, or move the binding with that id to it. */
export function addBinding(
  ctx: BindingPlanContext,
  owner: AutomationRef,
  shape: AutomationRef,
  id: unknown
): AutomationValue {
  targetOf(ctx, owner, 'presentation');
  const key = bindingId(id);
  const at = shapeAt(ctx, shape);
  bind(ctx, key, targetIn(ctx.batch.pkg, at.owner, at.shapeId));
  return { kind: 'applied' };
}

/** Bind the one selected shape. Any other selection refuses. */
export function addBindingFromSelection(
  ctx: BindingPlanContext,
  owner: AutomationRef,
  id: unknown
): AutomationValue {
  targetOf(ctx, owner, 'presentation');
  const key = bindingId(id);
  if (!ctx.selection)
    throw new AutomationFailure(
      'unsupported-capability',
      'this host has no selection',
      'selection'
    );
  const shapes = ctx.selection.read().shapes;
  if (!shapes || shapes.shapeIds.length !== 1)
    throw new AutomationFailure(
      'invalid-argument',
      'select exactly one shape',
      'binding-selection'
    );
  const at: AutomationOwner = { kind: 'slide', sldId: shapes.sldId };
  const shapeId = shapes.shapeIds[0]!;
  // The selected shape must be in the working package.
  shapeAt(ctx, ctx.handles.mint({ kind: 'shape', owner: at, shapeId }));
  bind(ctx, key, targetIn(ctx.batch.pkg, at, shapeId));
  return { kind: 'applied' };
}

export function deleteBinding(
  ctx: BindingPlanContext,
  owner: AutomationRef,
  id: unknown
): AutomationValue {
  targetOf(ctx, owner, 'presentation');
  bindingOf(ctx, id);
  ctx.bindings.delete(id as string);
  return { kind: 'applied' };
}

/**
 * The handle of the shape a binding names. A binding whose shape is gone refuses with
 * `item-not-found`, as a lookup of a shape that is not there does.
 */
export function getBindingShape(
  ctx: BindingPlanContext,
  owner: AutomationRef,
  id: unknown
): AutomationValue {
  targetOf(ctx, owner, 'presentation');
  bindingOf(ctx, id);
  // The batch so far may have removed the shape, or put another in its place.
  sweepBindings(ctx.bindings, ctx.batch.pkg);
  const target = bindingOf(ctx, id);
  if (target.orphaned)
    throw new AutomationFailure('item-not-found', 'the bound shape is gone', 'binding-shape');
  return {
    kind: 'handle',
    handle: ctx.handles.mint({ kind: 'shape', owner: target.owner, shapeId: target.shapeId }),
  };
}
