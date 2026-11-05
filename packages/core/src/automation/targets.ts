// Where a handle points in the working package: the slide, layout, or master part, and the
// shape element. Every lookup reads the batch's working package, so an operation sees the writes
// before it in the same batch.

import {
  attr,
  child,
  elements,
  presentationInfo,
  relationshipsOf,
  REL,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
  type SlideRef,
} from '../store/index.ts';
import { CommandRefusal, type OpBatch } from './commands/op-batch.ts';
import type { AutomationOwner, AutomationTarget, HandleTable } from './handles.ts';
import type {
  AutomationErrorCode,
  AutomationHandle,
  AutomationObjectKind,
  AutomationRef,
  AutomationValue,
} from './protocol.ts';

/** A refusal raised while planning one operation. The host reports it at that operation. */
export class AutomationFailure extends Error {
  constructor(
    readonly code: AutomationErrorCode,
    message: string,
    readonly detail?: string
  ) {
    super(message);
  }
}

/** Throw an `invalid-argument` failure that names the argument. */
export function badArgument(name: string): never {
  throw new AutomationFailure('invalid-argument', `invalid ${name}`, name);
}

/** What one operation plans against. */
export interface PlanContext {
  readonly batch: OpBatch;
  readonly handles: HandleTable;
  /** The values earlier operations of this batch answered with, by index. */
  readonly results: readonly (AutomationValue | undefined)[];
}

/** Thrown when a reference names a null result: a read answers null, a write fails. */
export class NullReference extends Error {}

/** The handle a ref names: itself, or the handle an earlier operation answered with. */
export function handleOf(ctx: PlanContext, ref: AutomationRef): AutomationHandle {
  if (typeof ref !== 'object' || ref === null)
    throw new AutomationFailure('invalid-handle', 'not a handle');
  if ('result' in ref) {
    const index = ref.result;
    if (!Number.isInteger(index) || index < 0 || index >= ctx.results.length)
      throw new AutomationFailure('invalid-handle', 'no such result', 'result');
    const value = ctx.results[index];
    if (value?.kind === 'null') throw new NullReference();
    if (value?.kind !== 'handle')
      throw new AutomationFailure('invalid-handle', 'that result is not an object', 'result');
    return value.handle;
  }
  return ref;
}

/** The target a ref names, checked against the kinds the operation takes. */
export function targetOf<K extends AutomationObjectKind>(
  ctx: PlanContext,
  ref: AutomationRef,
  ...kinds: K[]
): Extract<AutomationTarget, { kind: K }> {
  const target = ctx.handles.resolve(handleOf(ctx, ref));
  if (!target || !(kinds as string[]).includes(target.kind))
    throw new AutomationFailure('invalid-handle', 'a handle of another kind or host');
  return target as Extract<AutomationTarget, { kind: K }>;
}

/** The slide a sldId names, with its index, or a stale-handle failure. */
export function slideOf(pkg: PptxPackage, sldId: string): { ref: SlideRef; index: number } {
  const slides = presentationInfo(pkg).slides;
  const index = slides.findIndex((s) => s.sldId === sldId);
  if (index < 0) throw new AutomationFailure('invalid-handle', 'that slide is gone', 'gone');
  return { ref: slides[index]!, index };
}

/** The part that holds an owner's shape tree. */
export function ownerPart(pkg: PptxPackage, owner: AutomationOwner): string {
  if (owner.kind === 'slide') return slideOf(pkg, owner.sldId).ref.partName;
  const root = xmlPartOf(pkg, owner.part)?.root;
  const local = owner.kind === 'slideLayout' ? 'sldLayout' : 'sldMaster';
  if (!root || root.local !== local)
    throw new AutomationFailure('invalid-handle', 'that part is gone', 'gone');
  return owner.part;
}

/** The slide index of an owner, or undefined for a layout or master. */
export function ownerSlideIndex(pkg: PptxPackage, owner: AutomationOwner): number | undefined {
  return owner.kind === 'slide' ? slideOf(pkg, owner.sldId).index : undefined;
}

const SHAPE_KINDS: ReadonlySet<string> = new Set(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame']);

/** Whether an element is a shape of a shape tree (`p:sp`, `p:pic`, `p:grpSp`, …). */
export function isShape(el: Element): boolean {
  return SHAPE_KINDS.has(el.kind) || (el.local === 'contentPart' && el.prefix === 'p');
}

/** The non-visual properties element (`p:nvSpPr`, `p:nvPicPr`, …) of a shape. */
export function nvOf(el: Element): Element | undefined {
  return elements(el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
}

/** `p:cNvPr` of a shape. */
export function cNvPrOf(el: Element): Element | undefined {
  return child(nvOf(el), 'p:cNvPr');
}

/** The shapes directly in a tree or group, in z-order. */
export function shapesIn(parent: Element | undefined): Element[] {
  return elements(parent).filter(isShape);
}

/** A shape by `p:cNvPr/@id`, anywhere in a tree (groups nest up to 32 deep). */
export function findShape(
  tree: Element | undefined,
  shapeId: string
): { el: Element; parent: Element; level: number } | undefined {
  const visit = (
    parent: Element,
    level: number
  ): { el: Element; parent: Element; level: number } | undefined => {
    if (level > 32) return undefined;
    for (const el of shapesIn(parent)) {
      if (attr(cNvPrOf(el), 'id') === shapeId) return { el, parent, level };
      if (el.kind === 'grpSp') {
        const hit = visit(el, level + 1);
        if (hit) return hit;
      }
    }
    return undefined;
  };
  return tree ? visit(tree, 0) : undefined;
}

/** A resolved shape: its part, element, parent, and group depth. */
export interface ShapeLocation {
  readonly owner: AutomationOwner;
  readonly part: string;
  readonly el: Element;
  readonly parent: Element;
  readonly level: number;
  readonly shapeId: string;
}

/** The shape a ref names, in the working package. */
export function shapeAt(ctx: PlanContext, ref: AutomationRef): ShapeLocation {
  const target = targetOf(ctx, ref, 'shape');
  const part = ownerPart(ctx.batch.pkg, target.owner);
  const hit = findShape(shapeTreeOf(ctx.batch.pkg, part), target.shapeId);
  if (!hit) throw new AutomationFailure('invalid-handle', 'that shape is gone', 'gone');
  return { owner: target.owner, part, ...hit, shapeId: target.shapeId };
}

/** How many shapes of a tree carry `p:cNvPr/@id` = `shapeId`, counting up to two. */
function shapesWithId(tree: Element | undefined, shapeId: string): number {
  let count = 0;
  const visit = (parent: Element, level: number): void => {
    if (level > 32 || count > 1) return;
    for (const el of shapesIn(parent)) {
      if (attr(cNvPrOf(el), 'id') === shapeId) count += 1;
      if (el.kind === 'grpSp') visit(el, level + 1);
    }
  };
  if (tree) visit(tree, 0);
  return count;
}

/**
 * Refuse a write to a shape whose `p:cNvPr/@id` another shape of the same tree also carries.
 * The id must be unique (`ST_DrawingElementId`, ECMA-376 Part 1, 19.3.1.12 cNvPr), and a handle
 * names a shape by it, so a write could reach the wrong shape. A reference that is not a shape
 * passes.
 */
export function assertUniqueShape(ctx: PlanContext, ref: unknown): void {
  if (ref === undefined) return;
  const target = ctx.handles.resolve(handleOf(ctx, ref as AutomationRef));
  if (target?.kind !== 'shape') return;
  const part = ownerPart(ctx.batch.pkg, target.owner);
  if (shapesWithId(shapeTreeOf(ctx.batch.pkg, part), target.shapeId) > 1)
    throw new AutomationFailure(
      'unsupported-content',
      'another shape on the same slide has that id',
      'duplicate-id'
    );
}

/** The shape tree element (or the group) a container ref names, with its owner. */
export function containerAt(
  ctx: PlanContext,
  ref: AutomationRef
): { owner: AutomationOwner; part: string; tree: Element; group?: ShapeLocation } {
  const target = targetOf(ctx, ref, 'slide', 'slideLayout', 'slideMaster', 'shape');
  if (target.kind === 'shape') {
    const group = shapeAt(ctx, ref);
    if (group.el.kind !== 'grpSp')
      throw new AutomationFailure('invalid-handle', 'that shape is not a group', 'not-group');
    return { owner: group.owner, part: group.part, tree: group.el, group };
  }
  const owner: AutomationOwner =
    target.kind === 'slide' ? target : { kind: target.kind, part: target.part };
  const part = ownerPart(ctx.batch.pkg, owner);
  const tree = shapeTreeOf(ctx.batch.pkg, part);
  if (!tree) throw new AutomationFailure('unsupported-content', 'that part has no shape tree');
  return { owner, part, tree };
}

/** Mint the handle of a shape. */
export function shapeHandle(
  ctx: PlanContext,
  owner: AutomationOwner,
  el: Element
): AutomationHandle<'shape'> {
  const shapeId = attr(cNvPrOf(el), 'id');
  if (!shapeId) throw new AutomationFailure('unsupported-content', 'a shape with no id');
  return ctx.handles.mint({ kind: 'shape', owner, shapeId });
}

/** The masters in `p:sldMasterIdLst` order, with their `p:sldMasterId/@id`. */
export function mastersOf(pkg: PptxPackage): { part: string; id: string }[] {
  const info = presentationInfo(pkg);
  const root = xmlPartOf(pkg, info.presentationPart)?.root;
  const rels = relationshipsOf(pkg, info.presentationPart);
  const out: { part: string; id: string }[] = [];
  for (const m of elements(child(root, 'p:sldMasterIdLst'))) {
    const rel = rels.find((r) => r.id === attr(m, 'r:id') && !r.external);
    if (rel?.target && xmlPartOf(pkg, rel.target)?.root.local === 'sldMaster')
      out.push({ part: rel.target, id: attr(m, 'id') ?? '' });
  }
  return out;
}

/** The layouts of a master in `p:sldLayoutIdLst` order, with their `p:sldLayoutId/@id`. */
export function layoutsOf(pkg: PptxPackage, master: string): { part: string; id: string }[] {
  const root = xmlPartOf(pkg, master)?.root;
  const rels = relationshipsOf(pkg, master);
  const out: { part: string; id: string }[] = [];
  for (const l of elements(child(root, 'p:sldLayoutIdLst'))) {
    const rel = rels.find((r) => r.id === attr(l, 'r:id') && !r.external);
    if (rel?.target && xmlPartOf(pkg, rel.target)?.root.local === 'sldLayout')
      out.push({ part: rel.target, id: attr(l, 'id') ?? '' });
  }
  return out;
}

/** The master a layout belongs to. */
export function masterOfLayout(pkg: PptxPackage, layout: string): string | undefined {
  return relationshipsOf(pkg, layout).find((r) => r.type === REL.slideMaster && !r.external)
    ?.target;
}

/** Run an edit from `commands/`, turning its refusal into a protocol failure. */
export function edit<T>(run: () => T): T {
  try {
    return run();
  } catch (e) {
    if (e instanceof CommandRefusal) {
      // A name the store cannot carry is the caller's argument, not a failed edit.
      if (e.reason === 'errors.invalidName') badArgument('name');
      const detail = e.detail ? `${e.reason}:${e.detail}` : e.reason;
      throw new AutomationFailure(
        e.reason.startsWith('disabled.') ? 'unsupported-content' : 'transaction-refused',
        'the edit was refused',
        detail
      );
    }
    throw e;
  }
}

/** A finite number in a range, or an `invalid-argument` failure that names it. */
export function numberIn(value: unknown, name: string, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    badArgument(name);
  return value;
}

/** A non-negative integer below `limit`, or an `invalid-argument` failure. */
export function indexIn(value: unknown, name: string, limit: number): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value >= limit)
    badArgument(name);
  return value;
}

/** EMU per point (ECMA-376 Part 1, 20.1.10.16 ST_Coordinate). */
export const EMU = 12_700;

/** Points to whole EMU. */
export const toEmu = (points: number): number => Math.round(points * EMU);

/** EMU to points, rounded to 1/100 pt so reads do not carry float noise. */
export const toPoints = (emu: number): number => Math.round((emu / EMU) * 100) / 100;
