// Slide writes, slide export, and the hyperlinks of a slide. Each slide-list edit runs on the
// repaired presentation part (`normalizePresentation`), as the editor's slide commands do.

import {
  attr,
  child,
  presentationInfo,
  relationshipsOf,
  shapeTreeOf,
  writePptx,
  type Element,
} from '../store/index.ts';
import { toBase64 } from './commands/base64.ts';
import { OpBatch } from './commands/op-batch.ts';
import {
  addSlide as addSlideEdit,
  deleteSlide as deleteSlideEdit,
  moveSlide as moveSlideEdit,
} from './commands/slide-commands.ts';
import { applyLayout as applyLayoutEdit } from './commands/slide-layouts.ts';
import { normalizePresentation } from './commands/slide-sections.ts';
import { repairTablesForSave } from './commands/table-repair.ts';
import { slideNumberOf } from './commands/text-model.ts';
import { linkSpans, readText, type TextPlanContext } from './plan-text.ts';
import type {
  AutomationHandle,
  AutomationHyperlinkRead,
  AutomationRef,
  AutomationValue,
} from './protocol.ts';
import type { AutomationTextSpan } from './operations.ts';
import {
  AutomationFailure,
  badArgument,
  cNvPrOf,
  edit,
  indexIn,
  layoutsOf,
  mastersOf,
  masterOfLayout,
  shapeAt,
  shapeHandle,
  shapesIn,
  slideOf,
  slidePartOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/** Repair the presentation part once, before the first slide-list edit of a batch. */
function normalized(ctx: PlanContext): void {
  edit(() => normalizePresentation(ctx.batch));
}

export function addSlide(
  ctx: PlanContext,
  index: number | undefined,
  layoutRef: AutomationRef | undefined,
  masterRef: AutomationRef | undefined,
  layoutId?: string,
  masterId?: string
): AutomationValue {
  const pkg = ctx.batch.pkg;
  const slides = presentationInfo(pkg).slides;
  const at = index === undefined ? slides.length : indexIn(index, 'index', slides.length + 1);
  const masters = mastersOf(pkg);
  const byId = (id: unknown, name: string, parts: readonly { part: string; id: string }[]) => {
    if (typeof id !== 'string') badArgument(name);
    const hit = parts.find((p) => p.id === id);
    if (!hit) throw new AutomationFailure('item-not-found', 'no such id', name);
    return hit.part;
  };
  const layoutPart = layoutRef
    ? targetOf(ctx, layoutRef, 'slideLayout').part
    : layoutId !== undefined
      ? byId(
          layoutId,
          'layoutId',
          masters.flatMap((m) => layoutsOf(pkg, m.part))
        )
      : undefined;
  let masterPart = masterRef
    ? targetOf(ctx, masterRef, 'slideMaster').part
    : masterId !== undefined
      ? byId(masterId, 'slideMasterId', masters)
      : undefined;
  // The master of the slide before, then the first master (`AddSlideOptions`).
  const previous = slides[at - 1] ?? slides[0];
  masterPart ??=
    (layoutPart && masterOfLayout(pkg, layoutPart)) ||
    previous?.masterPart ||
    (previous?.layoutPart && masterOfLayout(pkg, previous.layoutPart)) ||
    mastersOf(pkg)[0]?.part;
  if (!masterPart)
    throw new AutomationFailure('unsupported-content', 'the presentation has no slide master');
  const layouts = layoutsOf(pkg, masterPart);
  if (layoutPart && !layouts.some((l) => l.part === layoutPart))
    throw new AutomationFailure(
      'invalid-argument',
      'that layout is not in the slide master',
      'layoutId'
    );
  const layout = layoutPart ?? layouts[0]?.part;
  if (!layout)
    throw new AutomationFailure('unsupported-content', 'the slide master has no layouts');
  normalized(ctx);
  edit(() => addSlideEdit(ctx.batch, at - 1, layout));
  const added = presentationInfo(ctx.batch.pkg).slides[at];
  if (!added) throw new AutomationFailure('transaction-refused', 'the slide was not added');
  return { kind: 'handle', handle: ctx.handles.mint({ kind: 'slide', sldId: added.sldId }) };
}

export function deleteSlide(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  normalized(ctx);
  const { index } = slideOf(ctx.batch.pkg, sldId);
  edit(() => deleteSlideEdit(ctx.batch, index));
  return { kind: 'applied' };
}

export function moveSlide(ctx: PlanContext, ref: AutomationRef, to: number): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const count = presentationInfo(ctx.batch.pkg).slides.length;
  const target = indexIn(to, 'index', count);
  normalized(ctx);
  const { index } = slideOf(ctx.batch.pkg, sldId);
  edit(() => moveSlideEdit(ctx.batch, index, target));
  return { kind: 'applied' };
}

export function applyLayout(
  ctx: PlanContext,
  ref: AutomationRef,
  layoutRef: AutomationRef
): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const { part } = targetOf(ctx, layoutRef, 'slideLayout');
  const { index } = slideOf(ctx.batch.pkg, sldId);
  edit(() => applyLayoutEdit(ctx.batch, index, part));
  return { kind: 'applied' };
}

/**
 * The listed slides as a `.pptx` of their own, in presentation order: a copy of the working
 * package with every other slide deleted. Nothing is committed.
 */
export function exportSlides(ctx: PlanContext, refs: readonly AutomationRef[]): AutomationValue {
  if (!Array.isArray(refs) || refs.length === 0) badArgument('slides');
  const keep = new Set(refs.map((ref) => targetOf(ctx, ref, 'slide').sldId));
  for (const sldId of keep) slideOf(ctx.batch.pkg, sldId);
  const copy = new OpBatch(ctx.batch.pkg, { reservedPartNames: ctx.batch.reservedPartNames });
  edit(() => {
    normalizePresentation(copy);
    const slides = presentationInfo(copy.pkg).slides;
    for (let i = slides.length - 1; i >= 0; i -= 1)
      if (!keep.has(slides[i]!.sldId)) deleteSlideEdit(copy, i);
  });
  return { kind: 'base64', base64: toBase64(writePptx(repairTablesForSave(copy.pkg))) };
}

/** The `a:hlinkClick` of a shape's own properties (`p:cNvPr`), if any. */
function shapeLink(el: Element): Element | undefined {
  return child(cNvPrOf(el), 'a:hlinkClick');
}

/** The address of a link relationship of `part`: its URL or relative file path, or null. */
function linkAddress(ctx: PlanContext, part: string): (rId: string | undefined) => string | null {
  const rels = relationshipsOf(ctx.batch.pkg, part);
  return (rId) => {
    const rel = rels.find((r) => r.id === rId);
    return rel?.external ? (rel.externalTarget ?? rel.relativeTarget ?? null) : null;
  };
}

/** The text links of one text body, as reads on the shape `handle`. */
function textLinks(
  txBody: Element,
  slideNumber: number | undefined,
  handle: AutomationHandle<'shape'>,
  address: (rId: string | undefined) => string | null
): AutomationHyperlinkRead[] {
  return linkSpans(txBody, slideNumber).map((span) => ({
    address: address(span.rId),
    screenTip: span.tooltip,
    type: 'textRange',
    shape: handle,
    start: span.start,
    length: span.length,
  }));
}

export function listHyperlinks(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const part = slidePartOf(ctx.batch.pkg, sldId);
  const address = linkAddress(ctx, part);
  const owner = { kind: 'slide' as const, sldId };
  const n = slideNumberOf(ctx.batch, part);
  const out: AutomationHyperlinkRead[] = [];
  const visit = (parent: Element | undefined, depth: number): void => {
    if (!parent || depth > 32) return;
    for (const el of shapesIn(parent)) {
      const handle: AutomationHandle<'shape'> = shapeHandle(ctx, owner, el);
      const link = shapeLink(el);
      if (link)
        out.push({
          address: address(attr(link, 'r:id')),
          screenTip: attr(link, 'tooltip') ?? null,
          type: 'shape',
          shape: handle,
          start: null,
          length: null,
        });
      const txBody = child(el, 'p:txBody');
      if (txBody) for (const read of textLinks(txBody, n, handle, address)) out.push(read);
      if (el.kind === 'grpSp') visit(el, depth + 1);
    }
  };
  visit(shapeTreeOf(ctx.batch.pkg, part), 0);
  return { kind: 'hyperlinks', hyperlinks: out };
}

/**
 * The text links of a shape's text (`a:rPr/a:hlinkClick`, ECMA-376 Part 1, 21.1.2.3.5) that
 * overlap `span`, or every text link without one. An empty span answers the link that holds the
 * point. Each read keeps the whole span of its link, so a link that starts before the range
 * reads from its own start. A shape link (`p:cNvPr/a:hlinkClick`) is not text and is left out.
 */
export function listTextHyperlinks(
  ctx: TextPlanContext,
  ref: AutomationRef,
  span: AutomationTextSpan | undefined
): AutomationValue {
  // Refuses a shape that holds no text, and a span outside the text.
  readText(ctx, ref, span);
  const at = shapeAt(ctx, ref);
  const txBody = child(at.el, 'p:txBody');
  if (!txBody) return { kind: 'hyperlinks', hyperlinks: [] };
  const links = textLinks(
    txBody,
    slideNumberOf(ctx.batch, at.part),
    shapeHandle(ctx, at.owner, at.el),
    linkAddress(ctx, at.part)
  );
  if (span === undefined) return { kind: 'hyperlinks', hyperlinks: links };
  const from = span.start;
  const to = span.start + span.length;
  const overlaps = (link: AutomationHyperlinkRead): boolean => {
    const start = link.start ?? 0;
    const end = start + (link.length ?? 0);
    return span.length === 0 ? start <= from && from < end : start < to && end > from;
  };
  return { kind: 'hyperlinks', hyperlinks: links.filter(overlaps) };
}
