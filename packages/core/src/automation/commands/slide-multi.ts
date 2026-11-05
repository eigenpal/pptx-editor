// Commands over several selected slides, as the slide pane runs them: delete them,
// duplicate them (the copies follow the last selected slide, in order), or move them as one
// block. Each command is one transaction, so one undo step.

import { path, presentationInfo, xmlPartOf } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { deleteSlide, duplicateSlide } from './slide-commands.ts';
import { blockIsInSection, placeBlock } from './slide-structure.ts';

/** The selected slides as distinct indices in ascending order. Refuses an empty or bad set. */
export function slideIndices(batch: OpBatch, indices: readonly number[]): number[] {
  const count = presentationInfo(batch.pkg).slides.length;
  const out = [...new Set(indices)].sort((a, b) => a - b);
  if (out.length === 0 || out.some((i) => !Number.isInteger(i) || i < 0 || i >= count))
    throw new CommandRefusal('errors.unsupported', 'slides');
  return out;
}

/**
 * Put the slide list in `order` (`p:sldId` node ids). Slide k moves to element position k; a
 * move removes the node and inserts it at the index, so earlier positions never shift.
 */
export function reorderSlides(batch: OpBatch, order: readonly string[]): void {
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const list = () => path(xmlPartOf(batch.pkg, presentation)?.root, 'p:sldIdLst');
  order.forEach((nodeId, k) => {
    const kids = list()?.children ?? [];
    const elementsBefore = kids.filter((c) => c.type === 'element');
    if (elementsBefore[k]?.id === nodeId) return;
    const index = kids.findIndex((c) => c === elementsBefore[k]);
    batch.push({
      op: 'moveNode',
      part: presentation,
      node: nodeId,
      parent: list()!.id,
      index: index < 0 ? kids.length : index,
    });
  });
}

/** Delete the selected slides. Returns the slide to show after: the one where the first was. */
export function deleteSlides(batch: OpBatch, indices: readonly number[]): number {
  const selected = slideIndices(batch, indices);
  const count = presentationInfo(batch.pkg).slides.length;
  if (selected.length >= count) throw new CommandRefusal('disabled.lastSlide');
  for (const index of [...selected].reverse()) deleteSlide(batch, index);
  return Math.min(selected[0], count - selected.length - 1);
}

/** Duplicate the selected slides. The copies follow the last selected slide, in order. */
export function duplicateSlides(batch: OpBatch, indices: readonly number[]): number[] {
  const selected = slideIndices(batch, indices);
  const originals = selected.map((i) => presentationInfo(batch.pkg).slides[i].nodeId);
  const copies: string[] = [];
  for (const nodeId of originals) {
    const index = presentationInfo(batch.pkg).slides.findIndex((s) => s.nodeId === nodeId);
    const at = duplicateSlide(batch, index);
    copies.push(presentationInfo(batch.pkg).slides[at].nodeId);
  }
  const rest = presentationInfo(batch.pkg)
    .slides.map((s) => s.nodeId)
    .filter((id) => !copies.includes(id));
  const start = rest.indexOf(originals[originals.length - 1]) + 1;
  reorderSlides(batch, [...rest.slice(0, start), ...copies, ...rest.slice(start)]);
  placeBlock(batch, copies);
  return copies.map((_, i) => start + i);
}

/**
 * Move the selected slides as one block, in their order, so the block starts at `to`, counted
 * among the slides that do not move. The block joins the section of the slide before it, or
 * `section` (`placeBlockInSection`). Returns the block's new indices.
 */
export function moveSlides(
  batch: OpBatch,
  indices: readonly number[],
  to: number,
  section?: string
): number[] {
  if (!Number.isInteger(to)) throw new CommandRefusal('errors.unsupported', 'to');
  const selected = slideIndices(batch, indices);
  const ids = presentationInfo(batch.pkg).slides.map((s) => s.nodeId);
  const moving = selected.map((i) => ids[i]);
  const rest = ids.filter((_, i) => !selected.includes(i));
  const at = Math.max(0, Math.min(Math.trunc(to), rest.length));
  const order = [...rest.slice(0, at), ...moving, ...rest.slice(at)];
  const moved = order.some((id, i) => id !== ids[i]);
  if (moved) reorderSlides(batch, order);
  // A step across a section header changes the section and keeps the place. A block that
  // stays where it is, in `section` already, needs no ops: no undo step.
  if (moved || (section !== undefined && !blockIsInSection(batch, moving, section)))
    placeBlock(batch, moving, section);
  return moving.map((_, i) => at + i);
}
