// The text a command formats: the selected shapes' text bodies (a table's cells), the
// selected range in them, and the autofit re-fit after the edit. Split from `commands.ts`, so
// the command table stays under the line cap; the text format commands share it.

import { child, type Element } from '../store/index.ts';
import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import { refit, refitAll } from '../automation/commands/autofit.ts';
import { tblOf } from '../automation/commands/table-model.ts';
import { tableTextTargets } from './table-commands.ts';
import type { ListTarget } from '../automation/commands/list-commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext } from './commands.ts';
import {
  ensureParagraphProps,
  paragraphsForFormatting,
  paragraphsOf,
  runsForFormatting,
} from '../automation/commands/text-model.ts';

/** The slide part of the current slide. */
export function slidePart(ctx: CommandContext): string {
  const ref = ctx.session.info.slides[ctx.slide];
  if (!ref) throw new CommandRefusal('disabled.noEditor');
  return ref.partName;
}

/** Selected shape ids (text selection counts as its shape). */
export function selectedIds(selection: EditorSelection): string[] {
  if (selection.kind === 'shapes') return [...selection.nodeIds];
  if (selection.kind === 'text' || selection.kind === 'cells') return [selection.nodeId];
  return [];
}

/** Text targets: each selected shape's txBody, with the text range when editing text. */
export function textTargets(
  batch: OpBatch,
  part: string,
  selection: EditorSelection
): { txBodyId: string; shapeId: string; range?: { from: TextPoint; to: TextPoint } }[] {
  const out: { txBodyId: string; shapeId: string; range?: { from: TextPoint; to: TextPoint } }[] =
    [];
  for (const id of selectedIds(selection)) {
    const sp = batch.element(part, id);
    // A table: the cell of the caret, the selected cells, or every cell.
    if (tblOf(sp)) out.push(...tableTextTargets(batch, part, selection, id));
    const txBody = child(sp, 'p:txBody');
    if (!txBody) continue;
    out.push({
      txBodyId: txBody.id,
      shapeId: id,
      ...(selection.kind === 'text'
        ? { range: { from: selection.anchor, to: selection.focus } }
        : {}),
    });
  }
  if (out.length === 0)
    throw new CommandRefusal(
      selection.kind === 'none' ? 'disabled.noSelection' : 'disabled.noText'
    );
  return out;
}

/** The paragraph indices a text target covers: the selected ones, or all of them. */
export function listTarget(
  batch: OpBatch,
  part: string,
  t: { txBodyId: string; shapeId: string; range?: { from: TextPoint; to: TextPoint } }
): ListTarget {
  const count = paragraphsOf(batch.element(part, t.txBodyId)).length;
  const lo = t.range ? Math.min(t.range.from.paragraph, t.range.to.paragraph) : 0;
  const hi = t.range ? Math.max(t.range.from.paragraph, t.range.to.paragraph) : count - 1;
  const paragraphs: number[] = [];
  for (let i = Math.max(0, lo); i <= Math.min(hi, count - 1); i += 1) paragraphs.push(i);
  return { shapeId: t.shapeId, txBodyId: t.txBodyId, paragraphs };
}

/**
 * The list targets that a check reads: at most {@link PROBE_TEXT_LIMIT} paragraphs in all for
 * a probe, as `forParagraphs` does, and every one otherwise.
 */
export function probeLists(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  targets: readonly Parameters<typeof listTarget>[2][]
): ListTarget[] {
  const lists = targets.map((t) => listTarget(batch, part, t));
  if (!ctx.probe) return lists;
  let left = PROBE_TEXT_LIMIT;
  const out: ListTarget[] = [];
  for (const t of lists) {
    if (left <= 0) break;
    out.push({ ...t, paragraphs: t.paragraphs.slice(0, left) });
    left -= t.paragraphs.length;
  }
  return out;
}

/** `refitTargets` for the shapes `ids`. */
export function refitIds(
  batch: OpBatch,
  ctx: CommandContext,
  part: string,
  ids: readonly string[]
): void {
  if (!ctx.probe) refitAll(batch, ctx.slide, part, ids, ctx.session.measurer);
}

/** Re-fit autofit text after an edit, unless the context only probes. */
export function refitShape(
  batch: OpBatch,
  ctx: CommandContext,
  part: string,
  nodeId: string
): void {
  if (!ctx.probe) refit(batch, ctx.slide, part, nodeId, ctx.session.measurer);
}

/**
 * Re-fit each shape of the targets once, after all of them changed: every cell of a table is
 * a target of the same frame, and one re-fit lays out the whole table.
 */
export function refitTargets(
  batch: OpBatch,
  ctx: CommandContext,
  part: string,
  targets: readonly { readonly shapeId: string }[]
): void {
  if (ctx.probe) return;
  const ids = new Set(targets.map((t) => t.shapeId));
  refitAll(batch, ctx.slide, part, ids, ctx.session.measurer);
}

/**
 * The runs or paragraphs that a probe (`can`) formats at most. Each toolbar button probes its
 * command on every change, and a probe of every run of a 50,000-run text takes a second per
 * button. No step on one run or paragraph refuses: a refusal comes from the targets, which the
 * probe still checks in full. The command itself (`exec`) formats every run.
 */
export const PROBE_TEXT_LIMIT = 64;

export function forRuns(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  apply: (run: Element) => void,
  wordAtCaret = true
): void {
  const targets = textTargets(batch, part, ctx.selection);
  const word = wordAtCaret && !ctx.caretArm;
  let left = ctx.probe ? PROBE_TEXT_LIMIT : Infinity;
  for (const t of targets)
    for (const run of runsForFormatting(batch, part, t.txBodyId, t.range, word)) {
      if (left-- <= 0) break;
      apply(run);
    }
  refitTargets(batch, ctx, part, targets);
}

export function forParagraphs(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  apply: (pPrId: string) => void
): void {
  const targets = textTargets(batch, part, ctx.selection);
  let left = ctx.probe ? PROBE_TEXT_LIMIT : Infinity;
  for (const t of targets)
    for (const pId of paragraphsForFormatting(batch, part, t.txBodyId, t.range)) {
      if (left-- <= 0) break;
      apply(ensureParagraphProps(batch, part, pId));
    }
  refitTargets(batch, ctx, part, targets);
}
