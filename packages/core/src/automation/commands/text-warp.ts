// The preset text warp of shapes' text (`a:bodyPr/a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19):
// set one with its adjust values, or remove it. The warp is the first child of `a:bodyPr`
// (`CT_TextBodyProperties`), so the autofit choice, 3-D, and extensions after it stay in schema
// order, and every attribute of `a:bodyPr` (`fromWordArt`, `wrap`, the insets) stays as it is.

import { isTextWarpName, presetTextWarpAdjustments } from '@eigenpal/ooxml-core';
import { attr, child, elements, type Element } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';

/** The adjust values a warp writes, by guide name, rounded and in the handle ranges. */
function warpAdjustments(
  preset: string,
  size: { readonly cx: number; readonly cy: number } | undefined,
  kept: ReadonlyMap<string, number>,
  wanted: Readonly<Record<string, number>>
): Map<string, number> {
  const slots = presetTextWarpAdjustments(preset, size?.cx ?? 0, size?.cy ?? 0, kept) ?? [];
  const out = new Map<string, number>();
  for (const s of slots) if (kept.has(s.name)) out.set(s.name, kept.get(s.name)!);
  for (const [name, raw] of Object.entries(wanted)) {
    // The names written are the warp table's, never the caller's text.
    const slot = slots.find((s) => s.name === name);
    if (!slot || typeof raw !== 'number' || !Number.isFinite(raw))
      throw new CommandRefusal('errors.unsupported', 'adjust value');
    let v = Math.round(raw);
    if (slot.min !== null && v < slot.min) v = Math.ceil(slot.min);
    if (slot.max !== null && v > slot.max) v = Math.floor(slot.max);
    out.set(slot.name, v);
  }
  return out;
}

/** The adjust values of an existing `a:prstTxWarp` (`val` formulas only). */
function keptAdjustments(warp: Element | undefined): Map<string, number> {
  const out = new Map<string, number>();
  for (const gd of elements(child(warp, 'a:avLst'))) {
    const name = attr(gd, 'name');
    const m = /^val\s+(-?\d+)$/.exec((attr(gd, 'fmla') ?? '').trim());
    if (name && m) out.set(name, Number(m[1]));
  }
  return out;
}

/**
 * Set the text warp of shapes with a text body, or remove it (`preset` null). `adjust` sets
 * adjust values by guide name; a value the warp keeps from before stays when the preset does not
 * change. `sizeOf` gives a shape's size, for the handle ranges. Returns the shapes changed.
 */
export function setTextWarp(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  preset: string | null,
  adjust: Readonly<Record<string, number>> | undefined,
  sizeOf: (id: string) => { readonly cx: number; readonly cy: number } | undefined
): string[] {
  if (preset !== null && (typeof preset !== 'string' || !isTextWarpName(preset)))
    throw new CommandRefusal('errors.unsupported', 'text warp');
  if (adjust !== undefined && (typeof adjust !== 'object' || adjust === null))
    throw new CommandRefusal('errors.unsupported', 'adjust value');
  const wanted = adjust ?? {};
  if (preset === null && Object.keys(wanted).length > 0)
    throw new CommandRefusal('errors.unsupported', 'adjust value');
  const out: string[] = [];
  let removed = false;
  for (const id of ids) {
    const bodyPr = child(child(batch.element(part, id), 'p:txBody'), 'a:bodyPr');
    if (!bodyPr) continue;
    const old = child(bodyPr, 'a:prstTxWarp');
    out.push(id);
    if (preset === null) {
      if (old) {
        batch.push({ op: 'removeNode', part, node: old.id });
        removed = true;
      }
      continue;
    }
    const kept = old && attr(old, 'prst') === preset ? keptAdjustments(old) : new Map();
    const values = warpAdjustments(preset, sizeOf(id), kept, wanted);
    const gds = [...values].map(([name, v]) => `<a:gd name="${name}" fmla="val ${v}"/>`).join('');
    if (old) batch.push({ op: 'removeNode', part, node: old.id });
    batch.insert({
      op: 'insertNode',
      part,
      parent: bodyPr.id,
      index: 0,
      spec: specFromXml(`<a:prstTxWarp prst="${preset}"><a:avLst>${gds}</a:avLst></a:prstTxWarp>`),
    });
  }
  if (out.length === 0) throw new CommandRefusal('disabled.noText');
  if (preset === null && !removed) throw new CommandRefusal('disabled.noEffectChange');
  return out;
}
