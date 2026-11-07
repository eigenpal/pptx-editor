// The preset text warp of shapes' text (`a:bodyPr/a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19):
// set one with its adjust values, or remove it. The warp is the first child of `a:bodyPr`
// (`CT_TextBodyProperties`), so the autofit choice, 3-D, and extensions after it stay in schema
// order, and every attribute of `a:bodyPr` (`fromWordArt`, `wrap`, the insets) stays as it is.

import { isTextWarpName, presetTextWarpAdjustments } from '@eigenpal/ooxml-core';
import { inheritedPlaceholders, type TextWarp } from '../../cascade/index.ts';
import { readAdjustValues, readTextWarp } from '../../cascade/text-warp.ts';
import { attr, child, type Element } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';

/** A value in a handle range, rounded. */
function inRange(raw: number, min: number | null, max: number | null): number {
  let v = Math.round(raw);
  if (min !== null && v < min) v = Math.ceil(min);
  if (max !== null && v > max) v = Math.floor(max);
  return v;
}

/** The adjust values a warp writes, by guide name, rounded and in the handle ranges. */
function warpAdjustments(
  preset: string,
  size: { readonly cx: number; readonly cy: number } | undefined,
  kept: ReadonlyMap<string, number>,
  wanted: Readonly<Record<string, number>>
): Map<string, number> {
  const slots = presetTextWarpAdjustments(preset, size?.cx ?? 0, size?.cy ?? 0, kept) ?? [];
  const out = new Map<string, number>();
  for (const s of slots)
    if (kept.has(s.name)) out.set(s.name, inRange(kept.get(s.name)!, s.min, s.max));
  for (const [name, raw] of Object.entries(wanted)) {
    // The names written are the warp table's, never the caller's text.
    const slot = slots.find((s) => s.name === name);
    if (!slot || typeof raw !== 'number' || !Number.isFinite(raw))
      throw new CommandRefusal('errors.unsupported', 'adjust value');
    out.set(slot.name, inRange(raw, slot.min, slot.max));
  }
  return out;
}

/** The warp a shape inherits from its layout and master placeholders, before its own. */
function inheritedWarp(batch: OpBatch, part: string, el: Element): TextWarp | undefined {
  const chain = inheritedPlaceholders(batch.pkg, part, el)
    .map((e) => child(child(e, 'p:txBody'), 'a:bodyPr'))
    .filter((b): b is Element => b !== undefined)
    .reverse();
  return readTextWarp(chain);
}

const warpXml = (preset: string, values: ReadonlyMap<string, number>): string =>
  `<a:prstTxWarp prst="${preset}"><a:avLst>${[...values]
    .map(([name, v]) => `<a:gd name="${name}" fmla="val ${v}"/>`)
    .join('')}</a:avLst></a:prstTxWarp>`;

/**
 * Set the text warp of shapes with a text body, or remove it (`preset` null). `adjust` sets
 * adjust values by guide name. The values of the warp a shape has, its own or inherited, stay
 * when the preset does not change. Removing a warp that a shape inherits from its layout
 * placeholder writes `textNoShape`, which draws the text flat. `sizeOf` gives a shape's size,
 * for the handle ranges. Returns the shapes changed.
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
  let changed = false;
  for (const id of ids) {
    const el = batch.element(part, id);
    const bodyPr = child(child(el, 'p:txBody'), 'a:bodyPr');
    if (!el || !bodyPr) continue;
    const old = child(bodyPr, 'a:prstTxWarp');
    const inherited = inheritedWarp(batch, part, el);
    out.push(id);
    let xml: string | undefined;
    if (preset === null) {
      // An inherited warp needs an own `textNoShape` to stop it.
      if (inherited) {
        if (old && attr(old, 'prst') === 'textNoShape') continue;
        xml = '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>';
      } else if (!old) continue;
    } else {
      const own = old && attr(old, 'prst') === preset;
      const kept = own
        ? readAdjustValues(child(old, 'a:avLst'))
        : inherited?.preset === preset
          ? inherited.adjust
          : new Map<string, number>();
      xml = warpXml(preset, warpAdjustments(preset, sizeOf(id), kept, wanted));
    }
    changed = true;
    if (old) batch.push({ op: 'removeNode', part, node: old.id });
    if (xml)
      batch.insert({ op: 'insertNode', part, parent: bodyPr.id, index: 0, spec: specFromXml(xml) });
  }
  if (out.length === 0) throw new CommandRefusal('disabled.noText');
  if (!changed) throw new CommandRefusal('disabled.noEffectChange');
  return out;
}
