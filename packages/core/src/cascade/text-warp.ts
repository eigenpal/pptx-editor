// The preset text warp of a text body (`a:bodyPr/a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19),
// and the adjust values that a preset geometry and a text warp share (`a:avLst`).

import { isTextWarpName } from '@eigenpal/ooxml-core';
import { attr, child, elements, type Element } from '../store/index.ts';
import type { TextWarp } from './types.ts';

/** The largest adjust value magnitude kept: an `ST_GeomGuideFormula` literal in 32 bits. */
const MAX_ADJUST = 2_147_483_647;

/**
 * The adjust values of an `a:avLst` (`a:gd` with a `val` formula), by guide name. A value
 * outside 32 bits is clamped; a guide with another formula is skipped.
 */
export function readAdjustValues(avLst: Element | undefined): Map<string, number> {
  const adjust = new Map<string, number>();
  for (const gd of elements(avLst)) {
    const name = attr(gd, 'name');
    const m = /^val\s+(-?\d+)$/.exec((attr(gd, 'fmla') ?? '').trim());
    if (name && m) adjust.set(name, Math.max(-MAX_ADJUST, Math.min(Number(m[1]), MAX_ADJUST)));
  }
  return adjust;
}

/**
 * The text warp of a `a:bodyPr` chain (inherited first, own last): the last `a:prstTxWarp` in
 * the chain wins. `textNoShape` and a name outside `ST_TextShapeType` draw flat text.
 */
export function readTextWarp(chain: readonly Element[]): TextWarp | undefined {
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const warp = child(chain[i], 'a:prstTxWarp');
    if (!warp) continue;
    const preset = attr(warp, 'prst') ?? '';
    if (preset === 'textNoShape' || !isTextWarpName(preset)) return undefined;
    return { preset, adjust: readAdjustValues(child(warp, 'a:avLst')) };
  }
  return undefined;
}
