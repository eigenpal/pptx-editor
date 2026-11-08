// The room that justified and distributed alignment adds between the characters of a line.
// Moved out of `text-layout.ts`, which places the lines.

import { isIdeographic } from './line-break.ts';

const EMU_PER_PT = 12_700;
const MARK = /^[\p{M}\uFE00-\uFE0F]$/u;

/** One UTF-16 unit of a line: its segment, its code, and its advance in EMU. */
export interface JustifyUnit {
  readonly seg: number;
  readonly code: number;
  readonly advance: number;
}

/** The width a space may grow to in East Asian justified text, as a share of its font size. */
const EAST_ASIAN_SPACE_EM = 0.5;

/**
 * The room that justified and distributed alignment adds after each unit of a line, `[start,
 * end)`. A line with no East Asian character widens its spaces only. A line with East Asian
 * characters first widens each space to half its font size, then spreads the rest evenly over
 * the gaps between all its characters, Latin letters included (reference PDF of the
 * `east-asian-typography` fixture).
 */
export function justifyExtra(
  units: readonly JustifyUnit[],
  segments: readonly { sizePt: number }[],
  start: number,
  end: number,
  room: number
): Float64Array | undefined {
  const out = new Float64Array(Math.max(0, end - start));
  const spaces: number[] = [];
  let eastAsian = false;
  for (let i = start; i < end; i += 1) {
    if (units[i].code === 32) spaces.push(i);
    else if (isIdeographic(units[i].code)) eastAsian = true;
  }
  if (!eastAsian) {
    if (spaces.length === 0) return undefined;
    for (const i of spaces) out[i - start] = room / spaces.length;
    return out;
  }
  let rest = room;
  if (spaces.length > 0) {
    const share = room / spaces.length;
    for (const i of spaces) {
      const grow = Math.max(
        0,
        segments[units[i].seg].sizePt * EAST_ASIAN_SPACE_EM * EMU_PER_PT - units[i].advance
      );
      const add = Math.min(share, grow);
      out[i - start] += add;
      rest -= add;
    }
  }
  if (rest <= 0) return out;
  const gaps: number[] = [];
  for (let i = start; i < end - 1; i += 1) {
    const next = units[i + 1].code;
    // Never inside a surrogate pair or before a combining mark.
    if (
      (next >= 0xdc00 && next <= 0xdfff) ||
      (next >= 0x0300 && MARK.test(String.fromCharCode(next)))
    )
      continue;
    gaps.push(i);
  }
  if (gaps.length === 0) return out;
  for (const i of gaps) out[i - start] += rest / gaps.length;
  return out;
}
