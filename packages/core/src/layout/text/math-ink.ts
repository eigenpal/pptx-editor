// Ink extents of math glyphs. Script placement and the line extent of an equation keep
// clearances between drawn glyphs, not between font boxes, because a font box reaches far above
// and below its glyphs.
//
// Adapted from DOCX Editor (eigenpal/docx-editor, packages/core/src/layout/equation-ink.ts at
// adf0e178c, Apache-2.0): the ink of a glyph is read once per face and character from its
// outline, and a glyph with no outline gives its font box instead, so clearances stay
// conservative. PPTX Editor reads the outline from the face that draws the glyph: the run's
// face, or the fallback face that draws a character the run's face lacks.

import type { ResolvedFace, TextMeasurer } from './measurer.ts';

/** The ink of one glyph, in em of its size, y up from the baseline. */
export interface MathInk {
  /** The lowest ink: negative below the baseline. */
  readonly lo: number;
  /** The highest ink. */
  readonly hi: number;
  /** The glyph has no outline, so the font box of the face that draws it stands in. */
  readonly fromFontBox?: true;
}

// Ink of one character in one face, per measurer. Equations repeat a small set of letters,
// digits, and operators, so most reads hit. The bound keeps hostile files finite.
const MAX_INK_ENTRIES = 4096;
const inkCache = new WeakMap<TextMeasurer, Map<string, MathInk>>();

// Spaces, invisible operators (U+2061 to U+2064), and format characters draw no ink.
const INVISIBLE = /^[\p{White_Space}\p{Default_Ignorable_Code_Point}]+$/u;

/** The text draws no ink: it holds only spaces, invisible operators, and format characters. */
export function isInvisibleMathText(text: string): boolean {
  return INVISIBLE.test(text);
}

const NO_INK: MathInk = Object.freeze({ lo: 0, hi: 0 });

/** The vertical extent of an outline, or undefined for an empty one. */
function outlineExtent(
  outline: ReturnType<NonNullable<TextMeasurer['glyphOutline']>>
): MathInk | undefined {
  if (!outline || outline.length === 0) return undefined;
  let lo = Infinity;
  let hi = -Infinity;
  for (const c of outline) {
    if (c.op === 'Z') continue;
    lo = Math.min(lo, c.y);
    hi = Math.max(hi, c.y);
    if (c.op === 'Q' || c.op === 'C') {
      lo = Math.min(lo, c.y1);
      hi = Math.max(hi, c.y1);
    }
    if (c.op === 'C') {
      lo = Math.min(lo, c.y2);
      hi = Math.max(hi, c.y2);
    }
  }
  return Number.isFinite(lo) ? { lo, hi } : undefined;
}

/**
 * The ink of character `cp` drawn in `face`, in em. The outline comes from the face that draws
 * the character: `face`, or the measurer's fallback face when `face` lacks it. A character with
 * no outline gives the ascent and descent of that face (`fromFontBox`). Invisible characters
 * have no ink.
 */
export function mathInkOf(measurer: TextMeasurer, face: ResolvedFace, cp: number): MathInk {
  let cache = inkCache.get(measurer);
  if (!cache) inkCache.set(measurer, (cache = new Map()));
  // A substituted face shares the id of its substitute but keeps its own metrics, which a
  // glyph with no outline takes: the metrics are part of the key.
  const { ascent, descent } = face.metrics;
  const key = `${face.id}\u0000${ascent}\u0000${descent}\u0000${cp}`;
  const hit = cache.get(key);
  if (hit) return hit;
  let out: MathInk;
  if (isInvisibleMathText(String.fromCodePoint(cp))) out = NO_INK;
  else {
    const drawing = measurer.fallbackFace?.(face, cp) ?? face;
    const extent = outlineExtent(measurer.glyphOutline?.(drawing, cp));
    out = Object.freeze(
      extent ?? { lo: -drawing.metrics.descent, hi: drawing.metrics.ascent, fromFontBox: true }
    );
  }
  if (cache.size < MAX_INK_ENTRIES) cache.set(key, out);
  return out;
}

/** The ink of `text` in `face`: the union of its characters, in em. */
export function mathTextInk(measurer: TextMeasurer, face: ResolvedFace, text: string): MathInk {
  let lo = 0;
  let hi = 0;
  let fromFontBox = false;
  let any = false;
  for (const ch of text) {
    const ink = mathInkOf(measurer, face, ch.codePointAt(0) ?? 32);
    if (ink === NO_INK) continue;
    lo = any ? Math.min(lo, ink.lo) : ink.lo;
    hi = any ? Math.max(hi, ink.hi) : ink.hi;
    any = true;
    fromFontBox ||= ink.fromFontBox === true;
  }
  if (!any) return NO_INK;
  return fromFontBox ? { lo, hi, fromFontBox: true } : { lo, hi };
}
