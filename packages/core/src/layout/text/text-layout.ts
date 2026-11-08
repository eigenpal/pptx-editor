// Slide text layout: one text body into positioned lines. DOM-free; the measurer is
// injected. All lengths are EMU.
//
// Order of operations (measured on reference PDFs): stored autofit (`fontScale`, `lnSpcReduction`) is
// applied first, then shaping, then character spacing, then greedy line filling at break
// opportunities, then paragraph spacing, alignment, columns, and vertical anchoring.

import { paintPieces } from './paint-clusters.ts';
import type { ResolvedParagraph, ResolvedTextBody, RunProps } from '../../cascade/index.ts';
import { slotLanguage } from '../../cascade/script-fonts.ts';
import { anchorOffset, ROTATED_VERT, shiftLines } from './anchor.ts';
import { AutonumberCounters, autonumberText, symbolBulletText } from './autonumber.ts';
import { bidiCaretX, markerRuns, placeBidiLine, wholeMarker, type BidiLine } from './bidi-line.ts';
import { lineBidiOf, shiftLineBidi, type LineBidi } from './line-visual-order.ts';
import { fontSlotPieces, slotFamily, type FontSlot } from './font-slots.ts';
import {
  breakOpportunities,
  forcedLineEnd,
  isHangingPunctuation,
  isTrailingSpace,
} from './line-break.ts';
import { layoutMath, type MathBox } from './math-layout.ts';
import type { ResolvedFace, TextMeasurer } from './measurer.ts';
import { justifyExtra } from './justify.ts';
import { leadingOffset, paragraphAlignment } from './paragraph-alignment.ts';
import { bidiPieces, paragraphBidiOf, type ParagraphBidi } from './rtl-paragraph.ts';
import type { WarpedRun } from './text-warp.ts';
import type { PathCommand } from '@eigenpal/ooxml-core';
import {
  autofitSizePt,
  baselineInLinePt,
  lineHeightInPt,
  paragraphSpacingPt,
  singleLinePt,
  type LineBasis,
  SCRIPT_SIZE_SCALE,
  SMALL_CAPS_SCALE,
  placeBaseline,
} from './line-metrics.ts';

const EMU_PER_PT = 12_700;
const MAX_LINES = 10_000;

/** One positioned glyph run on a line. */
export interface GlyphRun {
  readonly kind: 'text' | 'bullet';
  /** Display text (after caps). */
  readonly text: string;
  /**
   * Left edge of each UTF-16 unit, EMU from the text box left. Length = text.length + 1. In a
   * run that reads right to left (`rtl`), the values are the cell edges from left to right, and
   * unit `i` draws in cell `text.length - 1 - i`.
   */
  readonly x: readonly number[];
  /** The run reads right to left: its units stand in visual order from right to left. */
  readonly rtl?: true;
  readonly face: ResolvedFace;
  /** Rendered size in points (after autofit and script scaling). */
  readonly sizePt: number;
  /** Baseline shift in EMU (positive raises). */
  readonly rise: number;
  readonly props: RunProps;
  /** Index into the paragraph's runs (-1 for bullets). */
  readonly runIndex: number;
  /** Source offset of `text[0]` within the paragraph. */
  readonly start: number;
  /**
   * The picture of a picture bullet (`a:buBlip`). Such a run has no text, and `x` holds its
   * left and right edge. The picture stands on the baseline.
   */
  readonly picture?: BulletPicture;
}

/** The picture of a picture bullet, as paint draws it. */
export interface BulletPicture {
  /** The part whose relationships resolve `rId` and `svgRId`. */
  readonly part: string;
  /** The relationship of the raster picture (`a:blip/@r:embed`). */
  readonly rId: string;
  /** The relationship of an `asvg:svgBlip` alternative, when present. */
  readonly svgRId?: string;
  /** The picture's height, EMU. */
  readonly height: number;
}

/**
 * The height of a picture bullet over its bullet size: measured on reference PDFs, where the
 * picture is 0.7 times the bullet size high, whatever the text's face, and keeps its aspect.
 */
export const PICTURE_BULLET_SCALE = 0.7;

/**
 * A shape that an equation draws itself: a fraction rule, a radical sign, a stretched
 * delimiter, a brace, or an arrow. Paint fills it (nonzero) with the run's fill.
 */
export interface LineMark {
  /** The outline: x in EMU from the text box left, y in EMU up from the line's baseline. */
  readonly commands: readonly PathCommand[];
  readonly props: RunProps;
  /** Index into the paragraph's runs: the math run. */
  readonly runIndex: number;
}

/** One laid-out line. */
export interface Line {
  readonly paragraph: number;
  /** Top of the line box and its baseline, EMU from the text box top. */
  readonly top: number;
  readonly baseline: number;
  readonly height: number;
  /** Column index (multi-column bodies). */
  readonly column: number;
  /** Source offsets covered: [start, end). */
  readonly start: number;
  readonly end: number;
  /** x of each source offset from start to end inclusive, for caret placement. */
  readonly caretX: readonly number[];
  /** The units in visual order (UAX #9), absent where they read in source order. */
  readonly bidi?: LineBidi;
  readonly runs: readonly GlyphRun[];
  /** Height of the tallest font, for the caret. */
  readonly fontHeight: number;
  /**
   * The bullet or number of an empty list item. An empty item shows none, so paint ignores it;
   * the editor draws it dimmed while the caret is in the paragraph.
   */
  readonly pendingBullet?: GlyphRun;
  /**
   * The marks of the line's equations (`a14:m`). The glyphs of an equation are text runs of the
   * line, each at its own size and rise; the equation is one object character in the source.
   */
  readonly marks?: readonly LineMark[];
}

/**
 * Whether a text layout holds a character. An empty text body still lays out one line, so the
 * caret of a press in it has a position and a height; that line shows nothing, so diagnostics,
 * effect silhouettes, and paint bounds treat the body as no text.
 */
export function hasCharacters(layout: TextLayout): boolean {
  return layout.lines.some((line) => line.end > line.start);
}

/** Text layout result, in the text box's own frame. */
export interface TextLayout {
  /** The inner text box (after insets) in shape-local EMU, before rotation. */
  readonly box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  /** Rotation applied around the box center: 0, 90, or 270 plus `bodyPr/@rot`. */
  readonly rotation: number;
  /**
   * Characters paint draws upright inside a rotated line: East Asian characters for `eaVert`,
   * every character for `wordArtVert`, none otherwise.
   */
  readonly upright: 'none' | 'eastAsian' | 'all';
  readonly lines: readonly Line[];
  /**
   * Height the text needs in its own frame, EMU, insets included (for autofit). With columns,
   * the box height, or more when the last column runs below the box.
   */
  readonly contentHeight: number;
  /** Paragraph source lengths, for caret navigation. */
  readonly paragraphLengths: readonly number[];
  /** Families that were not available and were substituted. */
  readonly substitutions: readonly string[];
  /**
   * The glyphs of a warped text body (`a:prstTxWarp`) as paths in this frame, one entry for each
   * run that draws. Paint fills these instead of drawing the runs as text; the lines keep the
   * flat layout, for the caret and the selection.
   */
  readonly warped?: readonly WarpedRun[];
}

interface Segment {
  readonly runIndex: number;
  readonly start: number;
  readonly text: string;
  readonly props: RunProps;
  readonly face: ResolvedFace;
  readonly sizePt: number;
  readonly rise: number;
  readonly advances: Float64Array;
  readonly isBreak: boolean;
  /** The laid-out equation of a math run, which is one unit as wide as the equation. */
  readonly math?: MathBox;
}

/** The face for a run: its Latin font, or the font of `slot`, bold and italic. */
function faceFor(
  props: RunProps,
  measurer: TextMeasurer,
  family?: string,
  slot: FontSlot = 'latin'
): ResolvedFace {
  const lang = slot === 'ea' ? slotLanguage('ea', props.lang, props.altLang) : undefined;
  return measurer.resolve({
    family: family ?? props.latin,
    bold: props.bold,
    italic: props.italic,
    ...(slot === 'ea' ? { script: 'eastAsian' as const } : {}),
    ...(lang ? { lang } : {}),
  });
}

function displayText(text: string, caps: RunProps['caps']): string {
  if (caps === 'none') return text;
  let out = '';
  for (const ch of text) {
    const up = ch.toUpperCase();
    out += up.length === ch.length ? up : ch;
  }
  return out;
}

/**
 * Split text by the face that draws each character: `face`, or another face for a character
 * it lacks (`TextMeasurer.fallbackFace`). ASCII stays in `face`.
 */
function glyphFallbackPieces(
  text: string,
  face: ResolvedFace,
  measurer: TextMeasurer,
  lang: string | undefined
): { text: string; face: ResolvedFace }[] {
  const fallback = measurer.fallbackFace;
  if (!fallback || !/[^\u0000-\u007f]/.test(text)) return [{ text, face }];
  const out: { text: string; face: ResolvedFace }[] = [];
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i) ?? 0;
    const len = cp > 0xffff ? 2 : 1;
    // A combining mark or joiner stays with the character before it.
    const joins =
      out.length > 0 &&
      cp >= 0x0300 &&
      (cp === 0x200d || /^\p{M}$/u.test(String.fromCodePoint(cp)));
    const drawn = joins
      ? out[out.length - 1].face
      : cp < 0x80
        ? face
        : (fallback.call(measurer, face, cp, lang) ?? face);
    const last = out[out.length - 1];
    if (last && last.face === drawn) last.text += text.slice(i, i + len);
    else out.push({ text: text.slice(i, i + len), face: drawn });
    i += len;
  }
  return out;
}

/**
 * The break opportunities of a paragraph's display text. They do not depend on the font scale,
 * so the `normAutofit` search, which lays out one paragraph at each step, finds them once.
 */
const breaksOf = new WeakMap<ResolvedParagraph, { text: string; breaks: Uint8Array }>();
function paragraphBreaks(p: ResolvedParagraph, text: string): Uint8Array {
  const kept = breaksOf.get(p);
  if (kept?.text === text) return kept.breaks;
  const breaks = breakOpportunities(text, p.eaLineBreak, p.latinLineBreak, p.kinsoku);
  breaksOf.set(p, { text, breaks });
  return breaks;
}

function buildSegments(
  p: ResolvedParagraph,
  measurer: TextMeasurer,
  fontScale: number,
  substitutions: Set<string>,
  stacked = false,
  bidi?: ParagraphBidi
): Segment[] {
  const out: Segment[] = [];
  let offset = 0;
  p.runs.forEach((run, runIndex) => {
    const props = run.props;
    const scriptScale = props.baseline !== 0 ? SCRIPT_SIZE_SCALE : 1;
    const runSize = autofitSizePt(props.size / 100, fontScale);
    const baseSize = runSize * scriptScale;
    const face = faceFor(props, measurer);
    const rise = (props.baseline / 100) * runSize * EMU_PER_PT;
    const spacing = (props.spacing / 100) * EMU_PER_PT;
    if (run.kind === 'math' && run.math) {
      const box = layoutMath(run.math, measurer, fontScale);
      out.push({
        runIndex,
        start: offset,
        text: run.text,
        props,
        face: box.face,
        sizePt: box.sizePt,
        rise: 0,
        advances: Float64Array.of(box.width),
        isBreak: false,
        math: box,
      });
      offset += run.text.length;
      return;
    }
    if (run.kind === 'break') {
      if (face.substituted) substitutions.add(props.latin);
      out.push({
        runIndex,
        start: offset,
        text: '\n',
        props,
        face,
        sizePt: baseSize,
        rise: 0,
        advances: new Float64Array(1),
        isBreak: true,
      });
      offset += 1;
      return;
    }
    const shown = displayText(run.text, props.caps);
    // Small caps: lowercase letters draw as capitals at a reduced size.
    const pieces: { text: string; small: boolean }[] = [];
    if (props.caps === 'small') {
      for (const ch of run.text) {
        const small = ch !== ch.toUpperCase() && ch.toUpperCase().length === ch.length;
        const last = pieces[pieces.length - 1];
        const t = small ? ch.toUpperCase() : ch;
        if (last && last.small === small) last.text += t;
        else pieces.push({ text: t, small });
      }
    } else pieces.push({ text: shown, small: false });
    // Each piece in the font of its characters' slot (`font-slots.ts`).
    const slotted: { text: string; small: boolean; face: ResolvedFace }[] = [];
    for (const piece of pieces) {
      const parts = fontSlotPieces(piece.text, props);
      // An empty run still sizes its line, in its Latin face.
      if (parts.length === 0) {
        if (face.substituted) substitutions.add(props.latin);
        slotted.push({ text: piece.text, small: piece.small, face });
      }
      for (const part of parts) {
        const family = slotFamily(props, part.slot);
        const slotFace = part.slot === 'latin' ? face : faceFor(props, measurer, family, part.slot);
        if (slotFace.substituted) substitutions.add(family);
        const text = piece.text.slice(part.start, part.end);
        const lang = slotLanguage('ea', props.lang, props.altLang);
        for (const run of glyphFallbackPieces(text, slotFace, measurer, lang))
          slotted.push({ text: run.text, small: piece.small, face: run.face });
      }
    }
    let local = 0;
    for (const piece of bidiPieces(slotted, bidi?.levels, offset)) {
      const face = piece.face;
      const sizePt = piece.small ? baseSize * SMALL_CAPS_SCALE : baseSize;
      const kerning = props.kern !== undefined && props.size >= props.kern;
      const em = measurer.advances(face, piece.text, { kerning });
      const adv = new Float64Array(piece.text.length);
      for (let i = 0; i < em.length; i += 1) {
        const lowSurrogate =
          piece.text.charCodeAt(i) >= 0xdc00 && piece.text.charCodeAt(i) <= 0xdfff;
        // `wordArtVert` stacks upright letters: each takes the face's line height, not its width.
        const unit = stacked && em[i] > 0 ? face.metrics.ascent + face.metrics.descent : em[i];
        adv[i] = unit * sizePt * EMU_PER_PT + (lowSurrogate ? 0 : spacing);
      }
      out.push({
        runIndex,
        start: offset + local,
        text: piece.text,
        props,
        face,
        sizePt,
        rise,
        advances: adv,
        isBreak: false,
      });
      local += piece.text.length;
    }
    offset += run.text.length;
  });
  return out;
}

/**
 * The glyphs and marks of an equation at `x` on its line: each glyph piece is a text run with
 * the equation's source offset, so the caret and the hit test treat it as one character.
 */
function placeMath(
  box: MathBox,
  x: number,
  runIndex: number,
  start: number,
  runs: GlyphRun[],
  marks: LineMark[]
): void {
  for (const g of box.glyphs)
    runs.push({
      kind: 'text',
      text: g.text,
      x: g.x.map((v) => v + x),
      face: g.face,
      sizePt: g.sizePt,
      rise: g.y,
      props: g.props,
      runIndex,
      start,
    });
  for (const m of box.marks) marks.push(moveMark({ ...m, runIndex }, x));
}

/**
 * The runs of a bidi line, from its pieces, with the line's visible content starting at `left`.
 * Runs stay in source order, so the run before an offset is the run that holds it.
 */
function placeBidiRuns(
  line: BidiLine,
  left: number,
  start: number,
  units: readonly Unit[],
  segments: readonly Segment[],
  runs: GlyphRun[],
  marks: LineMark[]
): void {
  const placed: GlyphRun[] = [];
  for (const piece of line.pieces) {
    const s = segments[piece.seg];
    const x = piece.x.map((v) => v + left);
    if (s.math) {
      placeMath(s.math, x[0], s.runIndex, start + piece.from, placed, marks);
      continue;
    }
    if (s.isBreak) continue;
    let text = '';
    for (let i = piece.from; i < piece.to; i += 1)
      text += String.fromCharCode(units[start + i].code);
    placed.push({
      kind: 'text',
      text,
      x,
      face: s.face,
      sizePt: s.sizePt,
      rise: s.rise,
      props: s.props,
      runIndex: s.runIndex,
      start: start + piece.from,
      ...(piece.rtl ? { rtl: true as const } : {}),
    });
  }
  placed.sort((a, b) => a.start - b.start);
  for (const run of placed) runs.push(run);
}

/** A mark moved right by `dx`. */
function moveMark(mark: LineMark, dx: number): LineMark {
  if (dx === 0) return mark;
  return {
    ...mark,
    commands: mark.commands.map((c) => {
      if (c.op === 'Z') return c;
      if (c.op === 'Q') return { ...c, x1: c.x1 + dx, x: c.x + dx };
      if (c.op === 'C') return { ...c, x1: c.x1 + dx, x2: c.x2 + dx, x: c.x + dx };
      return { ...c, x: c.x + dx };
    }),
  };
}

interface Unit {
  readonly seg: number;
  readonly code: number;
  advance: number;
}

/** Lay out one text body inside a box (shape-local EMU). */
export function layoutText(
  body: ResolvedTextBody,
  rect: { x: number; y: number; width: number; height: number },
  measurer: TextMeasurer
): TextLayout {
  return layoutTextOn(body, rect, measurer, 'size');
}

/**
 * `layoutText` with the single line that the body's lines are spaced on (`LineBasis`): diagram
 * text takes the face's line (`face`). With `untilOverflow`, a single column stops after the
 * first paragraph that ends below the box: line heights and spacings are never negative, so
 * the text cannot fit after that. Only a fit check reads such a layout (`contentHeight`).
 * @internal
 */
export function layoutTextOn(
  body: ResolvedTextBody,
  rect: { x: number; y: number; width: number; height: number },
  measurer: TextMeasurer,
  basis: LineBasis,
  untilOverflow = false
): TextLayout {
  const bp = body.body;
  const vert = bp.vert;
  const rotated = ROTATED_VERT.has(vert);
  const rotation = (vert === 'vert270' ? 270 : rotated ? 90 : 0) + bp.rot;
  // Rotated text lays out in the swapped frame, then paint rotates it around the box center.
  const outerW = rotated ? rect.height : rect.width;
  const outerH = rotated ? rect.width : rect.height;
  const innerW = Math.max(0, outerW - bp.lIns - bp.rIns);
  const innerH = Math.max(0, outerH - bp.tIns - bp.bIns);
  const fontScale = bp.autofit.kind === 'normal' ? bp.autofit.fontScale : 1;
  const reduction = bp.autofit.kind === 'normal' ? bp.autofit.lineSpaceReduction : 0;
  const columns = bp.numCol;
  const colW = columns > 1 ? (innerW - bp.spcCol * (columns - 1)) / columns : innerW;
  const wrap = bp.wrap === 'square';
  const substitutions = new Set<string>();
  const lines: Line[] = [];
  const paragraphLengths: number[] = [];
  let y = 0;
  const numbering = new AutonumberCounters();
  const overflowAt = untilOverflow && columns <= 1 ? outerH + 1 - bp.tIns - bp.bIns : Infinity;
  let overflowed = false;
  const stackedVert = vert === 'wordArtVert';

  body.paragraphs.forEach((p, pi) => {
    if (overflowed) return;
    const bidi = paragraphBidiOf(p);
    const segments = buildSegments(p, measurer, fontScale, substitutions, stackedVert, bidi);
    const sourceLength = p.runs.reduce((n, r) => n + (r.kind === 'break' ? 1 : r.text.length), 0);
    paragraphLengths.push(sourceLength);
    const firstProps = p.runs.find((r) => r.kind !== 'break')?.props ?? p.endProps;
    const firstSize = autofitSizePt(firstProps.size / 100, fontScale);
    const isFirst = pi === 0;
    const isLast = pi === body.paragraphs.length - 1;
    if (!isFirst || bp.spcFirstLastPara)
      y += paragraphSpacingPt(p.spaceBefore, firstSize) * EMU_PER_PT;

    // Bullet. An empty paragraph shows none, but its would-be bullet still
    // places the caret where the text will start, and the editor draws it dimmed while the
    // caret is there (`Line.pendingBullet`). Empty paragraphs consume no number.
    let marker: GlyphRun | null = null;
    const number = sourceLength > 0 ? numbering.next(p) : numbering.peek(p);
    if (p.bullet?.type === 'picture') {
      const sizePt =
        'pct' in p.bullet.size
          ? firstSize * p.bullet.size.pct
          : autofitSizePt(p.bullet.size.pts / 100, fontScale);
      const height = sizePt * PICTURE_BULLET_SCALE * EMU_PER_PT;
      const x0 = p.marL + p.indent;
      const { part, rId, svgRId, aspect } = p.bullet;
      marker = {
        kind: 'bullet',
        text: '',
        x: [x0, x0 + height * aspect],
        face: faceFor(firstProps, measurer),
        sizePt,
        rise: 0,
        props: firstProps,
        runIndex: -1,
        start: 0,
        picture: { part, rId, ...(svgRId ? { svgRId } : {}), height },
      };
    } else if (p.bullet) {
      const raw =
        p.bullet.type === 'char' ? p.bullet.char : autonumberText(p.bullet.scheme, number ?? 1);
      const sizePt =
        'pct' in p.bullet.size
          ? firstSize * p.bullet.size.pct
          : autofitSizePt(p.bullet.size.pts / 100, fontScale);
      const bulletProps: RunProps = {
        ...firstProps,
        ...(p.bullet.color ? { fill: { type: 'solid', color: p.bullet.color } } : {}),
        underline: 'none',
        strike: 'noStrike',
      };
      // A number in complex-script or East Asian characters draws in the first run's font for
      // that script, not the bullet font (reference PDF of `rtl-text`: `أ-` in the run's
      // `a:cs` face, where the bullet font has no Arabic glyphs).
      const scriptSlot =
        p.bullet.type === 'autonum'
          ? // A copy of the props: the marker must not replace the first run's cached split,
            // which the `normAutofit` search reuses at each step.
            fontSlotPieces(raw, { ...firstProps }).find((s) => s.slot === 'cs' || s.slot === 'ea')
              ?.slot
          : undefined;
      const face =
        (scriptSlot === 'cs' || scriptSlot === 'ea') && slotFamily(firstProps, scriptSlot) !== ''
          ? faceFor(firstProps, measurer, slotFamily(firstProps, scriptSlot), scriptSlot)
          : measurer.resolve({
              family: p.bullet.font ?? firstProps.latin,
              bold: p.bullet.font ? false : firstProps.bold,
              italic: p.bullet.font ? false : firstProps.italic,
            });
      // A Wingdings or Symbol bullet is a byte of that font. When the font is not loaded, the
      // bullet draws its Unicode twin (DOCX Editor's table), not the plain letter; the file
      // keeps its character.
      const text =
        p.bullet.type === 'char'
          ? symbolBulletText(raw, p.bullet.font, () => !face.substituted)
          : raw;
      const em = measurer.advances(face, text);
      const x0 = p.marL + p.indent;
      const xs = [x0];
      for (let i = 0; i < em.length; i += 1) xs.push(xs[i] + em[i] * sizePt * EMU_PER_PT);
      marker = {
        kind: 'bullet',
        text,
        x: xs,
        face,
        sizePt,
        rise: 0,
        props: bulletProps,
        runIndex: -1,
        start: 0,
      };
    }

    // Flatten to units.
    const units: Unit[] = [];
    segments.forEach((s, si) => {
      for (let i = 0; i < s.text.length; i += 1)
        units.push({ seg: si, code: s.text.charCodeAt(i), advance: s.advances[i] });
    });
    const paragraphText = segments.map((s) => s.text).join('');
    const breaks = paragraphBreaks(p, paragraphText);
    const codes = bidi ? units.map((u) => u.code) : [];

    const bullet = sourceLength > 0 ? marker : null;
    const pending = sourceLength > 0 ? null : marker;
    const firstLineStart = marker
      ? Math.max(p.marL, marker.x[marker.x.length - 1])
      : p.marL + p.indent;
    const lineLeft = (first: boolean): number => (first ? firstLineStart : p.marL);
    // The bullet stands at the paragraph's leading edge, and moves with its line: centered and
    // right-aligned text keep it next to the text, at the same distance as on a left-aligned
    // line. In a right-to-left paragraph the bullet's text reads in the paragraph's direction
    // (`markerRuns`).
    const placeMarker = (m: GlyphRun, shift: number): GlyphRun => {
      const x0 = m.x[0];
      const w = m.x[m.x.length - 1] - x0;
      const at = (p.rtl ? colW - x0 - w - shift : x0 + shift) + bp.lIns;
      return { ...m, x: m.x.map((v) => v - x0 + at) };
    };
    const avail = (first: boolean): number => colW - lineLeft(first) - p.marR;

    /**
     * The advance of the tab at unit `i`, which starts at `x` (ECMA-376 Part 1, 21.1.2.1.36
     * tab). The text after a tab runs to the next tab or line end. A left stop starts that
     * text at the stop, a right stop ends it there, a centered stop centers it on the stop,
     * and a decimal stop puts its first full stop there (or ends it there, with no full stop).
     * Text that would start before `x` starts at `x`. Default stops are left stops. In a
     * wrapped line, text at a centered, right, or decimal stop that would pass `lineEnd` moves
     * left to end there (reference PDF of `svg-tab-stops`).
     */
    // Only record-high stops can be the first stop after a cursor. Keep their original
    // order, including the first of equal positions, even for an unsorted input list.
    // Searching this increasing index avoids scanning every stop for every tab character.
    const tabStops: ResolvedParagraph['tabs'][number][] = [];
    let furthestStop = -Infinity;
    for (const stop of p.tabs) {
      if (stop.pos > furthestStop) {
        tabStops.push(stop);
        furthestStop = stop.pos;
      }
    }
    const tabAdvance = (i: number, x: number, lineEnd: number): number => {
      let lo = 0;
      let hi = tabStops.length;
      while (lo < hi) {
        const mid = lo + Math.floor((hi - lo) / 2);
        if (tabStops[mid].pos > x + 1) hi = mid;
        else lo = mid + 1;
      }
      const stop = tabStops[lo];
      if (!stop) {
        const d = p.defaultTab > 0 ? p.defaultTab : 914_400;
        return (Math.floor(x / d) + 1) * d - x;
      }
      if (stop.align !== 'ctr' && stop.align !== 'r' && stop.align !== 'dec') return stop.pos - x;
      let width = 0;
      let beforePoint = -1;
      for (let j = i + 1; j < units.length; j += 1) {
        const code = units[j].code;
        if (code === 9 || code === 10) break;
        if (code === 46 && beforePoint < 0) beforePoint = width;
        width += units[j].advance;
      }
      const span =
        stop.align === 'ctr'
          ? width / 2
          : stop.align === 'dec' && beforePoint >= 0
            ? beforePoint
            : width;
      const advance = stop.pos - x - span;
      return Math.max(0, wrap ? Math.min(advance, lineEnd - x - width) : advance);
    };

    // Greedy fill.
    const lineRanges: { start: number; end: number; forced: boolean }[] = [];
    let start = 0;
    while (start < units.length && lineRanges.length < MAX_LINES) {
      const first = lineRanges.length === 0;
      const left = lineLeft(first);
      const limit = avail(first);
      let x = left;
      let lastBreak = -1;
      let end = start;
      let forced = false;
      for (let i = start; i < units.length; i += 1) {
        const u = units[i];
        if (u.code === 10) {
          end = i + 1;
          forced = true;
          break;
        }
        if (u.code === 9) u.advance = tabAdvance(i, x, left + limit);
        if (i > start && breaks[i]) lastBreak = i;
        const next = x + u.advance;
        // `hangingPunct`: a comma or full stop that starts inside the line hangs past its end.
        const hangs =
          p.hangingPunct && i > start && isHangingPunctuation(u.code) && x - left <= limit + 1;
        if (wrap && next - left > limit + 1 && !isTrailingSpace(u.code) && !hangs) {
          if (lastBreak > start) end = lastBreak;
          else end = forcedLineEnd(paragraphText, breaks, start, i);
          break;
        }
        x = next;
        end = i + 1;
      }
      lineRanges.push({ start, end, forced });
      start = end;
    }
    // A forced break at the end of a paragraph starts a new, empty line: the caret can reach
    // it, and anchoring and `spAutoFit` count its height (sized by the paragraph mark).
    const lastRange = lineRanges[lineRanges.length - 1];
    if (!lastRange || (lastRange.forced && lastRange.end === units.length))
      lineRanges.push({ start: units.length, end: units.length, forced: false });

    lineRanges.forEach((range, li) => {
      // A fit check reads only the content height: past the box, the rest cannot make it fit.
      if (y > overflowAt) return;
      const first = li === 0;
      let maxSize = 0;
      let maxFace: ResolvedFace | null = null;
      let maxRun = -2;
      // A segment's units are contiguous: visit each segment on the line once, in order.
      for (let i = range.start, last = -1; i < range.end; i += 1) {
        if (units[i].seg === last) continue;
        last = units[i].seg;
        const s = segments[last];
        if (s.sizePt >= maxSize) {
          maxSize = s.sizePt;
          maxFace = s.face;
          maxRun = s.runIndex;
        }
      }
      // One run in several fonts (its East Asian and Latin text, `font-slots.ts`): the tallest
      // ascent and the deepest descent of its faces place the baseline (reference PDFs of
      // `east-asian-line-breaking`, slide 4, and `east-asian-theme-fonts`).
      let metrics = maxFace?.metrics;
      if (metrics)
        for (let i = range.start, last = -1; i < range.end; i += 1) {
          if (units[i].seg === last) continue;
          last = units[i].seg;
          const s = segments[last];
          if (s.runIndex !== maxRun || s.sizePt !== maxSize || s.face === maxFace) continue;
          metrics = {
            ...metrics,
            ascent: Math.max(metrics.ascent, s.face.metrics.ascent),
            descent: Math.max(metrics.descent, s.face.metrics.descent),
          };
        }
      if (maxSize === 0 || !maxFace) {
        // An empty line: an empty paragraph, or the line after a trailing break. The paragraph
        // mark (`a:endParaRPr`) sizes it.
        const endSize = autofitSizePt(p.endProps.size / 100, fontScale);
        const afterBreak = range.start > 0 || segments.length === 0;
        maxSize = afterBreak ? endSize : segments[segments.length - 1].sizePt;
        maxFace = afterBreak ? faceFor(p.endProps, measurer) : segments[segments.length - 1].face;
      }
      // A bullet larger than the text does not make its line taller (fixture
      // `bullets-numbering`, "Custom character, 150%").
      const lineMetrics = metrics ?? maxFace.metrics;
      const single = singleLinePt(maxSize, lineMetrics, basis);
      const heightPt = lineHeightInPt(single, p.lineSpacing, reduction);
      let height = heightPt * EMU_PER_PT;
      let baseline = baselineInLinePt(single, heightPt, lineMetrics) * EMU_PER_PT;
      // An equation that needs more room than the line has makes the line taller. Above single
      // percent spacing (`a:spcPct`), the equation grows the single line, and the spacing then
      // adds its share of the text band on top, as DOCX Editor's `equationSpacingBand` does for
      // its `auto` rule: the line keeps the equation's own extent. The extra goes above and
      // below the baseline as on a text line. At 150 % (`equation-lines`, slide 1), the first
      // line of stacked display fractions comes 8.2 pt above the reference and the line pitch
      // 3 pt over it, against 14.2 pt and 5 pt when the equation fills the extra space. At
      // single spacing and below, and at exact spacing (`a:spcPts`), the equation fills the
      // line's own space first (the reference PDFs of `equations` and `preserved-content`).
      const raised = 'pct' in p.lineSpacing && heightPt > single + 1e-9;
      const singleBase = baselineInLinePt(single, single, lineMetrics) * EMU_PER_PT;
      let growAbove = 0;
      let growBelow = 0;
      for (let i = range.start, last = -1; i < range.end; i += 1) {
        if (units[i].seg === last) continue;
        last = units[i].seg;
        const box = segments[last].math;
        if (!box) continue;
        if (raised) {
          growAbove = Math.max(growAbove, box.lineAscent - singleBase);
          growBelow = Math.max(growBelow, box.lineDescent - (single * EMU_PER_PT - singleBase));
          continue;
        }
        const above = Math.max(0, box.lineAscent - baseline);
        const below = Math.max(0, box.lineDescent - (height - baseline));
        height += above + below;
        baseline += above;
      }
      height += growAbove + growBelow;
      baseline += growAbove;

      // Positions, then alignment. Positions include the left inset, so no later pass copies
      // the position arrays to add it.
      const left = lineLeft(first);
      let contentEnd = range.end;
      while (
        contentEnd > range.start &&
        (isTrailingSpace(units[contentEnd - 1].code) || units[contentEnd - 1].code === 10)
      )
        contentEnd -= 1;
      let width = 0;
      for (let i = range.start; i < contentEnd; i += 1) width += units[i].advance;
      // A bidi line in visual order (UAX #9 L1/L2); its visible extent is what aligns.
      const segOf = (i: number): number => units[i].seg;
      const placeLine = (advance: (i: number) => number): BidiLine | undefined =>
        bidi && placeBidiLine(bidi, codes, range.start, range.end, contentEnd, advance, segOf);
      let bidiLine = placeLine((i) => units[i].advance);
      if (bidiLine) width = bidiLine.contentRight - bidiLine.contentLeft;
      // A hanging comma or full stop stays outside the line for alignment.
      const hanging =
        p.hangingPunct &&
        contentEnd > range.start &&
        isHangingPunctuation(units[contentEnd - 1].code) &&
        width > avail(first) + 1;
      if (hanging) width -= units[contentEnd - 1].advance;
      const room = avail(first) - width;
      const lastOfPara = li === lineRanges.length - 1 || range.forced;
      let shift = 0;
      let extra: Float64Array | undefined;
      // `shift` is measured from the paragraph's leading edge: the right edge when it is `rtl`.
      const alignment = paragraphAlignment(p.align);
      const justifies = p.align === 'just' || p.align === 'justLow';
      const distributes = p.align === 'dist' || p.align === 'thaiDist';
      if (alignment !== 'both') shift = leadingOffset(alignment, p.rtl, room);
      // A justified line that ends with hanging punctuation fills the line, also the last line.
      else if (((justifies && (!lastOfPara || hanging)) || distributes) && room > 0)
        // The hanging character keeps its gap: the line ends where the character starts.
        extra = justifyExtra(units, segments, range.start, contentEnd, room);
      if (!wrap && p.align === 'ctr') shift = (avail(first) - width) / 2;
      let caretX: number[] = [];
      let lineBidi: LineBidi | undefined;
      const origin = (p.rtl ? colW - left - shift : left + shift) + bp.lIns;
      let x = origin;
      const runs: GlyphRun[] = [];
      const marks: LineMark[] = [];
      let cur: { seg: number; text: string; xs: number[]; start: number } | null = null;
      const flush = (): void => {
        if (!cur) return;
        const s = segments[cur.seg];
        if (s.math) {
          placeMath(s.math, cur.xs[0], s.runIndex, cur.start, runs, marks);
          cur = null;
          return;
        }
        if (!s.isBreak && cur.text.length > 0) {
          cur.xs.push(x);
          runs.push({
            kind: 'text',
            text: cur.text,
            x: cur.xs,
            face: s.face,
            sizePt: s.sizePt,
            rise: s.rise,
            props: s.props,
            runIndex: s.runIndex,
            start: cur.start,
          });
        }
        cur = null;
      };
      // Display text keeps source length (caps never change length), so unit index = offset.
      const sourceOffset = range.start;
      const advance = (i: number): number => units[i].advance + (extra?.[i - range.start] ?? 0);
      if (bidiLine && extra) bidiLine = placeLine(advance);
      if (bidiLine) {
        // Justification widened the line: its right edge moved with it.
        const shown = bidiLine.contentRight - bidiLine.contentLeft;
        const lineLeftX = (p.rtl ? colW - left - shift - shown : left + shift) + bp.lIns;
        placeBidiRuns(bidiLine, lineLeftX, range.start, units, segments, runs, marks);
        caretX = bidiCaretX(bidiLine, range.start, range.end - range.start, advance).map(
          (v) => v + lineLeftX
        );
        lineBidi = lineBidiOf(bidiLine, lineLeftX, range.start, advance, p.rtl);
      }
      // A bidi line placed its runs above; a left-to-right line places them here.
      const lineEnd = bidiLine ? range.start : range.end;
      for (let i = range.start; i < lineEnd; i += 1) {
        const u = units[i];
        if (!cur || cur.seg !== u.seg) {
          flush();
          cur = { seg: u.seg, text: '', xs: [], start: i };
        }
        caretX.push(x);
        cur.xs.push(x);
        cur.text += String.fromCharCode(u.code);
        // `extra` ends at the content end: trailing spaces and breaks after it get no extra.
        x += u.advance + (extra?.[i - range.start] ?? 0);
      }
      flush();
      if (!bidiLine) caretX.push(x);
      const lineRuns =
        first && bullet
          ? [
              ...(p.rtl ? markerRuns(placeMarker(bullet, shift)) : [placeMarker(bullet, shift)]),
              ...runs,
            ]
          : runs;
      lines.push({
        paragraph: pi,
        top: y,
        baseline: y + baseline,
        height,
        column: 0,
        start: sourceOffset,
        end: sourceOffset + (range.end - range.start),
        caretX: range.end === range.start ? [origin] : caretX,
        ...(lineBidi ? { bidi: lineBidi } : {}),
        runs: paintPieces(lineRuns),
        fontHeight: maxSize * EMU_PER_PT,
        ...(first && pending
          ? { pendingBullet: wholeMarker(placeMarker(pending, shift), p.rtl) }
          : {}),
        ...(marks.length > 0 ? { marks } : {}),
      });
      y += height;
    });
    if (!isLast || bp.spcFirstLastPara)
      y += paragraphSpacingPt(p.spaceAfter, firstSize) * EMU_PER_PT;
    if (y > overflowAt) overflowed = true;
  });

  // Columns: flow lines into the next column when the box height is exceeded. Baselines are
  // placed from the top of their column, before anchoring (`placeBaseline`).
  let placed: Line[];
  let contentHeight = y;
  // What autofit reads: with columns, the bottom of the lowest line, so an overflowing last
  // column counts.
  let needed = y;
  if (columns <= 1 || innerH <= 0)
    placed = lines.map((line) => ({ ...line, baseline: placeBaseline(line.baseline) }));
  else {
    placed = [];
    let col = 0;
    let offset = 0;
    for (const line of lines) {
      if (line.top - offset + line.height > innerH && col < columns - 1 && line.top - offset > 0) {
        col += 1;
        offset = line.top;
      }
      const dx = col * (colW + bp.spcCol);
      placed.push({
        ...line,
        column: col,
        top: line.top - offset,
        baseline: placeBaseline(line.baseline - offset),
        caretX: line.caretX.map((v) => v + dx),
        ...(line.bidi ? { bidi: shiftLineBidi(line.bidi, dx) } : {}),
        runs: line.runs.map((r) => ({ ...r, x: r.x.map((v) => v + dx) })),
        ...(line.marks ? { marks: line.marks.map((m) => moveMark(m, dx)) } : {}),
      });
    }
    contentHeight = Math.min(y, innerH);
    needed = contentHeight;
    for (const l of placed) needed = Math.max(needed, l.top + l.height);
  }

  // Vertical anchoring moves lines down; their position arrays stay as they are.
  const dy = bp.tIns + anchorOffset(bp.anchor, innerH, contentHeight);
  const finalLines = dy === 0 ? placed : shiftLines(placed, dy);
  const frame = {
    x: rect.x + (rect.width - outerW) / 2,
    y: rect.y + (rect.height - outerH) / 2,
    width: outerW,
    height: outerH,
  };
  return {
    box: frame,
    rotation,
    upright: vert === 'eaVert' ? 'eastAsian' : vert === 'wordArtVert' ? 'all' : 'none',
    lines: finalLines,
    contentHeight: needed + bp.tIns + bp.bIns,
    paragraphLengths,
    substitutions: [...substitutions],
  };
}
