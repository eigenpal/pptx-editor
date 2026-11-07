// Office Math layout (ECMA-376 Part 1, 22.1): one math zone into a box of positioned glyphs and
// filled marks (rules, the radical sign, stretched delimiters, braces). DOM-free; the measurer
// is injected, as for text. All lengths are EMU, with y up from the math baseline.
//
// The rules follow TeX's math layout: atoms with classes and the spaces between them, script
// sizes (73 % and 60 %, rounded to half a point), cramped styles under radicals and in
// subscripts, and shifts that keep a minimum gap. The constants are measured on the reference
// PDF of the `equations` fixture (STIX Two Math, 18 pt). Where a run's font is not available,
// the runs draw in STIX Two Math or Latin Modern Math when the host has them, else in the serif
// text face, with italic letters instead of math alphanumerics. The known approximation
// `equations` lists what differs.

import type { PathCommand } from '@eigenpal/ooxml-core';
import type { MathNode, MathRow, MathRun, MathStyle, MathZone } from '../../cascade/math.ts';
import type { RunProps } from '../../cascade/types.ts';
import { autofitSizePt } from './line-metrics.ts';
import {
  atomSpace,
  charClass,
  INTEGRALS,
  isMathLetter,
  mathAlphanumeric,
  plainLetter,
  spacingAccent,
  type AtomClass,
} from './math-glyphs.ts';
import {
  arrow,
  delimiter,
  delimiterWidth,
  horizontalGroup,
  moveCommands,
  radicalSign,
  rect,
  stroke,
} from './math-marks.ts';
import type { ResolvedFace, TextMeasurer } from './measurer.ts';

const EMU_PER_PT = 12_700;

/** Constants in em of the construct's size (measured, see the header). */
const C = {
  axis: 0.29,
  rule: 0.055,
  numDisplay: 0.76,
  denDisplay: 0.667,
  numText: 0.52,
  denText: 0.453,
  gapDisplay: 0.165,
  gapText: 0.055,
  sup: 0.36,
  supCramped: 0.293,
  sub: 0.2,
  subWithSup: 0.24,
  supWithSub: 0.4,
  preSub: 0.24,
  preSup: 0.4,
  scriptSpace: 0.03,
  radWidth: 0.655,
  radGap: 0.178,
  radBottom: 0.087,
  radMinContent: 0.66,
  delimMin: 0.927,
  delimFactor: 0.9,
  delimPad: 0.02,
  sepPad: 0.11,
  matrixRow: 1.16,
  matrixCol: 1,
  eqArrRow: 1.187,
  limUpper: 0.72,
  limLower: 0.613,
  limGap: 0.11,
  limGapLow: 0.167,
  barGap: 0.11,
  barGapBottom: 0.167,
  accentGap: 0.04,
  groupGap: 0.1,
  boxPad: 0.05,
  boxMinAscent: 0.61,
  boxMinDescent: 0.17,
} as const;

/** The scale of each script level (73 % and 60 % of the base size). */
const LEVEL_SCALE = [1, 0.73, 0.6] as const;

/** Math fonts that draw math runs whose own font is missing, in order. */
const MATH_FAMILIES = ['STIX Two Math', 'Latin Modern Math'] as const;
/** The text face that draws math runs when no math font is available. */
const SERIF_FAMILY = 'Liberation Serif';

/** Positioned glyphs of one math run piece. */
export interface MathGlyphs {
  readonly text: string;
  /** The left edge of each UTF-16 unit, EMU from the box origin; length `text.length + 1`. */
  readonly x: readonly number[];
  /** The baseline of the glyphs, EMU above the math baseline. */
  readonly y: number;
  readonly face: ResolvedFace;
  readonly sizePt: number;
  readonly props: RunProps;
}

/** A filled mark: a rule, the radical sign, a delimiter, a brace, or an arrow. */
export interface MathMark {
  /** The outline, EMU, x from the box origin and y up from the math baseline. Fill nonzero. */
  readonly commands: readonly PathCommand[];
  readonly props: RunProps;
}

/** A laid-out math zone. */
export interface MathBox {
  readonly width: number;
  readonly ascent: number;
  readonly descent: number;
  readonly glyphs: readonly MathGlyphs[];
  readonly marks: readonly MathMark[];
  /** The face and size of the zone's base runs: they size the line. */
  readonly face: ResolvedFace;
  readonly sizePt: number;
  /**
   * The room the equation needs above and below the baseline in a line: the ink of its glyphs
   * and its marks. A line grows to hold it.
   */
  readonly lineAscent: number;
  readonly lineDescent: number;
}

interface Box {
  w: number;
  asc: number;
  desc: number;
  glyphs: MathGlyphs[];
  marks: MathMark[];
}

interface Item {
  box: Box;
  cls: AtomClass;
}

interface Env {
  readonly measurer: TextMeasurer;
  readonly fontScale: number;
  readonly level: 0 | 1 | 2;
  readonly display: boolean;
  readonly cramped: boolean;
  /** The style of runs that state none: plain in a function name. */
  readonly sty?: MathStyle;
  /** The run properties of the zone, for marks of constructs with none. */
  readonly props: RunProps;
}

const emptyBox = (): Box => ({ w: 0, asc: 0, desc: 0, glyphs: [], marks: [] });

/** The largest of `values`, and at least 0. A loop: the values come from the file. */
function maxOf(values: readonly number[]): number {
  let max = 0;
  for (const v of values) if (v > max) max = v;
  return max;
}

/** `box` moved by `dx`, `dy`, into `into`. */
function place(into: Box, box: Box, dx: number, dy: number): void {
  for (const g of box.glyphs) into.glyphs.push({ ...g, x: g.x.map((v) => v + dx), y: g.y + dy });
  for (const m of box.marks) into.marks.push({ ...m, commands: moveCommands(m.commands, dx, dy) });
  into.asc = Math.max(into.asc, box.asc + dy);
  into.desc = Math.max(into.desc, box.desc - dy);
}

const nextLevel = (level: 0 | 1 | 2): 0 | 1 | 2 => (level === 0 ? 1 : 2);

/** The size in points of `props` at a script level. */
function sizePt(props: RunProps, env: Env, level = env.level): number {
  const base = autofitSizePt(props.size / 100, env.fontScale);
  if (level === 0) return base;
  return Math.max(1, Math.round(base * LEVEL_SCALE[level] * 2) / 2);
}

const emOf = (props: RunProps, env: Env, level = env.level): number =>
  sizePt(props, env, level) * EMU_PER_PT;

/** The face that draws a math run, and whether it is a math font (math alphanumerics). */
interface FaceChoice {
  readonly face: ResolvedFace;
  readonly math: boolean;
}

const faceCache = new WeakMap<TextMeasurer, Map<string, FaceChoice>>();

function mathFace(
  measurer: TextMeasurer,
  family: string,
  bold: boolean,
  italic: boolean
): FaceChoice {
  let cache = faceCache.get(measurer);
  if (!cache) faceCache.set(measurer, (cache = new Map()));
  const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}`;
  const hit = cache.get(key);
  if (hit) return hit;
  let choice: FaceChoice | undefined;
  for (const f of [family, ...MATH_FAMILIES]) {
    const face = measurer.resolve({ family: f, bold: false, italic: false });
    if (!face.substituted && /math/i.test(face.family)) {
      choice = { face, math: true };
      break;
    }
  }
  if (!choice) {
    const own = measurer.resolve({ family, bold, italic });
    choice = {
      face: own.substituted ? measurer.resolve({ family: SERIF_FAMILY, bold, italic }) : own,
      math: false,
    };
  }
  if (cache.size < 256) cache.set(key, choice);
  return choice;
}

/** Ink extents in em: [lowest, highest], y up. */
const inkCache = new WeakMap<TextMeasurer, Map<string, readonly [number, number]>>();
const DESCENDERS = /[gjpqyβγζημξρφχψ,;]/u;
const SHORT = /[a-zα-ω=+−\-×·∘<>~*]/u;

function inkOf(measurer: TextMeasurer, face: ResolvedFace, cp: number): readonly [number, number] {
  let cache = inkCache.get(measurer);
  if (!cache) inkCache.set(measurer, (cache = new Map()));
  const key = `${face.id}\u0000${cp}`;
  const hit = cache.get(key);
  if (hit) return hit;
  let out: readonly [number, number] | undefined;
  const outline = measurer.glyphOutline?.(face, cp);
  if (outline && outline.length > 0) {
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
    if (Number.isFinite(lo)) out = [lo, hi];
  }
  if (!out) {
    // An estimate when the measurer has no outline: by the plain character's shape, and at
    // least the face's ascent and descent. The line grows by ink only (`layoutMath`), so a glyph
    // with no outline (a fallback face draws it, an East Asian character for example) must not
    // pass the line's top or bottom.
    const ch = String.fromCodePoint(cp).normalize('NFKD').charAt(0);
    const { ascent, descent } = face.metrics;
    if (cp === 32) out = [0, 0];
    else if (DESCENDERS.test(ch)) out = [Math.min(-0.21, -descent), Math.max(0.46, ascent)];
    else if (SHORT.test(ch)) out = [Math.min(0, -descent), Math.max(0.46, ascent)];
    else out = [Math.min(0, -descent), Math.max(0.68, ascent)];
  }
  if (cache.size < 4096) cache.set(key, out);
  return out;
}

/** One piece of text in one face as a box. */
function textBox(text: string, face: ResolvedFace, size: number, props: RunProps, env: Env): Box {
  const em = size * EMU_PER_PT;
  const adv = env.measurer.advances(face, text);
  const x = [0];
  for (let i = 0; i < adv.length; i += 1) x.push(x[i]! + adv[i]! * em);
  let lo = 0;
  let hi = 0;
  for (const ch of text) {
    const [a, b] = inkOf(env.measurer, face, ch.codePointAt(0) ?? 32);
    lo = Math.min(lo, a);
    hi = Math.max(hi, b);
  }
  // A character outside the BMP (a math alphanumeric) is a piece of its own: an SVG `x` list
  // positions each character, not each UTF-16 unit.
  const glyphs: MathGlyphs[] = [];
  for (let i = 0; i < text.length;) {
    const astral = (text.codePointAt(i) ?? 0) > 0xffff;
    let j = i + (astral ? 2 : 1);
    if (!astral) while (j < text.length && (text.codePointAt(j) ?? 0) <= 0xffff) j += 1;
    glyphs.push({ text: text.slice(i, j), x: x.slice(i, j + 1), y: 0, face, sizePt: size, props });
    i = j;
  }
  return { w: x[x.length - 1]!, asc: hi * em, desc: -lo * em, glyphs, marks: [] };
}

/** Run properties for glyphs and marks: no text decoration, which math does not draw. */
const plainProps = (props: RunProps): RunProps =>
  props.underline === 'none' && props.strike === 'noStrike'
    ? props
    : { ...props, underline: 'none', strike: 'noStrike' };

/** The atoms of a math run: letters and digits together, each operator alone. */
function runItems(run: MathRun, env: Env): Item[] {
  const props = plainProps(run.props);
  // Normal text (`m:nor`) keeps the run's own size at every script level: the script sizes
  // apply to math text only. Its place (the shift of a superscript, for example) still follows
  // the construct. The rule comes from the reference PDFs, not from the standard's text: the
  // `equation-lines` fixture, slide 2, draws normal text at the run's size in a subscript and a
  // superscript of a math base, in a nested script, and in a radical degree, in STIX Two Math
  // and in Liberation Serif, as `preserved-content`, slide 1, does.
  const size = sizePt(props, env, run.normal ? 0 : env.level);
  if (run.normal) {
    const face = env.measurer.resolve({
      family: props.latin,
      bold: props.bold,
      italic: props.italic,
    });
    return run.text ? [{ box: textBox(run.text, face, size, props, env), cls: 'ord' }] : [];
  }
  const style: MathStyle = run.sty ?? env.sty ?? 'i';
  const bold = style === 'b' || style === 'bi';
  const italic = style === 'i' || style === 'bi';
  const upright = mathFace(env.measurer, props.latin, bold, false);
  const letters = upright.math ? upright : mathFace(env.measurer, props.latin, bold, italic);
  const items: Item[] = [];
  let pending: { text: string; face: ResolvedFace } | null = null;
  const flush = (): void => {
    if (!pending) return;
    items.push({ box: textBox(pending.text, pending.face, size, props, env), cls: 'ord' });
    pending = null;
  };
  // Faces of the plain letters that stand for math alphanumerics, when no math font draws.
  const styled = new Map<MathStyle, ResolvedFace>();
  const plainFace = (s: MathStyle): ResolvedFace => {
    let face = styled.get(s);
    if (!face) {
      const b = s === 'b' || s === 'bi';
      face = mathFace(env.measurer, props.latin, b, s === 'i' || s === 'bi').face;
      styled.set(s, face);
    }
    return face;
  };
  for (const original of run.text) {
    let ch = original;
    let cp = ch.codePointAt(0) ?? 32;
    // Spaces in a math run do not draw: math spacing comes from the atom classes.
    if (cp === 32) continue;
    let own: ResolvedFace | undefined;
    const plain = upright.math ? undefined : plainLetter(cp);
    if (plain) {
      cp = plain.cp;
      ch = String.fromCodePoint(cp);
      own = plainFace(plain.style);
    }
    const shown = cp === 0x2d ? '−' : ch;
    const cls = charClass(shown);
    const letter = isMathLetter(cp) && !own;
    const face = own ?? (letter ? letters.face : upright.face);
    const drawn =
      letter && upright.math ? String.fromCodePoint(mathAlphanumeric(cp, style)) : shown;
    if (cls !== 'ord') {
      flush();
      items.push({ box: textBox(drawn, face, size, props, env), cls });
      continue;
    }
    if (pending && pending.face !== face) flush();
    if (pending) pending.text += drawn;
    else pending = { text: drawn, face };
  }
  flush();
  return items;
}

/** A row: atoms side by side with the spaces of their classes. */
function rowBox(row: MathRow, env: Env): Box {
  const items: Item[] = [];
  for (const node of row) for (const item of nodeItems(node, env)) items.push(item);
  // A binary operator with no left operand, or before a relation or a closing, is ordinary.
  for (let i = 0; i < items.length; i += 1) {
    if (items[i]!.cls !== 'bin') continue;
    const prev = items[i - 1]?.cls;
    const next = items[i + 1]?.cls;
    if (
      !prev ||
      prev === 'bin' ||
      prev === 'op' ||
      prev === 'rel' ||
      prev === 'open' ||
      prev === 'punct'
    )
      items[i]!.cls = 'ord';
    else if (!next || next === 'rel' || next === 'close' || next === 'punct') items[i]!.cls = 'ord';
  }
  const out = emptyBox();
  const mu = emOf(env.props, env) / 18;
  let x = 0;
  items.forEach((item, i) => {
    if (i > 0) x += atomSpace(items[i - 1]!.cls, item.cls, env.level > 0) * mu;
    place(out, item.box, x, 0);
    x += item.box.w;
  });
  out.w = x;
  return out;
}

const childEnv = (env: Env, over: Partial<Env> = {}): Env => ({
  ...env,
  display: false,
  sty: undefined,
  ...over,
});

/** Boxes of a row stacked on one another, centered, `gap` em apart at least. */
function stackRows(rows: readonly Box[], rowSpace: number, em: number, axis: number): Box {
  const out = emptyBox();
  if (rows.length === 0) return out;
  const width = maxOf(rows.map((r) => r.w));
  let y = 0;
  const ys: number[] = [];
  rows.forEach((r, i) => {
    if (i > 0) y -= Math.max(rowSpace * em, rows[i - 1]!.desc + r.asc + 0.1 * em);
    ys.push(y);
  });
  const top = rows[0]!.asc;
  const bottom = y - rows[rows.length - 1]!.desc;
  const dy = axis - (top + bottom) / 2;
  rows.forEach((r, i) => place(out, r, (width - r.w) / 2, ys[i]! + dy));
  out.w = width;
  return out;
}

function fraction(node: Extract<MathNode, { kind: 'frac' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const t = C.rule * em;
  const a = C.axis * em;
  const out = emptyBox();
  if (node.type === 'lin' || node.type === 'skw') {
    const level = node.type === 'skw' ? nextLevel(env.level) : env.level;
    const n = rowBox(node.num, childEnv(env, { level }));
    const d = rowBox(node.den, childEnv(env, { level, cramped: true }));
    if (node.type === 'lin') {
      const { face } = mathFace(env.measurer, props.latin, false, false);
      const slash = textBox('/', face, sizePt(props, env), props, env);
      place(out, n, 0, 0);
      place(out, slash, n.w, 0);
      place(out, d, n.w + slash.w, 0);
      out.w = n.w + slash.w + d.w;
      return out;
    }
    // Skewed: the numerator raised, a slash, the denominator lowered.
    const gap = 0.12 * em;
    place(out, n, 0, 0.35 * em);
    const sx = n.w + gap * 0.5;
    out.marks.push({ props, commands: stroke(sx, -0.25 * em, sx + 0.45 * em, 0.85 * em, t) });
    place(out, d, sx + 0.45 * em + gap * 0.5, -0.15 * em);
    out.w = sx + 0.45 * em + gap + d.w;
    return out;
  }
  const n = rowBox(node.num, childEnv(env));
  const d = rowBox(node.den, childEnv(env, { cramped: true }));
  const w = Math.max(n.w, d.w);
  // A stack with no bar (`noBar`) keeps the places of a fraction with one.
  const gap = (env.display ? C.gapDisplay : C.gapText) * em;
  const u = Math.max((env.display ? C.numDisplay : C.numText) * em, a + t / 2 + gap + n.desc);
  const v = Math.max((env.display ? C.denDisplay : C.denText) * em, d.asc + gap + t / 2 - a);
  if (node.type === 'bar') out.marks.push({ props, commands: rect(0, a - t / 2, w, t) });
  place(out, n, (w - n.w) / 2, u);
  place(out, d, (w - d.w) / 2, -v);
  out.w = w;
  return out;
}

function scripts(node: Extract<MathNode, { kind: 'script' }>, env: Env): Box {
  const em = emOf(plainProps(node.props), env);
  const base = rowBox(node.base, env);
  const level = nextLevel(env.level);
  const sup = node.sup ? rowBox(node.sup, childEnv(env, { level })) : undefined;
  const sub = node.sub ? rowBox(node.sub, childEnv(env, { level, cramped: true })) : undefined;
  let u = 0;
  let v = 0;
  if (sup && sub) {
    u = Math.max(C.supWithSub * em, base.asc - 0.35 * em);
    v = Math.max(C.subWithSup * em, base.desc + 0.05 * em);
    const gap = u - sup.desc - (sub.asc - v);
    if (gap < 4 * C.rule * em) v += 4 * C.rule * em - gap;
  } else if (sup) {
    u = Math.max(
      (env.cramped ? C.supCramped : C.sup) * em,
      base.asc - 0.3 * em,
      sup.desc + 0.1 * em
    );
  } else if (sub) {
    v = Math.max(C.sub * em, base.desc + 0.05 * em, sub.asc - 0.36 * em);
  }
  const out = emptyBox();
  const scriptW = Math.max(sup?.w ?? 0, sub?.w ?? 0);
  if (node.type === 'pre') {
    u = Math.max(C.preSup * em, u);
    v = Math.max(C.preSub * em, sub && sup ? v * 0.7 : v);
    if (sup) place(out, sup, scriptW - sup.w, u);
    if (sub) place(out, sub, scriptW - sub.w, -v);
    place(out, base, scriptW, 0);
    out.w = scriptW + base.w;
    return out;
  }
  place(out, base, 0, 0);
  if (sup) place(out, sup, base.w, u);
  if (sub) place(out, sub, base.w, -v);
  out.w = base.w + scriptW + C.scriptSpace * em;
  return out;
}

function radical(node: Extract<MathNode, { kind: 'rad' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const t = C.rule * em;
  const content = rowBox(node.base, childEnv(env, { cramped: true }));
  const top = Math.max(content.asc, C.radMinContent * em) + C.radGap * em + t;
  const bottom = Math.min(-C.radBottom * em, -(content.desc + 0.06 * em));
  const signW = C.radWidth * em + Math.max(0, top - bottom - 1.1 * em) * 0.08;
  const out = emptyBox();
  let signX = 0;
  if (node.degree) {
    const degree = rowBox(node.degree, childEnv(env, { level: 2 }));
    signX = Math.max(0, degree.w - 0.4 * signW);
    place(out, degree, Math.max(0, 0.4 * signW - degree.w) / 2, bottom + (top - bottom) * 0.6);
  }
  out.marks.push({ props, commands: radicalSign(signX, bottom, top, signW, em) });
  const cx = signX + signW - 0.06 * em;
  place(out, content, cx, 0);
  const end = cx + content.w + 0.02 * em;
  out.marks.push({ props, commands: rect(signX + signW, top - t, end - signX - signW, t) });
  out.asc = Math.max(out.asc, top);
  out.desc = Math.max(out.desc, -bottom);
  out.w = end + 0.04 * em;
  return out;
}

/** A delimiter character at `height`, from `x`: a stretched mark, else the glyph itself. */
function delimiterAt(
  ch: string,
  x: number,
  bottom: number,
  top: number,
  props: RunProps,
  env: Env,
  out: Box
): number {
  if (!ch) return 0;
  const em = emOf(props, env);
  const w = delimiterWidth(ch, em, top - bottom);
  if (w !== undefined) {
    out.marks.push({ props, commands: delimiter(ch, x, bottom, top, w, em) });
    out.asc = Math.max(out.asc, top);
    out.desc = Math.max(out.desc, -bottom);
    return w;
  }
  const { face } = mathFace(env.measurer, props.latin, false, false);
  const glyph = textBox(ch, face, sizePt(props, env), props, env);
  place(out, glyph, x, 0);
  return glyph.w;
}

function delimited(node: Extract<MathNode, { kind: 'delim' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const a = C.axis * em;
  // The items keep the display style of the row (reference PDF: a fraction in brackets).
  const items = node.items.map((r) => rowBox(r, { ...env, sty: undefined }));
  const asc = maxOf(items.map((b) => b.asc));
  const desc = maxOf(items.map((b) => b.desc));
  const half = Math.max(asc - a, desc + a);
  const h = node.grow ? Math.max(C.delimMin * em, 2 * half * C.delimFactor) : C.delimMin * em;
  const bottom = a - h / 2;
  const top = a + h / 2;
  const out = emptyBox();
  const pad = C.delimPad * em;
  let x = 0;
  const begin = delimiterAt(node.begin, x, bottom, top, props, env, out);
  x += begin > 0 ? begin + pad : 0;
  items.forEach((item, i) => {
    if (i > 0 && node.separator) {
      x += C.sepPad * em;
      x += delimiterAt(node.separator, x, bottom, top, props, env, out);
      x += C.sepPad * em;
    }
    place(out, item, x, 0);
    x += item.w;
  });
  const end = delimiterAt(node.end, x + pad, bottom, top, props, env, out);
  x += end > 0 ? end + pad : 0;
  out.w = x;
  return out;
}

function matrix(node: Extract<MathNode, { kind: 'matrix' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const cells = node.rows.map((row) => row.map((c) => rowBox(c, childEnv(env))));
  const cols = maxOf(cells.map((r) => r.length));
  const colW = Array.from({ length: cols }, (_, j) => maxOf(cells.map((r) => r[j]?.w ?? 0)));
  const rows: Box[] = cells.map((row) => {
    const box = emptyBox();
    let x = 0;
    for (let j = 0; j < cols; j += 1) {
      const cell = row[j];
      if (cell) place(box, cell, x + (colW[j]! - cell.w) / 2, 0);
      x += colW[j]! + (j < cols - 1 ? C.matrixCol * em : 0);
    }
    box.w = x;
    return box;
  });
  return stackRows(rows, C.matrixRow, em, C.axis * em);
}

function accent(node: Extract<MathNode, { kind: 'acc' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const base = rowBox(node.base, childEnv(env, { cramped: true }));
  const out = emptyBox();
  place(out, base, 0, 0);
  const top = Math.max(base.asc, 0.46 * em) + C.accentGap * em;
  const chr = node.chr;
  const glyph = spacingAccent(chr);
  if (glyph === undefined) {
    if (chr === '⃗' || chr === '⃖' || chr === '⃡') {
      const y = top + 0.12 * em;
      out.marks.push({ props, commands: arrow(0, Math.max(base.w, 0.4 * em), y, em, chr === '⃖') });
      out.asc = Math.max(out.asc, y + 0.12 * em);
    } else {
      const t = C.rule * em;
      out.marks.push({ props, commands: rect(0, top + 0.06 * em, base.w, t) });
      out.asc = Math.max(out.asc, top + 0.06 * em + t);
    }
    out.w = base.w;
    return out;
  }
  const { face } = mathFace(env.measurer, props.latin, false, false);
  const mark = textBox(glyph, face, sizePt(props, env), props, env);
  const [lo] = inkOf(env.measurer, face, glyph.codePointAt(0) ?? 32);
  // Italic bases lean right: the accent sits a little right of the center.
  place(out, mark, (base.w - mark.w) / 2 + 0.04 * em, top - lo * emOf(props, env));
  out.w = Math.max(base.w, mark.w);
  return out;
}

function overUnder(base: Box, limit: Box, pos: 'low' | 'upp', em: number): Box {
  const out = emptyBox();
  const w = Math.max(base.w, limit.w);
  place(out, base, (w - base.w) / 2, 0);
  const y =
    pos === 'low'
      ? -Math.max(C.limLower * em, base.desc + C.limGapLow * em + limit.asc)
      : Math.max(C.limUpper * em, base.asc + C.limGap * em + limit.desc);
  place(out, limit, (w - limit.w) / 2, y);
  out.w = w;
  return out;
}

function nary(node: Extract<MathNode, { kind: 'nary' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const integral = INTEGRALS.has(node.chr);
  const scale = env.display ? (integral ? 1.6 : 1.35) : 1;
  const { face } = mathFace(env.measurer, props.latin, false, false);
  const size = sizePt(props, env) * scale;
  const op = textBox(node.chr, face, size, props, env);
  // The operator centers on the math axis.
  const dy = C.axis * em - (op.asc - op.desc) / 2;
  const sym = emptyBox();
  place(sym, op, 0, dy);
  sym.w = op.w;
  const level = nextLevel(env.level);
  const sub = node.sub ? rowBox(node.sub, childEnv(env, { level, cramped: true })) : undefined;
  const sup = node.sup ? rowBox(node.sup, childEnv(env, { level })) : undefined;
  const out = emptyBox();
  const underOver = node.limLoc ? node.limLoc === 'undOvr' : !integral;
  if (underOver) {
    const w = Math.max(sym.w, sub?.w ?? 0, sup?.w ?? 0);
    place(out, sym, (w - sym.w) / 2, 0);
    const gap = 0.1 * em;
    if (sup) place(out, sup, (w - sup.w) / 2, sym.asc + gap + sup.desc);
    if (sub) place(out, sub, (w - sub.w) / 2, -(sym.desc + gap + sub.asc));
    out.w = w;
  } else {
    place(out, sym, 0, 0);
    const x = sym.w + (integral ? -0.05 * em : 0.02 * em);
    if (sup) place(out, sup, x + (integral ? 0.15 * em : 0), sym.asc - sup.asc * 0.7);
    if (sub) place(out, sub, x, -sym.desc + sub.asc * 0.2);
    out.w = Math.max(sym.w, x + Math.max(sup?.w ?? 0, sub?.w ?? 0) + (integral ? 0.15 * em : 0));
  }
  const body = rowBox(node.base, env);
  const x = out.w + (3 / 18) * em;
  place(out, body, x, 0);
  out.w = x + body.w;
  return out;
}

function groupChr(node: Extract<MathNode, { kind: 'groupChr' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const base = rowBox(node.base, childEnv(env));
  const out = emptyBox();
  place(out, base, 0, 0);
  out.w = base.w;
  const below = node.pos === 'bot';
  const y = below ? -(base.desc + C.groupGap * em) : base.asc + C.groupGap * em;
  const ch = node.chr;
  if ('⏟⏞⎴⎵⏠⏡'.includes(ch)) {
    out.marks.push({ props, commands: horizontalGroup(ch, 0, base.w, y, em) });
    if (below) out.desc = Math.max(out.desc, -y + 0.25 * em);
    else out.asc = Math.max(out.asc, y + 0.25 * em);
  } else if (ch === '→' || ch === '←') {
    const ay = below ? y - 0.12 * em : y + 0.12 * em;
    out.marks.push({ props, commands: arrow(0, base.w, ay, em, ch === '←') });
    if (below) out.desc = Math.max(out.desc, -ay + 0.12 * em);
    else out.asc = Math.max(out.asc, ay + 0.12 * em);
  } else if (ch) {
    const { face } = mathFace(env.measurer, props.latin, false, false);
    const glyph = textBox(ch, face, sizePt(props, env), props, env);
    place(out, glyph, (base.w - glyph.w) / 2, below ? y - glyph.asc : y + glyph.desc);
    out.w = Math.max(base.w, glyph.w);
  }
  return out;
}

function borderBox(node: Extract<MathNode, { kind: 'borderBox' }>, env: Env): Box {
  const props = plainProps(node.props);
  const em = emOf(props, env);
  const t = C.rule * em;
  const pad = C.boxPad * em;
  const base = rowBox(node.base, childEnv(env));
  const asc = Math.max(base.asc, C.boxMinAscent * em) + pad;
  const desc = Math.max(base.desc, C.boxMinDescent * em) + pad;
  const out = emptyBox();
  const inner = t + pad;
  place(out, base, inner, 0);
  const w = base.w + 2 * inner;
  const { hide } = node;
  if (!hide.t) out.marks.push({ props, commands: rect(0, asc, w, t) });
  if (!hide.b) out.marks.push({ props, commands: rect(0, -desc - t, w, t) });
  if (!hide.l) out.marks.push({ props, commands: rect(0, -desc - t, t, asc + desc + 2 * t) });
  if (!hide.r) out.marks.push({ props, commands: rect(w - t, -desc - t, t, asc + desc + 2 * t) });
  out.asc = Math.max(out.asc, asc + t);
  out.desc = Math.max(out.desc, desc + t);
  out.w = w;
  return out;
}

/** The atoms of one construct. */
function nodeItems(node: MathNode, env: Env): Item[] {
  switch (node.kind) {
    case 'run':
      return runItems(node, env);
    case 'func': {
      const name = rowBox(node.name, { ...env, sty: 'p' });
      return [
        { box: name, cls: 'op' },
        { box: rowBox(node.base, { ...env, sty: undefined }), cls: 'ord' },
      ];
    }
    case 'frac':
      return [{ box: fraction(node, env), cls: 'ord' }];
    case 'script':
      return [{ box: scripts(node, env), cls: 'ord' }];
    case 'rad':
      return [{ box: radical(node, env), cls: 'ord' }];
    case 'nary':
      return [{ box: nary(node, env), cls: 'ord' }];
    case 'delim':
      return [{ box: delimited(node, env), cls: 'ord' }];
    case 'matrix':
      return [{ box: matrix(node, env), cls: 'ord' }];
    case 'acc':
      return [{ box: accent(node, env), cls: 'ord' }];
    case 'bar': {
      const props = plainProps(node.props);
      const em = emOf(props, env);
      const t = C.rule * em;
      const base = rowBox(node.base, childEnv(env, { cramped: node.pos === 'top' || env.cramped }));
      const out = emptyBox();
      place(out, base, 0, 0);
      const y =
        node.pos === 'top'
          ? Math.max(base.asc, 0.46 * em) + C.barGap * em
          : -(base.desc + C.barGapBottom * em) - t;
      out.marks.push({ props, commands: rect(0, y, base.w, t) });
      out.asc = Math.max(out.asc, y + t);
      out.desc = Math.max(out.desc, -y);
      out.w = base.w;
      return [{ box: out, cls: 'ord' }];
    }
    case 'limit': {
      const em = emOf(plainProps(node.props), env);
      const base = rowBox(node.base, childEnv(env));
      const limit = rowBox(
        node.limit,
        childEnv(env, { level: nextLevel(env.level), cramped: node.pos === 'low' })
      );
      // A limit under a function name (lim, max) keeps the name's operator spacing.
      return [{ box: overUnder(base, limit, node.pos, em), cls: 'op' }];
    }
    case 'groupChr':
      return [{ box: groupChr(node, env), cls: 'ord' }];
    case 'box':
      return [{ box: rowBox(node.base, env), cls: 'ord' }];
    case 'borderBox':
      return [{ box: borderBox(node, env), cls: 'ord' }];
    case 'eqArr': {
      const em = emOf(env.props, env);
      const rows = node.rows.map((r) => rowBox(r, childEnv(env)));
      return [{ box: stackRows(rows, C.eqArrRow, em, C.axis * em), cls: 'ord' }];
    }
    case 'phant': {
      const base = rowBox(node.base, env);
      const out: Box = node.show ? base : { ...base, glyphs: [], marks: [] };
      return [
        {
          box: {
            ...out,
            w: node.zeroWidth ? 0 : out.w,
            asc: node.zeroAscent ? 0 : out.asc,
            desc: node.zeroDescent ? 0 : out.desc,
          },
          cls: 'ord',
        },
      ];
    }
  }
}

const boxCache = new WeakMap<MathZone, WeakMap<TextMeasurer, Map<number, MathBox>>>();

/**
 * Lay out a math zone at the body's autofit font scale. A display zone (`m:oMathPara`) stacks
 * its equations, centered on one another.
 */
export function layoutMath(zone: MathZone, measurer: TextMeasurer, fontScale: number): MathBox {
  let byMeasurer = boxCache.get(zone);
  if (!byMeasurer) boxCache.set(zone, (byMeasurer = new WeakMap()));
  let byScale = byMeasurer.get(measurer);
  if (!byScale) byMeasurer.set(measurer, (byScale = new Map()));
  const hit = byScale.get(fontScale);
  if (hit) return hit;
  const env: Env = {
    measurer,
    fontScale,
    level: 0,
    display: zone.display,
    cramped: false,
    props: zone.props,
  };
  const em = emOf(zone.props, env);
  const lines = zone.lines.map((line) => rowBox(line, env));
  const box = lines.length === 1 ? lines[0]! : stackRows(lines, C.eqArrRow, em, C.axis * em);
  const base = mathFace(measurer, zone.props.latin, false, false).face;
  // The line already holds the base runs' face (`face`, `sizePt`). Past that, the equation
  // needs room for its ink only, above and below: the glyph outlines (`box`) and the marks. A
  // raised script whose ink stays inside the line does not make the line taller, also when its
  // face's ascent would pass the line, and a line after an equation starts where the ink ends,
  // with no gap added. Measured on the reference PDFs of `equations`, `equation-lines`, and
  // `preserved-content`, slide 1: stacked equations and text lines above and below them come
  // within 0.2 pt, where the faces' ascent and descent put the next line 3 to 18 pt too low.
  // A partial fit: the reference sometimes draws a numerator or a superscript lower, closer to
  // the base, than this layout does (`equation-lines`, slide 1), and makes a line of an
  // equation at more than single spacing taller than the spacing gives. Neither is modeled.
  let lineAscent = Math.max(0, box.asc);
  let lineDescent = Math.max(0, box.desc);
  for (const m of box.marks)
    for (const c of m.commands)
      if (c.op !== 'Z') {
        lineAscent = Math.max(lineAscent, c.y);
        lineDescent = Math.max(lineDescent, -c.y);
      }
  const out: MathBox = {
    lineAscent,
    lineDescent,
    width: box.w,
    ascent: box.asc,
    descent: box.desc,
    glyphs: box.glyphs,
    marks: box.marks,
    face: base,
    sizePt: sizePt(zone.props, env),
  };
  byScale.set(fontScale, out);
  return out;
}
