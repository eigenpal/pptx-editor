// Slide size: `p:sldSz` of the presentation part (ECMA-376 Part 1, 19.2.1.36), and the boxes
// of the content that sits on the slides, layouts, and masters.
//
// With `content: 'scale'`, the content scales by one factor, the smaller of the width and
// height ratios, and is centered on the new slide. So nothing leaves the slide and nothing
// distorts. A group scales its own box only: its children keep their child coordinates, so they
// follow the group. A table also scales its column widths and row heights, because its grid sets
// the drawn size. Text scales with its box: font and bullet sizes, character spacing, spacing
// in points, indents, tab stops, and text and cell insets, on the slides, layouts, masters
// (with their text styles), and the presentation's default text style. Line widths and effect
// distances keep their values. With `content: 'keep'`, nothing but the size changes.

import {
  attr,
  child,
  elements,
  find,
  NS,
  presentationInfo,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { coordinate, positiveCoordinate } from './package-edits.ts';

/** `ST_SlideSizeCoordinate`: 1 to 56 inches, in EMU (ECMA-376 Part 1, 19.7.17). @internal */
export const MIN_SLIDE_SIZE = 914_400;
/** @internal */
export const MAX_SLIDE_SIZE = 51_206_400;

/** What happens to the content when the slide size changes. */
export type SlideSizeContent = 'scale' | 'keep';

/** The part roots whose shape trees hold content on the slide canvas. */
const CANVAS_ROOTS: ReadonlySet<string> = new Set(['sld', 'sldLayout', 'sldMaster']);
/** The deepest group or alternate-content nesting the walk follows: the reader's XML depth cap. */
const MAX_DEPTH = 256;

/**
 * The DrawingML text attributes that scale with the content, by element, with the range of
 * their type: `ST_TextFontSize`, `ST_TextPoint`, `ST_TextSpacingPoint`, `ST_TextMargin`,
 * `ST_TextIndent`, and `ST_Coordinate32` (ECMA-376 Part 1, 20.1.10 and 21.1.10).
 */
const FONT_SIZE = [100, 400_000] as const;
const TEXT_POINT = [-400_000, 400_000] as const;
const MARGIN = [0, 51_206_400] as const;
const INDENT = [-51_206_400, 51_206_400] as const;
const COORDINATE32 = [-2_147_483_648, 2_147_483_647] as const;
type Range = readonly [number, number];
const PARAGRAPH: Readonly<Record<string, Range>> = { marL: MARGIN, marR: MARGIN, indent: INDENT };
const RUN: Readonly<Record<string, Range>> = { sz: FONT_SIZE, spc: TEXT_POINT };
const INSETS = (...names: string[]): Readonly<Record<string, Range>> =>
  Object.fromEntries(names.map((n) => [n, COORDINATE32]));
const TEXT_ATTRIBUTES: ReadonlyMap<string, Readonly<Record<string, Range>>> = new Map<
  string,
  Readonly<Record<string, Range>>
>([
  ['rPr', RUN],
  ['defRPr', RUN],
  ['endParaRPr', RUN],
  ['buSzPts', { val: FONT_SIZE }],
  ['spcPts', { val: [0, 158_400] }],
  ['tab', { pos: COORDINATE32 }],
  ['bodyPr', INSETS('lIns', 'tIns', 'rIns', 'bIns')],
  ['tcPr', INSETS('marL', 'marR', 'marT', 'marB')],
  ['pPr', PARAGRAPH],
  ['defPPr', PARAGRAPH],
  ...[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n): [string, Readonly<Record<string, Range>>] => [
    `lvl${n}pPr`,
    PARAGRAPH,
  ]),
]);

interface Scale {
  readonly s: number;
  readonly dx: number;
  readonly dy: number;
}

/** A slide size in EMU, refused out of `ST_SlideSizeCoordinate`. */
function slideCoordinate(value: number, what: string): number {
  const v = Math.round(value);
  if (!Number.isFinite(v) || v < MIN_SLIDE_SIZE || v > MAX_SLIDE_SIZE)
    throw new CommandRefusal('errors.slideSizeRange', what);
  return v;
}

/**
 * Set the slide size to `width` by `height` EMU. Returns false when the size does not change.
 * The edit is one transaction, so one undo step restores the size and every box.
 */
export function setSlideSize(
  batch: OpBatch,
  width: number,
  height: number,
  content: SlideSizeContent
): boolean {
  const cx = slideCoordinate(width, 'width');
  const cy = slideCoordinate(height, 'height');
  const info = presentationInfo(batch.pkg);
  const part = info.presentationPart;
  const sz = child(xmlPartOf(batch.pkg, part)?.root, 'p:sldSz');
  if (!sz) throw new CommandRefusal('errors.malformed', 'p:sldSz');
  if (cx === info.width && cy === info.height) return false;
  const set = (p: string, node: Element, name: string, value: string | null): void =>
    batch.push({
      op: 'setAttribute',
      part: p,
      node: node.id,
      name: { ns: '', local: name },
      value,
    });
  set(part, sz, 'cx', String(cx));
  set(part, sz, 'cy', String(cy));
  // `type` names a preset size (19.7.15); a new size is a custom one, the default.
  if (attr(sz, 'type') !== undefined) set(part, sz, 'type', null);
  if (content === 'keep') return true;
  const s = Math.min(cx / info.width, cy / info.height);
  const scale: Scale = { s, dx: (cx - info.width * s) / 2, dy: (cy - info.height * s) / 2 };
  for (const [name, p] of batch.pkg.parts) {
    const root = p.kind === 'xml' ? p.xml.root : undefined;
    if (!root || root.ns !== NS.p || !CANVAS_ROOTS.has(root.local)) continue;
    const tree = find(root, 'p:spTree');
    if (tree) scaleShapes(batch, name, tree, scale, 0);
    scaleText(batch, name, root, s, 0);
  }
  const defaults = child(xmlPartOf(batch.pkg, part)?.root, 'p:defaultTextStyle');
  if (defaults) scaleText(batch, part, defaults, s, 0);
  return true;
}

/** Scale the text attributes of every DrawingML element under `el` (`TEXT_ATTRIBUTES`). */
function scaleText(batch: OpBatch, part: string, el: Element, s: number, depth: number): void {
  if (depth > MAX_DEPTH) throw new CommandRefusal('errors.tooLarge');
  const names = el.ns === NS.a ? TEXT_ATTRIBUTES.get(el.local) : undefined;
  if (names)
    for (const [name, [min, max]] of Object.entries(names)) {
      const text = attr(el, name);
      const v = Number(text);
      if (text === undefined || !Number.isFinite(v)) continue;
      const next = Math.min(max, Math.max(min, Math.round(v * s)));
      if (next === v) continue;
      batch.push({
        op: 'setAttribute',
        part,
        node: el.id,
        name: { ns: '', local: name },
        value: String(next),
      });
    }
  for (const c of elements(el)) scaleText(batch, part, c, s, depth + 1);
}

/** Scale the boxes of the shapes in a shape tree, and of each copy in `mc:AlternateContent`. */
function scaleShapes(
  batch: OpBatch,
  part: string,
  parent: Element,
  scale: Scale,
  depth: number
): void {
  if (depth > MAX_DEPTH) throw new CommandRefusal('errors.tooLarge');
  for (const el of elements(parent)) {
    if (el.ns === NS.mc && el.local === 'AlternateContent') {
      // Every branch holds a copy of the same object, so every copy scales.
      for (const branch of elements(el)) scaleShapes(batch, part, branch, scale, depth + 1);
      continue;
    }
    if (el.ns !== NS.p) continue;
    const xfrm = transformOf(el);
    if (!xfrm) continue;
    scaleTransform(batch, part, xfrm, scale);
    if (el.local === 'graphicFrame') scaleTable(batch, part, el, scale.s);
  }
}

/** A shape's own transform: `a:xfrm` of its shape properties, or a frame's `p:xfrm`. */
function transformOf(el: Element): Element | undefined {
  switch (el.local) {
    case 'sp':
    case 'pic':
    case 'cxnSp':
      return child(child(el, 'p:spPr'), 'a:xfrm');
    case 'grpSp':
      return child(child(el, 'p:grpSpPr'), 'a:xfrm');
    case 'graphicFrame':
      return child(el, 'p:xfrm');
    case 'contentPart':
      return child(child(el, 'p:nvContentPartPr'), 'p:xfrm') ?? child(el, 'p14:xfrm');
    default:
      return undefined;
  }
}

function scaleTransform(batch: OpBatch, part: string, xfrm: Element, k: Scale): void {
  const set = (node: Element, name: string, value: string): void =>
    batch.push({ op: 'setAttribute', part, node: node.id, name: { ns: '', local: name }, value });
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const num = (el: Element, name: string): number | undefined => {
    const v = Number(attr(el, name));
    return attr(el, name) === undefined || !Number.isFinite(v) ? undefined : v;
  };
  if (off) {
    const x = num(off, 'x');
    const y = num(off, 'y');
    if (x !== undefined) set(off, 'x', coordinate(x * k.s + k.dx, 'x'));
    if (y !== undefined) set(off, 'y', coordinate(y * k.s + k.dy, 'y'));
  }
  if (ext) {
    const cx = num(ext, 'cx');
    const cy = num(ext, 'cy');
    if (cx !== undefined) set(ext, 'cx', positiveCoordinate(cx * k.s, 'cx'));
    if (cy !== undefined) set(ext, 'cy', positiveCoordinate(cy * k.s, 'cy'));
  }
}

/** Scale a table's column widths and row heights, which set its drawn size. */
function scaleTable(batch: OpBatch, part: string, frame: Element, s: number): void {
  const data = child(child(frame, 'a:graphic'), 'a:graphicData');
  const tbl = child(data, 'a:tbl');
  if (!tbl) return;
  const scaleAttr = (el: Element, name: string): void => {
    const v = Number(attr(el, name));
    if (attr(el, name) === undefined || !Number.isFinite(v)) return;
    batch.push({
      op: 'setAttribute',
      part,
      node: el.id,
      name: { ns: '', local: name },
      value: positiveCoordinate(v * s, name),
    });
  };
  for (const col of elements(child(tbl, 'a:tblGrid')))
    if (col.local === 'gridCol') scaleAttr(col, 'w');
  for (const row of elements(tbl)) if (row.local === 'tr') scaleAttr(row, 'h');
}
