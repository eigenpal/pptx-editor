// Shape style commands: fill, outline color, weight, and dash, and text fitting. Each writes
// `p:spPr` / `a:ln` / `a:bodyPr` children in schema order (CT_ShapeProperties,
// CT_LineProperties, CT_TextBodyProperties), so saved files stay valid.

import { attr, child, NS, type Element } from '../../store/index.ts';
import type {
  ColorChoice,
  GradientFill,
  PatternFill,
  LineDash,
  LineEndType,
  TextFitting,
  ThemeColorSlot,
  VerticalAlign,
} from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { patternXml, pictureFillXml, pictureType } from './fill-commands.ts';
import { addImagePart, specFromXml } from './shape-commands.ts';

const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
const GEOMETRY = new Set(['xfrm', 'custGeom', 'prstGeom']);
const FITS = new Set(['noAutofit', 'normAutofit', 'spAutoFit']);
const DASHES: readonly LineDash[] = [
  'solid',
  'sysDot',
  'sysDash',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
];
const HEX = /^#?[0-9a-fA-F]{6}$/;
/** The widest outline the schema allows (`ST_LineWidth`), EMU. */
export const MAX_LINE_WIDTH = 20_116_800;

const isA = (c: { type: string }, names: ReadonlySet<string>): c is Element =>
  c.type === 'element' && (c as Element).ns === NS.a && names.has((c as Element).local);

/** The theme colors a pick can name: the color map names a slide uses (`ST_SchemeColorVal`). */
const THEME_SLOTS: ReadonlySet<ThemeColorSlot> = new Set<ThemeColorSlot>([
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
]);

/** True for a value a color command accepts: `#RRGGBB` or a known theme color and variant. */
export function isColorChoice(value: unknown): value is ColorChoice {
  if (typeof value === 'string') return HEX.test(value);
  if (typeof value !== 'object' || value === null) return false;
  const { theme, shift = 0 } = value as { theme?: unknown; shift?: unknown };
  return (
    typeof theme === 'string' &&
    THEME_SLOTS.has(theme as ThemeColorSlot) &&
    typeof shift === 'number' &&
    Number.isInteger(shift) &&
    Math.abs(shift) < 100
  );
}

/**
 * The color element for a choice: `a:srgbClr`, or `a:schemeClr` with the `a:lumMod` and
 * `a:lumOff` written for "Lighter n%" and "Darker n%".
 */
export function colorXml(color: ColorChoice): string {
  if (!isColorChoice(color)) throw new CommandRefusal('errors.unsupported', 'color');
  if (typeof color === 'string')
    return `<a:srgbClr val="${color.replace(/^#/, '').toUpperCase()}"/>`;
  const shift = color.shift ?? 0;
  if (shift === 0) return `<a:schemeClr val="${color.theme}"/>`;
  const mod = `<a:lumMod val="${(100 - Math.abs(shift)) * 1000}"/>`;
  const off = shift > 0 ? `<a:lumOff val="${shift * 1000}"/>` : '';
  return `<a:schemeClr val="${color.theme}">${mod}${off}</a:schemeClr>`;
}

/** A fill element for a color: `a:solidFill`, or `a:noFill` for null. */
export function fillXml(color: ColorChoice | null): string {
  if (color === null) return '<a:noFill/>';
  return `<a:solidFill>${colorXml(color)}</a:solidFill>`;
}

function spPrOf(batch: OpBatch, part: string, id: string): Element {
  const el = batch.element(part, id);
  if (el.kind !== 'sp' && el.kind !== 'cxnSp' && el.kind !== 'pic')
    throw new CommandRefusal('disabled.unsupportedElement');
  const spPr = child(el, 'p:spPr');
  if (!spPr) throw new CommandRefusal('errors.malformed');
  return spPr;
}

/** Replace the children named in `names` with `xml`, inserted after the last of `after`. */
function replaceIn(
  batch: OpBatch,
  part: string,
  parentId: string,
  names: ReadonlySet<string>,
  after: ReadonlySet<string>,
  xml: string
): string {
  for (const c of batch.element(part, parentId).children)
    if (isA(c, names)) batch.push({ op: 'removeNode', part, node: c.id });
  const parent = batch.element(part, parentId);
  let index = 0;
  parent.children.forEach((c, i) => {
    if (isA(c, after)) index = i + 1;
  });
  return batch.insert({ op: 'insertNode', part, parent: parentId, index, spec: specFromXml(xml) });
}

/** Set the fill of shapes (`p:sp` only; pictures, connectors, and groups have none). */
export function setShapeFill(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  color: ColorChoice | null
): void {
  const shapes = ids.filter((id) => batch.element(part, id).kind === 'sp');
  if (shapes.length === 0) throw new CommandRefusal('disabled.noShapeFill');
  const xml = fillXml(color);
  for (const id of shapes) replaceIn(batch, part, spPrOf(batch, part, id).id, FILLS, GEOMETRY, xml);
}

/** `a:gradFill` for a gradient, after validating it. */
export function gradientXml(g: GradientFill): string {
  if (g.kind !== 'linear' && g.kind !== 'radial')
    throw new CommandRefusal('errors.unsupported', 'gradient');
  if (!Array.isArray(g.stops) || g.stops.length < 2 || g.stops.length > 10)
    throw new CommandRefusal('errors.unsupported', 'stops');
  const stops = [...g.stops]
    .map((s) => {
      if (!HEX.test(s.color) || !Number.isFinite(s.pos))
        throw new CommandRefusal('errors.unsupported', 'stop');
      return { pos: Math.round(Math.max(0, Math.min(1, s.pos)) * 100_000), color: s.color };
    })
    .sort((a, b) => a.pos - b.pos);
  const gs = stops
    .map(
      (s) =>
        `<a:gs pos="${s.pos}"><a:srgbClr val="${s.color.replace(/^#/, '').toUpperCase()}"/></a:gs>`
    )
    .join('');
  const angle = Number.isFinite(g.angle)
    ? ((Math.round(g.angle * 60_000) % 21_600_000) + 21_600_000) % 21_600_000
    : 0;
  const shade =
    g.kind === 'radial'
      ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
      : `<a:lin ang="${angle}" scaled="0"/>`;
  return `<a:gradFill rotWithShape="${g.rotWithShape ? 1 : 0}"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
}

/** Set a gradient fill on shapes. */
export function setShapeGradient(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  gradient: GradientFill
): void {
  const shapes = ids.filter((id) => batch.element(part, id).kind === 'sp');
  if (shapes.length === 0) throw new CommandRefusal('disabled.noShapeFill');
  const xml = gradientXml(gradient);
  for (const id of shapes) replaceIn(batch, part, spPrOf(batch, part, id).id, FILLS, GEOMETRY, xml);
}

/** Set a preset pattern fill on shapes. */
export function setShapePattern(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  pattern: PatternFill
): void {
  const shapes = ids.filter((id) => batch.element(part, id).kind === 'sp');
  if (shapes.length === 0) throw new CommandRefusal('disabled.noShapeFill');
  const xml = patternXml(pattern);
  for (const id of shapes) replaceIn(batch, part, spPrOf(batch, part, id).id, FILLS, GEOMETRY, xml);
}

/** Set a picture fill on shapes: one media part, stretched or tiled. */
export function setShapePicture(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  bytes: Uint8Array,
  tile: boolean
): void {
  const shapes = ids.filter((id) => batch.element(part, id).kind === 'sp');
  if (shapes.length === 0) throw new CommandRefusal('disabled.noShapeFill');
  const xml = pictureFillXml(addImagePart(batch, part, bytes, pictureType(bytes)), tile === true);
  for (const id of shapes) replaceIn(batch, part, spPrOf(batch, part, id).id, FILLS, GEOMETRY, xml);
}

/** The shape's `a:ln`, created after the fill when absent. */
function lineOf(batch: OpBatch, part: string, id: string): Element {
  const spPr = spPrOf(batch, part, id);
  const ln = child(spPr, 'a:ln');
  if (ln) return ln;
  const after = new Set([...GEOMETRY, ...FILLS]);
  const lnId = replaceIn(batch, part, spPr.id, new Set(), after, '<a:ln/>');
  return batch.element(part, lnId);
}

/** Set the outline color (null: no outline). */
export function setLineColor(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  color: ColorChoice | null
): void {
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const xml = fillXml(color);
  for (const id of ids) replaceIn(batch, part, lineOf(batch, part, id).id, FILLS, new Set(), xml);
}

/**
 * Set the outline width. A shape whose outline draws nothing gets the theme text color, so a
 * weight change is always visible.
 */
export function setLineWidth(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  points: number,
  visible: (id: string) => boolean
): void {
  if (!(points > 0 && points <= 1584)) throw new CommandRefusal('errors.unsupported', 'width');
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const w = String(Math.min(MAX_LINE_WIDTH, Math.round(points * 12_700)));
  for (const id of ids) {
    const ln = lineOf(batch, part, id);
    batch.push({ op: 'setAttribute', part, node: ln.id, name: { ns: '', local: 'w' }, value: w });
    if (!visible(id))
      replaceIn(
        batch,
        part,
        ln.id,
        FILLS,
        new Set(),
        '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
      );
  }
}

/** Set the outline dash preset. */
export function setLineDash(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  dash: LineDash
): void {
  if (!DASHES.includes(dash)) throw new CommandRefusal('errors.unsupported', 'dash');
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  for (const id of ids)
    replaceIn(
      batch,
      part,
      lineOf(batch, part, id).id,
      new Set(['prstDash', 'custDash']),
      FILLS,
      `<a:prstDash val="${dash}"/>`
    );
}

/** Set text fitting on shapes with a text body. Returns the shapes changed. */
export function setTextFitting(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  mode: TextFitting
): string[] {
  const xml =
    mode === 'shape' ? '<a:spAutoFit/>' : mode === 'shrink' ? '<a:normAutofit/>' : '<a:noAutofit/>';
  if (mode !== 'none' && mode !== 'shrink' && mode !== 'shape')
    throw new CommandRefusal('errors.unsupported', 'fitting');
  const out: string[] = [];
  for (const id of ids) {
    const bodyPr = child(child(batch.element(part, id), 'p:txBody'), 'a:bodyPr');
    if (!bodyPr) continue;
    replaceIn(batch, part, bodyPr.id, FITS, new Set(['prstTxWarp']), xml);
    out.push(id);
  }
  if (out.length === 0) throw new CommandRefusal('disabled.noText');
  return out;
}

/** Set the vertical alignment of text in shapes. */
export function setVerticalAlign(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  anchor: VerticalAlign
): void {
  if (anchor !== 't' && anchor !== 'ctr' && anchor !== 'b')
    throw new CommandRefusal('errors.unsupported', 'anchor');
  let changed = 0;
  for (const id of ids) {
    const bodyPr = child(child(batch.element(part, id), 'p:txBody'), 'a:bodyPr');
    if (!bodyPr) continue;
    batch.push({
      op: 'setAttribute',
      part,
      node: bodyPr.id,
      name: { ns: '', local: 'anchor' },
      value: anchor,
    });
    changed += 1;
  }
  if (changed === 0) throw new CommandRefusal('disabled.noText');
}

const LINE_PRESETS =
  /^(line|lineInv|straightConnector1|bentConnector[2-5]|curvedConnector[2-5]|arc)$/;
const END_TYPES: readonly LineEndType[] = [
  'none',
  'triangle',
  'arrow',
  'stealth',
  'oval',
  'diamond',
];

/** Whether an element is a line: a connector, or a shape with a line preset. */
export function isLine(el: Element): boolean {
  if (el.kind === 'cxnSp') return true;
  if (el.kind !== 'sp') return false;
  return LINE_PRESETS.test(attr(child(child(el, 'p:spPr'), 'a:prstGeom'), 'prst') ?? '');
}

/** Set the arrowheads of lines. `head` is the start of the path, `tail` its end. */
export function setLineEnds(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  head: LineEndType | undefined,
  tail: LineEndType | undefined
): void {
  for (const t of [head, tail])
    if (t !== undefined && !END_TYPES.includes(t))
      throw new CommandRefusal('errors.unsupported', 'lineEnd');
  const lines = ids.filter((id) => isLine(batch.element(part, id)));
  if (lines.length === 0) throw new CommandRefusal('disabled.notALine');
  const before = new Set([...FILLS, 'prstDash', 'custDash', 'round', 'bevel', 'miter']);
  for (const id of lines) {
    if (head !== undefined)
      replaceIn(
        batch,
        part,
        lineOf(batch, part, id).id,
        new Set(['headEnd']),
        before,
        `<a:headEnd type="${head}"/>`
      );
    if (tail !== undefined)
      replaceIn(
        batch,
        part,
        lineOf(batch, part, id).id,
        new Set(['tailEnd']),
        new Set([...before, 'headEnd']),
        `<a:tailEnd type="${tail}"/>`
      );
  }
}
