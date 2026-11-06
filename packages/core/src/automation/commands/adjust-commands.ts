// Write the adjust values of a preset geometry (`a:prstGeom/a:avLst/a:gd`, ECMA-376 Part 1,
// 20.1.9.5) onto a shape. The editor's adjust handles and the editing API's `adjustments.set`
// share this path. A shape that takes its geometry from its layout or master gets its own
// `a:prstGeom` first, with the values it showed, so the edit changes only that shape.

import { attr, child, elements, type Element } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';

/** A guide value is a 32-bit integer (`ST_GeomGuideFormula` literals). */
export const GUIDE_LIMIT = 2_147_483_647;

/** Where a shape's preset geometry is written. */
export interface PresetTarget {
  readonly part: string;
  /** The shape's `p:spPr`. */
  readonly spPr: Element;
  /** The shape's own `a:prstGeom`, or undefined when the geometry comes from its layout. */
  readonly own: Element | undefined;
  readonly preset: string;
  /** The adjust values the shape shows now, for a new `a:prstGeom`. */
  readonly adjust: ReadonlyMap<string, number>;
}

/**
 * Write `values` (guide name to whole guide units) into the shape's `a:avLst`. The names must
 * be names of the preset table, never text from a file or a caller, because they are written
 * as attribute values.
 */
export function writeAdjustments(
  batch: OpBatch,
  t: PresetTarget,
  values: ReadonlyMap<string, number>
): void {
  for (const v of values.values())
    if (!Number.isInteger(v) || Math.abs(v) > GUIDE_LIMIT)
      throw new CommandRefusal('errors.unsupported', 'adjust value');
  const part = t.part;
  let prstGeom = t.own;
  if (!prstGeom) {
    // CT_ShapeProperties: `a:xfrm`, then the geometry (ECMA-376 Part 1, 20.1.2.2.35).
    const kids = batch.element(part, t.spPr.id).children;
    const xfrm = kids.findIndex((n) => n.type === 'element' && n.local === 'xfrm');
    const gds = [...t.adjust]
      .map(([name, v]) => `<a:gd name="${name}" fmla="val ${Math.round(v)}"/>`)
      .join('');
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: t.spPr.id,
      index: xfrm + 1,
      spec: specFromXml(`<a:prstGeom prst="${t.preset}"><a:avLst>${gds}</a:avLst></a:prstGeom>`),
    });
    prstGeom = batch.element(part, id);
  }
  // CT_PresetGeometry2D holds one optional `a:avLst`.
  let avLst = child(batch.element(part, prstGeom.id), 'a:avLst');
  if (!avLst) {
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: prstGeom.id,
      index: 0,
      spec: specFromXml('<a:avLst/>'),
    });
    avLst = batch.element(part, id);
  }
  for (const [name, v] of values) {
    const fmla = `val ${v}`;
    const list = batch.element(part, avLst.id);
    const gd = elements(list).find((g) => g.local === 'gd' && attr(g, 'name') === name);
    if (gd)
      batch.push({
        op: 'setAttribute',
        part,
        node: gd.id,
        name: { ns: '', local: 'fmla' },
        value: fmla,
      });
    else
      batch.insert({
        op: 'insertNode',
        part,
        parent: list.id,
        index: list.children.length,
        spec: specFromXml(`<a:gd name="${name}" fmla="${fmla}"/>`),
      });
  }
}
