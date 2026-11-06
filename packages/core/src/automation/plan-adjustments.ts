// The adjust values of a preset geometry (`a:prstGeom/a:avLst/a:gd`, ECMA-376 Part 1,
// 20.1.9.5). The order, the defaults, and the handle ranges come from the preset's own
// definition (`reference/ecma-376/part1/drawingml-geometries/presetShapeDefinitions.xml`).

import { presetAdjustments, type PresetAdjustment } from '../layout/index.ts';
import { attr, child, elements, intAttr, type Element } from '../store/index.ts';
import { GUIDE_LIMIT, writeAdjustments } from './commands/adjust-commands.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import { resolvedShape } from './plan-reads.ts';
import type { TextPlanContext } from './plan-text.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  indexIn,
  shapeAt,
  type ShapeLocation,
} from './targets.ts';

/** Guide units per object-model unit: a length is a fraction, an angle is in degrees. */
const LENGTH_UNIT = 100_000;
const ANGLE_UNIT = 60_000;

/** The adjust values a shape stores in its own `a:avLst` (`val` formulas only). */
function storedValues(prstGeom: Element | undefined): Map<string, number> {
  const out = new Map<string, number>();
  for (const gd of elements(child(prstGeom, 'a:avLst'))) {
    const name = attr(gd, 'name');
    const m = /^val\s+(-?\d+)$/.exec((attr(gd, 'fmla') ?? '').trim());
    if (name && m) out.set(name, Number(m[1]));
  }
  return out;
}

/** What a shape's preset geometry is, and where it comes from. */
interface PresetOf {
  readonly preset: string;
  readonly adjust: ReadonlyMap<string, number>;
  /** The shape's own `a:prstGeom`, or undefined when it takes the geometry of its layout. */
  readonly own: Element | undefined;
  readonly spPr: Element;
  readonly adjustments: PresetAdjustment[];
}

function noPreset(): never {
  throw new AutomationFailure(
    'unsupported-content',
    'the shape has no preset geometry',
    'no-preset-geometry'
  );
}

/** The preset geometry of a shape, a picture, or a connector, or `no-preset-geometry`. */
function presetOf(ctx: TextPlanContext, at: ShapeLocation): PresetOf {
  const spPr = at.el.kind === 'grpSp' ? undefined : child(at.el, 'p:spPr');
  if (!spPr) noPreset();
  if (child(spPr, 'a:custGeom')) noPreset();
  const own = child(spPr, 'a:prstGeom');
  const resolved = resolvedShape(ctx, at);
  const spec =
    resolved &&
    (resolved.kind === 'shape' || resolved.kind === 'connector' || resolved.kind === 'picture')
      ? resolved.geometry
      : undefined;
  let preset: string;
  let adjust: ReadonlyMap<string, number>;
  if (own) {
    preset = attr(own, 'prst') ?? '';
    adjust = storedValues(own);
  } else if (spec?.kind === 'preset') {
    // A placeholder takes its geometry from its layout or master (ECMA-376 Part 1, 19.3.1.36).
    preset = spec.preset;
    adjust = spec.adjust;
  } else noPreset();
  const xfrm = child(spPr, 'a:xfrm');
  const w = resolved?.xfrm.cx ?? intAttr(child(xfrm, 'a:ext'), 'cx') ?? 0;
  const h = resolved?.xfrm.cy ?? intAttr(child(xfrm, 'a:ext'), 'cy') ?? 0;
  // An unknown name is refused here, so a preset and its guide names written below are always
  // names of the preset table, never text from the file.
  const adjustments = presetAdjustments(preset, w, h, adjust);
  if (!adjustments) noPreset();
  return { preset, adjust, own, spPr, adjustments };
}

const unitOf = (a: PresetAdjustment): number => (a.angle ? ANGLE_UNIT : LENGTH_UNIT);

export function readAdjustments(ctx: TextPlanContext, ref: AutomationRef): AutomationValue {
  const { preset, adjustments } = presetOf(ctx, shapeAt(ctx, ref));
  return {
    kind: 'adjustments',
    adjustments: { preset, values: adjustments.map((a) => a.value / unitOf(a)) },
  };
}

/**
 * Set one adjust value. The value is bounded to the range of the preset's adjust handle at the
 * shape's size; a value outside the 32-bit guide range is refused. A shape that takes its
 * geometry from its layout gets its own `a:prstGeom`, with the values it showed.
 */
export function setAdjustment(
  ctx: TextPlanContext,
  ref: AutomationRef,
  index: number,
  value: number
): AutomationValue {
  const at = shapeAt(ctx, ref);
  const geometry = presetOf(ctx, at);
  const slot = geometry.adjustments[indexIn(index, 'index', geometry.adjustments.length)]!;
  if (typeof value !== 'number' || !Number.isFinite(value)) badArgument('value');
  let guide = Math.round(value * unitOf(slot));
  if (Math.abs(guide) > GUIDE_LIMIT) badArgument('value');
  if (slot.min !== null && guide < slot.min) guide = Math.ceil(slot.min);
  if (slot.max !== null && guide > slot.max) guide = Math.floor(slot.max);
  edit(() =>
    writeAdjustments(ctx.batch, { part: at.part, ...geometry }, new Map([[slot.name, guide]]))
  );
  return { kind: 'applied' };
}
