// The transform (`a:xfrm`, ECMA-376 Part 1, 20.1.7.6) and the geometry (`a:prstGeom`,
// `a:custGeom`) of a shape's properties, as the slide cascade reads them.

import { attr, boolAttr, child, intAttr, type Element } from '../store/index.ts';
import { readAdjustValues } from './text-warp.ts';
import type { GeometrySpec, Xfrm } from './types.ts';

/** The largest offset or size the cascade reads from `a:xfrm`, in EMU (four slide widths). @internal */
export const XFRM_BOUND = 51_206_400 * 4;

export function readXfrm(xfrm: Element | undefined): Xfrm | undefined {
  if (!xfrm) return undefined;
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  if (!off && !ext) return undefined;
  const bound = (v: number | undefined): number =>
    Math.max(-XFRM_BOUND, Math.min(v ?? 0, XFRM_BOUND));
  return {
    x: bound(intAttr(off, 'x')),
    y: bound(intAttr(off, 'y')),
    cx: Math.max(0, bound(intAttr(ext, 'cx'))),
    cy: Math.max(0, bound(intAttr(ext, 'cy'))),
    rot: ((intAttr(xfrm, 'rot') ?? 0) / 60_000) % 360,
    flipH: boolAttr(xfrm, 'flipH') ?? false,
    flipV: boolAttr(xfrm, 'flipV') ?? false,
  };
}

export const ZERO_XFRM: Xfrm = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

export function readGeometry(spPr: Element | undefined): GeometrySpec | undefined {
  const prst = child(spPr, 'a:prstGeom');
  if (prst) {
    const adjust = readAdjustValues(child(prst, 'a:avLst'));
    return { kind: 'preset', preset: attr(prst, 'prst') ?? 'rect', adjust };
  }
  const cust = child(spPr, 'a:custGeom');
  if (cust) return { kind: 'custom', element: cust };
  return undefined;
}
