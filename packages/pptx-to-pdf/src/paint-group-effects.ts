/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The effects of a group (`p:grpSpPr/a:effectLst`) as PDF paint, as in the SVG paint: one shadow,
// glow, and soft edge of all the group's shapes together, from the silhouettes and edge fades of
// `paint-effects.ts`.

import { drawObject, type PDFRef } from 'pdf-lib';
import type { LayoutItem, Matrix } from '@pptx-editor.dev/core/layout';
import { blurSigma, effectBoxCorners } from '@pptx-editor.dev/core/output';
import {
  MAX_MORPH_SPREAD,
  reflectionBlurSigma,
  softEdgeSpread,
} from '@pptx-editor.dev/pro/drawing';
import {
  castGlow,
  castShadows,
  drawThroughMask,
  edgeFadeOps,
  formBox,
  MIN_SIGMA,
  proEffects,
  silhouette,
  softEdgeSilhouette,
  type Drawable,
  type Part,
  type PartClass,
} from './paint-effects.ts';
import type { FillScope } from './paint-fill.ts';
import type { PageOps } from './page-ops.ts';

/** A shape or picture that a group paints, with the fill scope it paints in. */
export interface GroupLeaf {
  readonly item: Drawable;
  readonly scope: FillScope;
  /** The scale of the group that paints it, for its line width (`drawItem`'s `groupScale`). */
  readonly groupScale: number;
}

/**
 * The silhouette of a group's content under `m`: the silhouettes of its shapes and pictures
 * together. Parts of one blur profile join one class, so a layer paints them as one
 * transparency group, and children that overlap cast one shadow, not two.
 */
async function groupSilhouette(
  leaves: readonly GroupLeaf[],
  m: Matrix,
  p: PageOps
): Promise<PartClass[]> {
  const byWidth = new Map<number, Part[]>();
  for (const leaf of leaves)
    for (const cls of await silhouette(leaf.item, m, p, leaf.scope)) {
      const parts = byWidth.get(cls.bandWidth);
      if (parts) parts.push(...cls.parts);
      else byWidth.set(cls.bandWidth, [...cls.parts]);
    }
  return [...byWidth].map(([bandWidth, parts]) => ({ parts, bandWidth }));
}

/**
 * The outer shadows and the glow of a group (`p:grpSpPr/a:effectLst`), painted before its
 * content, as in the SVG paint: one shadow of all `leaves` together, placed by the group's box.
 */
export async function drawGroupShadowsAndGlow(
  group: Extract<LayoutItem, { kind: 'group' }>,
  leaves: readonly GroupLeaf[],
  p: PageOps
): Promise<void> {
  const effects = group.element.effects;
  if (!effects || leaves.length === 0) return;
  const { cx: w, cy: h } = group.element.xfrm;
  const identity: Matrix = [1, 0, 0, 1, 0, 0];
  const under = (m: Matrix): Promise<PartClass[]> => groupSilhouette(leaves, m, p);
  await castShadows(effects, group.matrix, w, h, identity, under, p);
  await castGlow(effects.glow, () => under(identity), p);
}

/** Report how the writer draws a group's effects, when they draw something. */
export function groupEffectDiagnostics(
  group: Extract<LayoutItem, { kind: 'group' }>,
  leaves: readonly GroupLeaf[],
  p: PageOps
): void {
  const fx = group.element.effects;
  if (!fx || leaves.length === 0) return;
  const pro = proEffects();
  const shadow = fx.outerShadow ?? fx.secondShadow;
  const shadowDrawn = shadow !== undefined && shadow.color.a > 0;
  const glowDrawn = pro && fx.glow !== undefined && fx.glow.radius > 0 && fx.glow.color.a > 0;
  if (shadowDrawn && blurSigma(shadow.blur) >= MIN_SIGMA)
    p.diag.add('effect.shadow-blur', 'approximation', 'Shadow blur is drawn as layered rings.');
  if (glowDrawn) p.diag.add('effect.glow', 'approximation', 'Glow is drawn as layered rings.');
  if (shadowDrawn || glowDrawn)
    p.diag.add(
      'effect.group-silhouette',
      'approximation',
      'Group shadows and glow use an opaque silhouette of the shapes in the group: a table casts its box, and the effects of the shapes cast nothing.'
    );
  const ref = fx.reflection;
  if (
    pro &&
    ref &&
    (ref.startAlpha > 0 || ref.endAlpha > 0) &&
    reflectionBlurSigma(ref.blur) >= MIN_SIGMA
  )
    p.diag.add('effect.reflection-blur', 'approximation', 'Reflections are drawn without blur.');
  if (pro && fx.softEdge && fx.softEdge.radius > 0)
    p.diag.add(
      'effect.group-soft-edge',
      'approximation',
      'The soft edge of a group fades the lightest of the soft edges of its shapes. Its text paints over it unfaded, cut at the outline of a later shape that covers it.'
    );
  // The mask covers the shapes' areas: their own shadows, glows, and reflections outside them do
  // not show, where the editor's filter fades them with the group.
  if (
    pro &&
    fx.softEdge &&
    fx.softEdge.radius > 0 &&
    leaves.some(({ item }) => {
      const own = item.element.effects;
      return !!(own?.outerShadow || own?.secondShadow || own?.glow || own?.reflection);
    })
  )
    p.diag.add(
      'effect.group-soft-edge-effects',
      'approximation',
      'Under the soft edge of a group, the shadows, glows, and reflections of its shapes show only over the shapes.'
    );
  if (pro && fx.blur)
    p.diag.add(
      'effect.group-body',
      'approximation',
      'The blur of a group applies to each shape in the group.'
    );
}

/**
 * Paint `drawBody` (the bodies of a group's shapes, in slide space) through the group's soft edge
 * (`p:grpSpPr/a:effectLst/a:softEdge`), as the SVG renderer's body filter over the group's
 * content. The mask is the lightest of the edge fades of the group's shapes (`leaves`): each fade
 * is an isolated gray group drawn with the Lighten blend mode. Where shapes overlap, an edge of
 * one inside another stays opaque; only near where their edges cross is the mask lighter than a
 * fade of the joined outline. A shape whose fade leaves nothing adds nothing.
 */
export async function drawGroupSoftEdge(
  group: Extract<LayoutItem, { kind: 'group' }>,
  leaves: readonly GroupLeaf[],
  p: PageOps,
  drawBody: () => Promise<void>
): Promise<void> {
  const soft = group.element.effects?.softEdge;
  if (!soft || soft.radius <= 0 || !proEffects() || leaves.length === 0) return drawBody();
  const { spread, sigma } = softEdgeSpread(soft.radius);
  const inset = Math.min(Math.abs(spread), MAX_MORPH_SPREAD);
  const fades: PDFRef[] = [];
  for (const leaf of leaves) {
    const m = leaf.item.matrix;
    if (Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-9) continue;
    const shape = softEdgeSilhouette(leaf.item, leaf.scope, leaf.groupScale);
    if (shape.filled.length === 0 && shape.outlines.length === 0) continue;
    const corners = effectBoxCorners(leaf.item);
    const xs = corners.map((c) => c.x);
    const ys = corners.map((c) => c.y);
    const narrow = Math.min(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys));
    if (Math.abs(spread) > narrow / 2) continue;
    const ops = await edgeFadeOps(shape, p, inset, sigma, corners);
    fades.push(p.form(ops, formBox(p), { S: 'Transparency', CS: 'DeviceGray', I: true }));
  }
  const body = p.form(await p.capture(drawBody), formBox(p), { S: 'Transparency' });
  const maskOps = await p.capture(async () => {
    p.state('blend:Lighten', () => p.res.doc.context.obj({ Type: 'ExtGState', BM: 'Lighten' }));
    for (const fade of fades) p.push(drawObject(p.name('XObject', fade)));
  });
  drawThroughMask(p, body, maskOps);
}
