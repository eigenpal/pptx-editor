/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Glow, inner shadow, reflection, and soft edge as SVG (design D1 of `effects-fidelity`), as
// plain data that the engine's SVG paint builds (`EffectNode`).
//
// A shape paints as one content group (geometry and text). A reflection is a `<use>` copy of that
// group under the effect transform, inside an untransformed group that holds the mask, so blur
// and fade always work in slide space. Glow filters the content; inner shadow and soft edge
// filter the geometry only.
//
// Approximations are listed on the public fidelity page (`docs/site/content/pptx-fidelity.mdx`)
// and in the harness report: the reflection blur kernel (`REFLECTION_BLUR_SIGMA_PER_RADIUS`, not
// measured yet), glow as a grown, blurred outline, and soft edges as a shrunk, blurred mask
// (`glowSpread`, `softEdgeSpread`). The alpha grows and shrinks by three rectangles, not a
// circle.

import { toCss, type ResolvedEffects } from '@pptx-editor.dev/core/cascade';
import {
  apply,
  type DrawingEffectsRenderer,
  type EffectBounds,
  type EffectCopyPaint,
  type EffectCopySource,
  type EffectFilterPaint,
  type EffectIds,
  type EffectNode,
  type Matrix,
} from '@pptx-editor.dev/core/layout';
import { blurSigma } from '@pptx-editor.dev/core/output';
import {
  boundsOf,
  glowSpread,
  grow,
  innerShadowShift,
  PX,
  reflectionBlurSigma,
  reflectionFade,
  softEdgeSpread,
} from './effects-math.ts';

/** Format a number for an SVG attribute, as the engine's paint does. */
function num(v: number): string {
  if (!Number.isFinite(v)) return '0';
  const r = Math.round(v * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
}

function node(
  tag: string,
  attrs: Record<string, string> = {},
  children: readonly EffectNode[] = []
): EffectNode {
  return children.length > 0 ? { tag, attrs, children } : { tag, attrs };
}

function matrixAttr(m: Matrix): string {
  return `matrix(${m.map(num).join(' ')})`;
}

/** The filter region attributes for bounds in user space. */
function region(b: EffectBounds): Record<string, string> {
  return {
    filterUnits: 'userSpaceOnUse',
    x: num(b.x),
    y: num(b.y),
    width: num(Math.max(1, b.r - b.x)),
    height: num(Math.max(1, b.b - b.y)),
    'color-interpolation-filters': 'sRGB',
  };
}

/**
 * The most a glow grows, or a soft edge shrinks, the outline: 100 pt, in EMU. The browser's
 * `feMorphology` costs time in proportion to its radius on every paint, so a radius of a few
 * inches stops the page for seconds. Glows of up to 200 pt and soft edges of up to 100 pt draw
 * as stored. A larger glow draws as a 200 pt glow, and a larger soft edge shrinks the outline
 * by 100 pt and keeps its blur.
 */
const MAX_MORPH_SPREAD = 1_270_000;

/** `SourceAlpha` as an opaque gray image: each color channel holds the alpha. */
const ALPHA_TO_GRAY = '0 0 0 1 0  0 0 0 1 0  0 0 0 1 0  0 0 0 0 1';
/** Back from the gray image: the red channel becomes the alpha. */
const GRAY_TO_ALPHA = '0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  1 0 0 0 0';

/**
 * Grow (`spread` > 0) or shrink (`spread` < 0) the alpha of `SourceAlpha` by `|spread|` in every
 * direction. `feMorphology` takes a rectangle, so a square alone grows a curve's diagonals by up
 * to √2 times the spread. Three rectangles do better: a horizontal bar, a vertical bar (each one
 * pixel thick, because a zero radius on one axis can turn the primitive off), and a square of
 * `spread / √2`. A grow is their union, a shrink their intersection. The reach is within 8 % of
 * the spread on a curved edge. At a sharp corner, a direction close to an edge reaches only the
 * square's side, so the corner can be up to 29 % short.
 *
 * A union is the largest of the three alphas, an intersection the smallest, so a fill of alpha
 * `a` stays `a`. Alpha compositing cannot take a maximum, so the alpha moves into the color
 * channels of an opaque image, where `feBlend` `lighten` and `darken` are exactly the maximum and
 * the minimum, and back. Every input stays opaque only inside the filter region, so a shrink
 * needs a region that holds its spread (`body`).
 */
function morphAlpha(spread: number, result: string, bounds: EffectBounds): EffectNode[] {
  const r = Math.min(Math.abs(spread), MAX_MORPH_SPREAD);
  // The alpha moves at most `r` from the content: the primitives work only on the bounds and
  // `r` around them, not on a filter region that a wide blur makes much larger.
  const b = grow(bounds, r + PX);
  const sub = {
    x: num(b.x),
    y: num(b.y),
    width: num(Math.max(1, b.r - b.x)),
    height: num(Math.max(1, b.b - b.y)),
  };
  const grows = spread > 0;
  const d = r / Math.SQRT2;
  const thin = Math.min(PX, r);
  const morph = (rx: number, ry: number, out: string): EffectNode =>
    node('feMorphology', {
      ...sub,
      in: `${result}-g`,
      operator: grows ? 'dilate' : 'erode',
      radius: `${num(rx)} ${num(ry)}`,
      result: out,
    });
  const mode = grows ? 'lighten' : 'darken';
  return [
    node('feColorMatrix', {
      ...sub,
      in: 'SourceAlpha',
      type: 'matrix',
      values: ALPHA_TO_GRAY,
      result: `${result}-g`,
    }),
    morph(r, thin, `${result}-h`),
    morph(thin, r, `${result}-v`),
    morph(d, d, `${result}-d`),
    node('feBlend', {
      ...sub,
      in: `${result}-h`,
      in2: `${result}-v`,
      mode,
      result: `${result}-hv`,
    }),
    node('feBlend', {
      ...sub,
      in: `${result}-d`,
      in2: `${result}-hv`,
      mode,
      result: `${result}-m`,
    }),
    node('feColorMatrix', {
      ...sub,
      in: `${result}-m`,
      type: 'matrix',
      values: GRAY_TO_ALPHA,
      result,
    }),
  ];
}

/**
 * Inner shadow and soft edge, applied to the shape's geometry. `bounds` are the geometry's
 * slide-space bounds.
 */
export function bodyEffects(
  effects: ResolvedEffects,
  bounds: EffectBounds,
  ids: EffectIds
): EffectFilterPaint | null {
  const prims: EffectNode[] = [];
  const { softEdge, innerShadow } = effects;
  let source = 'SourceGraphic';
  // `bounds` hold the geometry and its outline. The region adds a tenth of the larger side, as
  // the old bounding-box region did, and room for the inverse that the inner shadow moves and
  // blurs, so the inverse never ends inside the blur.
  let margin = Math.max(bounds.r - bounds.x, bounds.b - bounds.y) / 10 + PX;
  if (innerShadow) {
    const sigma = blurSigma(innerShadow.blur);
    const { dx, dy } = innerShadowShift(innerShadow.dist, innerShadow.dir);
    margin += Math.abs(innerShadow.dist) + 3 * sigma;
    prims.push(
      node('feComponentTransfer', { in: 'SourceAlpha', result: 'ia' }, [
        node('feFuncA', { type: 'table', tableValues: '1 0' }),
      ]),
      node('feGaussianBlur', { in: 'ia', stdDeviation: num(sigma), result: 'ib' }),
      node('feOffset', { in: 'ib', dx: num(dx), dy: num(dy), result: 'io' }),
      node('feFlood', {
        'flood-color': toCss({ ...innerShadow.color, a: 1 }),
        'flood-opacity': num(innerShadow.color.a),
        result: 'ic',
      }),
      node('feComposite', { in: 'ic', in2: 'io', operator: 'in', result: 'ish' }),
      node('feComposite', { in: 'ish', in2: 'SourceAlpha', operator: 'in', result: 'isc' }),
      node('feMerge', { result: 'inner' }, [
        node('feMergeNode', { in: 'SourceGraphic' }),
        node('feMergeNode', { in: 'isc' }),
      ])
    );
    source = 'inner';
  }
  if (softEdge && softEdge.radius > 0) {
    const { spread, sigma } = softEdgeSpread(softEdge.radius);
    // A shrink by more than half the narrower side of the bounds leaves nothing: the horizontal
    // or the vertical rectangle of `morphAlpha` fits nowhere. The shape then paints nothing,
    // with no `feMorphology` that costs time in proportion to the radius.
    if (Math.abs(spread) > Math.min(bounds.r - bounds.x, bounds.b - bounds.y) / 2) {
      const id = ids.id('fx');
      const empty = node('feFlood', { 'flood-opacity': '0' });
      return { defs: [node('filter', { id, ...region(bounds) }, [empty])], filter: id };
    }
    // A kernel that reaches past the region reads transparent pixels, which `darken` skips.
    margin = Math.max(margin, Math.min(Math.abs(spread), MAX_MORPH_SPREAD) + PX);
    prims.push(
      ...morphAlpha(spread, 'se', bounds),
      node('feGaussianBlur', { in: 'se', stdDeviation: num(sigma), result: 'sm' }),
      // The mask holds the shape's alpha already: the shape's colors, made opaque, take it.
      node('feComponentTransfer', { in: source, result: 'so' }, [
        node('feFuncA', { type: 'discrete', tableValues: '1' }),
      ]),
      node('feComposite', { in: 'so', in2: 'sm', operator: 'in', result: 'sc' }),
      // Opaque colors are black outside the outline, where the blurred mask still has up to
      // 0.6 %: the outline, made binary, cuts them off.
      node('feComponentTransfer', { in: 'SourceAlpha', result: 'sb' }, [
        node('feFuncA', { type: 'linear', slope: '255', intercept: '0' }),
      ]),
      node('feComposite', { in: 'sc', in2: 'sb', operator: 'in', result: 'soft' })
    );
  }
  if (prims.length === 0) return null;
  const id = ids.id('fx');
  return { defs: [node('filter', { id, ...region(grow(bounds, margin)) }, prims)], filter: id };
}

/** Glow around the content, under it. `bounds` are the content's slide-space bounds. */
export function glowEffect(
  effects: ResolvedEffects,
  bounds: EffectBounds,
  ids: EffectIds
): EffectFilterPaint | null {
  const glow = effects.glow;
  if (!glow || glow.radius <= 0) return null;
  const id = ids.id('glow');
  // A glow past 200 pt draws as a 200 pt glow: it grows the outline by the most that
  // `morphAlpha` grows it, and its blur stays within a filter region that the browser draws.
  const { spread, sigma } = glowSpread(Math.min(glow.radius, 2 * MAX_MORPH_SPREAD));
  const filter = node('filter', { id, ...region(grow(bounds, spread + 3 * sigma + PX)) }, [
    ...morphAlpha(spread, 'gd', bounds),
    node('feGaussianBlur', { in: 'gd', stdDeviation: num(sigma), result: 'gb' }),
    node('feFlood', {
      'flood-color': toCss({ ...glow.color, a: 1 }),
      'flood-opacity': num(glow.color.a),
      result: 'gc',
    }),
    node('feComposite', { in: 'gc', in2: 'gb', operator: 'in', result: 'glow' }),
    node('feMerge', {}, [
      node('feMergeNode', { in: 'glow' }),
      node('feMergeNode', { in: 'SourceGraphic' }),
    ]),
  ]);
  return { defs: [filter], filter: id };
}

/** The reflection: a transformed, blurred copy of the content that fades along `fadeDir`. */
export function reflectionEffect(
  effects: ResolvedEffects,
  source: EffectCopySource,
  ids: EffectIds
): EffectCopyPaint | null {
  const ref = effects.reflection;
  if (!ref || (ref.startAlpha <= 0 && ref.endAlpha <= 0)) return null;
  const {
    matrix: m,
    line,
    start,
    end,
  } = reflectionFade(ref, source.matrix, source.width, source.height);
  const sigma = reflectionBlurSigma(ref.blur);
  const contentBounds = grow(
    boundsOf(source.corners.map((p) => apply(m, p.x, p.y))),
    3 * sigma + PX
  );
  const defs: EffectNode[] = [];
  const gradId = ids.id('fade');
  defs.push(
    node(
      'linearGradient',
      {
        id: gradId,
        gradientUnits: 'userSpaceOnUse',
        x1: num(line.x1),
        y1: num(line.y1),
        x2: num(line.x2),
        y2: num(line.y2),
      },
      [
        node('stop', {
          offset: num(start),
          'stop-color': '#ffffff',
          'stop-opacity': num(ref.startAlpha),
        }),
        node('stop', {
          offset: num(end),
          'stop-color': '#ffffff',
          'stop-opacity': num(ref.endAlpha),
        }),
      ]
    )
  );
  const maskId = ids.id('reflect');
  const area = {
    x: num(contentBounds.x),
    y: num(contentBounds.y),
    width: num(Math.max(1, contentBounds.r - contentBounds.x)),
    height: num(Math.max(1, contentBounds.b - contentBounds.y)),
  };
  defs.push(
    node('mask', { id: maskId, maskUnits: 'userSpaceOnUse', ...area }, [
      node('rect', { ...area, fill: `url(#${gradId})` }),
    ])
  );
  let copy = node('use', { href: `#${source.contentId}`, transform: matrixAttr(m) });
  if (sigma > 0) {
    const blurId = ids.id('rblur');
    defs.push(
      node('filter', { id: blurId, ...region(contentBounds) }, [
        node('feGaussianBlur', { in: 'SourceGraphic', stdDeviation: num(sigma) }),
      ])
    );
    copy = node('g', { filter: `url(#${blurId})` }, [copy]);
  }
  return {
    defs,
    node: node('g', { mask: `url(#${maskId})`, 'data-effect': 'reflection' }, [copy]),
  };
}

/** The effects renderer that `installDrawing` installs. */
export const drawingEffectsRenderer: DrawingEffectsRenderer = Object.freeze({
  body: bodyEffects,
  glow: glowEffect,
  reflection: reflectionEffect,
});
