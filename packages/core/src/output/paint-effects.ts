// Shape effects as SVG (design D1 of `effects-fidelity`).
//
// A shape paints as one content group (geometry and text). Outer shadows and reflections are
// `<use>` copies of that group under the effect transform, inside untransformed groups that
// hold the filters, so blur and color always work in slide space. Glow filters the content;
// inner shadow and soft edge filter the geometry only.
//
// Approximations are listed on the public fidelity page (`docs/site/content/pptx-fidelity.mdx`)
// and in the harness report: the blur kernel (`BLUR_SIGMA_PER_RADIUS`, measured, and
// `REFLECTION_BLUR_SIGMA_PER_RADIUS`, not measured yet), glow as a
// grown, blurred outline, and soft edges as a shrunk, blurred mask (`glowSpread`,
// `softEdgeSpread`). The alpha grows and shrinks by three rectangles, not a circle.

import { toCss, type ResolvedEffects } from '../cascade/index.ts';
import { apply, multiply, type Matrix } from '../layout/index.ts';
import {
  blurSigma,
  boundsOf,
  cornersUnder,
  effectMatrix,
  fadeLine,
  glowSpread,
  grow,
  innerShadowShift,
  reflectionBlurSigma,
  softEdgeSpread,
  PX,
  type Bounds,
} from './effect-transforms.ts';
import { Defs, matrixAttr, num, svg } from './svg.ts';

/** The filter region attributes for bounds in user space. */
function region(b: Bounds): Record<string, string> {
  return {
    filterUnits: 'userSpaceOnUse',
    x: num(b.x),
    y: num(b.y),
    width: num(Math.max(1, b.r - b.x)),
    height: num(Math.max(1, b.b - b.y)),
    'color-interpolation-filters': 'sRGB',
  };
}

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
 * needs a region that holds its spread (`bodyFilter`).
 */
function morphAlpha(doc: Document, spread: number, result: string): SVGElement[] {
  const r = Math.abs(spread);
  const grows = spread > 0;
  const d = r / Math.SQRT2;
  const thin = Math.min(PX, r);
  const morph = (rx: number, ry: number, out: string): SVGElement =>
    svg(doc, 'feMorphology', {
      in: `${result}-g`,
      operator: grows ? 'dilate' : 'erode',
      radius: `${num(rx)} ${num(ry)}`,
      result: out,
    });
  const mode = grows ? 'lighten' : 'darken';
  return [
    svg(doc, 'feColorMatrix', {
      in: 'SourceAlpha',
      type: 'matrix',
      values: ALPHA_TO_GRAY,
      result: `${result}-g`,
    }),
    morph(r, thin, `${result}-h`),
    morph(thin, r, `${result}-v`),
    morph(d, d, `${result}-d`),
    svg(doc, 'feBlend', { in: `${result}-h`, in2: `${result}-v`, mode, result: `${result}-hv` }),
    svg(doc, 'feBlend', { in: `${result}-d`, in2: `${result}-hv`, mode, result: `${result}-m` }),
    svg(doc, 'feColorMatrix', {
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
export function bodyFilter(
  defs: Defs,
  effects: ResolvedEffects | undefined,
  bounds: Bounds
): string | undefined {
  if (!effects) return undefined;
  const doc = defs.doc;
  const prims: SVGElement[] = [];
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
      svg(doc, 'feComponentTransfer', { in: 'SourceAlpha', result: 'ia' }, [
        svg(doc, 'feFuncA', { type: 'table', tableValues: '1 0' }),
      ]),
      svg(doc, 'feGaussianBlur', { in: 'ia', stdDeviation: num(sigma), result: 'ib' }),
      svg(doc, 'feOffset', { in: 'ib', dx: num(dx), dy: num(dy), result: 'io' }),
      svg(doc, 'feFlood', {
        'flood-color': toCss({ ...innerShadow.color, a: 1 }),
        'flood-opacity': num(innerShadow.color.a),
        result: 'ic',
      }),
      svg(doc, 'feComposite', { in: 'ic', in2: 'io', operator: 'in', result: 'ish' }),
      svg(doc, 'feComposite', { in: 'ish', in2: 'SourceAlpha', operator: 'in', result: 'isc' }),
      svg(doc, 'feMerge', { result: 'inner' }, [
        svg(doc, 'feMergeNode', { in: 'SourceGraphic' }),
        svg(doc, 'feMergeNode', { in: 'isc' }),
      ])
    );
    source = 'inner';
  }
  if (softEdge && softEdge.radius > 0) {
    const { spread, sigma } = softEdgeSpread(softEdge.radius);
    // A kernel that reaches past the region reads transparent pixels, which `darken` skips.
    margin = Math.max(margin, Math.abs(spread) + PX);
    prims.push(
      ...morphAlpha(doc, spread, 'se'),
      svg(doc, 'feGaussianBlur', { in: 'se', stdDeviation: num(sigma), result: 'sm' }),
      // The mask holds the shape's alpha already: the shape's colors, made opaque, take it.
      svg(doc, 'feComponentTransfer', { in: source, result: 'so' }, [
        svg(doc, 'feFuncA', { type: 'discrete', tableValues: '1' }),
      ]),
      svg(doc, 'feComposite', { in: 'so', in2: 'sm', operator: 'in', result: 'sc' }),
      // Opaque colors are black outside the outline, where the blurred mask still has up to
      // 0.6 %: the outline, made binary, cuts them off.
      svg(doc, 'feComponentTransfer', { in: 'SourceAlpha', result: 'sb' }, [
        svg(doc, 'feFuncA', { type: 'linear', slope: '255', intercept: '0' }),
      ]),
      svg(doc, 'feComposite', { in: 'sc', in2: 'sb', operator: 'in', result: 'soft' })
    );
  }
  if (prims.length === 0) return undefined;
  const id = defs.id('fx');
  return defs.add(svg(doc, 'filter', { id, ...region(grow(bounds, margin)) }, prims));
}

/** Glow around the content, under it. `bounds` are the content's slide-space bounds. */
export function glowFilter(
  defs: Defs,
  effects: ResolvedEffects | undefined,
  bounds: Bounds
): string | undefined {
  const glow = effects?.glow;
  if (!glow || glow.radius <= 0) return undefined;
  const doc = defs.doc;
  const id = defs.id('glow');
  const { spread, sigma } = glowSpread(glow.radius);
  return defs.add(
    svg(doc, 'filter', { id, ...region(grow(bounds, spread + 3 * sigma + PX)) }, [
      ...morphAlpha(doc, spread, 'gd'),
      svg(doc, 'feGaussianBlur', { in: 'gd', stdDeviation: num(sigma), result: 'gb' }),
      svg(doc, 'feFlood', {
        'flood-color': toCss({ ...glow.color, a: 1 }),
        'flood-opacity': num(glow.color.a),
        result: 'gc',
      }),
      svg(doc, 'feComposite', { in: 'gc', in2: 'gb', operator: 'in', result: 'glow' }),
      svg(doc, 'feMerge', {}, [
        svg(doc, 'feMergeNode', { in: 'glow' }),
        svg(doc, 'feMergeNode', { in: 'SourceGraphic' }),
      ]),
    ])
  );
}

/** What effect copies need about the shape. */
export interface EffectSource {
  /** The id of the painted content group. */
  readonly contentId: string;
  /** The shape's local box → slide. */
  readonly matrix: Matrix;
  readonly width: number;
  readonly height: number;
  /** Slide-space corners of everything the content paints (box and text). */
  readonly corners: readonly { x: number; y: number }[];
}

function copyOf(doc: Document, source: EffectSource, m: Matrix): SVGElement {
  return svg(doc, 'use', { href: `#${source.contentId}`, transform: matrixAttr(m) });
}

/** The outer shadow: a transformed, blurred, single-color copy of the content. */
export function shadowCopy(
  defs: Defs,
  effects: ResolvedEffects | undefined,
  source: EffectSource
): SVGElement | null {
  const shadow = effects?.outerShadow;
  if (!shadow) return null;
  const doc = defs.doc;
  const m = effectMatrix(shadow, source.matrix, source.width, source.height);
  const sigma = blurSigma(shadow.blur);
  const bounds = grow(boundsOf(source.corners.map((p) => apply(m, p.x, p.y))), 3 * sigma + PX);
  const id = defs.id('shadow');
  const prims = [
    sigma > 0
      ? svg(doc, 'feGaussianBlur', { in: 'SourceAlpha', stdDeviation: num(sigma), result: 'sb' })
      : svg(doc, 'feOffset', { in: 'SourceAlpha', result: 'sb' }),
    svg(doc, 'feFlood', {
      'flood-color': toCss({ ...shadow.color, a: 1 }),
      'flood-opacity': num(shadow.color.a),
      result: 'sc',
    }),
    svg(doc, 'feComposite', { in: 'sc', in2: 'sb', operator: 'in' }),
  ];
  const filter = defs.add(svg(doc, 'filter', { id, ...region(bounds) }, prims));
  return svg(doc, 'g', { filter, 'data-effect': 'shadow' }, [copyOf(doc, source, m)]);
}

/** The reflection: a transformed, blurred copy of the content that fades along `fadeDir`. */
export function reflectionCopy(
  defs: Defs,
  effects: ResolvedEffects | undefined,
  source: EffectSource
): SVGElement | null {
  const ref = effects?.reflection;
  if (!ref || (ref.startAlpha <= 0 && ref.endAlpha <= 0)) return null;
  const doc = defs.doc;
  const m = effectMatrix(ref, source.matrix, source.width, source.height);
  const sigma = reflectionBlurSigma(ref.blur);
  const contentBounds = grow(
    boundsOf(source.corners.map((p) => apply(m, p.x, p.y))),
    3 * sigma + PX
  );
  // The fade runs across the reflected shape box, in the slide's frame or the shape's.
  const box = boundsOf(cornersUnder(multiply(m, source.matrix), source.width, source.height));
  const turn = ref.rotWithShape
    ? (Math.atan2(source.matrix[1], source.matrix[0]) * 180) / Math.PI
    : 0;
  const line = fadeLine(box, ref.fadeDir + turn);
  const gradId = defs.id('fade');
  defs.add(
    svg(
      doc,
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
        svg(doc, 'stop', {
          offset: num(ref.startPos),
          'stop-color': '#ffffff',
          'stop-opacity': num(ref.startAlpha),
        }),
        svg(doc, 'stop', {
          offset: num(Math.max(ref.startPos, ref.endPos)),
          'stop-color': '#ffffff',
          'stop-opacity': num(ref.endAlpha),
        }),
      ]
    )
  );
  const maskId = defs.id('reflect');
  const area = {
    x: num(contentBounds.x),
    y: num(contentBounds.y),
    width: num(Math.max(1, contentBounds.r - contentBounds.x)),
    height: num(Math.max(1, contentBounds.b - contentBounds.y)),
  };
  const mask = defs.add(
    svg(doc, 'mask', { id: maskId, maskUnits: 'userSpaceOnUse', ...area }, [
      svg(doc, 'rect', { ...area, fill: `url(#${gradId})` }),
    ])
  );
  let copy = copyOf(doc, source, m);
  if (sigma > 0) {
    const blurId = defs.id('rblur');
    const blur = defs.add(
      svg(doc, 'filter', { id: blurId, ...region(contentBounds) }, [
        svg(doc, 'feGaussianBlur', { in: 'SourceGraphic', stdDeviation: num(sigma) }),
      ])
    );
    copy = svg(doc, 'g', { filter: blur }, [copy]);
  }
  return svg(doc, 'g', { mask, 'data-effect': 'reflection' }, [copy]);
}
