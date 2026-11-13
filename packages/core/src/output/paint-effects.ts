// Shape effects as SVG (design D1 of `effects-fidelity`).
//
// A shape paints as one content group (geometry and text). An outer shadow is a `<use>` copy of
// that group under the effect transform, inside an untransformed group that holds the filter, so
// blur and color always work in slide space. A group with effects paints its children as one
// content group in the same way, placed by the group's box (`paintGroupEffects`).
//
// Glow, inner shadow, reflection, and soft edge come from the installed effects renderer
// (`drawing-renderers.ts`, `@pptx-editor.dev/pro/drawing`) as plain data, which `effectDefs`
// and `effectElement` build here. Without the renderer, a shape paints without them.
//
// Approximations are listed on the public fidelity page (`docs/site/content/pptx-fidelity.mdx`)
// and in the harness report: the blur kernel (`BLUR_SIGMA_PER_RADIUS`, measured).

import { toCss, type ResolvedEffects } from '../cascade/index.ts';
import {
  apply,
  drawingRenderers,
  multiply,
  type EffectCopySource,
  type LayoutItem,
  type EffectNode,
  type Matrix,
  type PathCommand,
  type ThreeDFacet,
} from '../layout/index.ts';
import { blurSigma, boundsOf, effectMatrix, grow, PX, type Bounds } from './effect-transforms.ts';
import { groupTextCovers } from './group-text-covers.ts';
import { groupContentCorners } from './item-corners.ts';
import { commandData } from './path-data.ts';
import type { PaintContext } from './paint-fill.ts';
import { matrixAttr, num, svg, type Defs } from './svg.ts';
import { finiteFacet } from './painter-math.ts';

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

/**
 * The elements an effects renderer may return: paint servers, masks, filters, their primitives,
 * groups, and copies of painted content. Anything else (a script, foreign content, a link)
 * builds as an empty group.
 */
const EFFECT_TAGS: ReadonlySet<string> = new Set([
  'g',
  'use',
  'rect',
  'mask',
  'linearGradient',
  'stop',
  'filter',
  'feBlend',
  'feColorMatrix',
  'feComponentTransfer',
  'feComposite',
  'feFlood',
  'feFuncA',
  'feFuncB',
  'feFuncG',
  'feFuncR',
  'feGaussianBlur',
  'feMerge',
  'feMergeNode',
  'feMorphology',
  'feOffset',
]);

/** Whether a renderer attribute may be set: no event handlers, and links only within the slide. */
function effectAttr(name: string, value: string): boolean {
  const key = name.toLowerCase();
  if (key.startsWith('on')) return false;
  if (key === 'href' || key === 'xlink:href') return value.startsWith('#');
  // A paint, mask, or filter reference names a definition of the slide, never a URL.
  return !/url\(\s*(?!['"]?#)/i.test(value);
}

/**
 * Build one renderer node as SVG: attributes through `setAttribute`, never parsed markup. Only
 * effect elements and attributes build (`EFFECT_TAGS`, `effectAttr`), so a renderer can never
 * add script or fetch a resource.
 */
export function effectElement(doc: Document, node: EffectNode): SVGElement {
  if (!EFFECT_TAGS.has(node.tag)) return svg(doc, 'g');
  const attrs = Object.fromEntries(
    Object.entries(node.attrs).filter(
      ([name, value]) => typeof value === 'string' && effectAttr(name, value)
    )
  );
  return svg(
    doc,
    node.tag,
    attrs,
    (node.children ?? []).map((child) => effectElement(doc, child))
  );
}

/** Add a renderer's definitions to `defs`. */
export function effectDefs(defs: Defs, nodes: readonly EffectNode[]): void {
  for (const node of nodes) defs.add(effectElement(defs.doc, node));
}

/** What effect copies need about the shape. */
export type EffectSource = EffectCopySource;

function copyOf(doc: Document, source: EffectSource, m: Matrix): SVGElement {
  return svg(doc, 'use', { href: `#${source.contentId}`, transform: matrixAttr(m) });
}

/**
 * The polygons of a 3-D paint (bevel strips, contour, extrusion sides) as paths. Each opaque one
 * has a thin outline of its own color, so neighbors leave no antialiased seam between them.
 */
export function facetElements(
  doc: Document,
  facets: readonly ThreeDFacet[],
  seam: number
): SVGElement[] {
  return facets.filter(finiteFacet).map((f) => {
    const d = f.polygons
      .filter((poly) => poly.length >= 3)
      .map((poly) => poly.map((p, i) => `${i ? 'L' : 'M'}${num(p.x)} ${num(p.y)}`).join('') + 'Z')
      .join('');
    const color = toCss({ ...f.color, a: 1 });
    const opacity = f.color.a < 1 ? num(f.color.a) : undefined;
    return svg(doc, 'path', {
      d,
      fill: color,
      'fill-opacity': opacity,
      // A translucent face light has no seams to hide, and an outline would darken its edge.
      ...(opacity === undefined
        ? { stroke: color, 'stroke-width': num(seam), 'stroke-linejoin': 'round' }
        : {}),
    });
  });
}

/** The outer shadow: a transformed, blurred, single-color copy of the content. */
export function shadowCopy(
  defs: Defs,
  effects: ResolvedEffects | undefined,
  source: EffectSource,
  which: 'outerShadow' | 'secondShadow' = 'outerShadow'
): SVGElement | null {
  const shadow = effects?.[which];
  if (!shadow) return null;
  const doc = defs.doc;
  const placed = effectMatrix(shadow, source.matrix, source.width, source.height);
  // On a turned face, the shadow is placed on the face and turns with it.
  const face = source.face;
  const m = face ? multiply(face.front, placed) : placed;
  const corners = face ? face.corners : source.corners;
  const sigma = blurSigma(shadow.blur);
  const bounds = grow(boundsOf(corners.map((p) => apply(m, p.x, p.y))), 3 * sigma + PX);
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
  const copy = face
    ? svg(doc, 'use', { href: `#${face.id}`, transform: matrixAttr(m) })
    : copyOf(doc, source, m);
  return svg(doc, 'g', { filter, 'data-effect': 'shadow' }, [copy]);
}

/**
 * Whether a group paints effects of its own (`p:grpSpPr/a:effectLst`): outer and preset shadows
 * always, and glow, inner shadow, reflection, soft edge, and blur through the installed effects
 * renderer, as for a shape. @internal
 */
export function groupPaintsEffects(group: Extract<LayoutItem, { kind: 'group' }>): boolean {
  const fx = group.element.effects;
  if (!fx) return false;
  if (fx.outerShadow || fx.secondShadow) return true;
  const pro = drawingRenderers().effects;
  if (!pro) return false;
  return !!(fx.glow || fx.reflection || fx.softEdge || (pro.blur && fx.blur));
}

/**
 * Paint a group's effects around its painted children (`inner`), into the group's node `outer`.
 * The effects apply to the group's content as one picture (ECMA-376 Part 1, 19.3.1.22): the
 * shadow, the glow, and the reflection copy all children together, and the soft edge, the inner
 * shadow, and the blur filter them together. The order is the shape's: reflection, shadows,
 * glow, then the content.
 */
export function paintGroupEffects(
  doc: Document,
  group: Extract<LayoutItem, { kind: 'group' }>,
  outer: SVGElement,
  inner: SVGElement,
  ctx: PaintContext
): void {
  const effects = group.element.effects!;
  const pro = drawingRenderers().effects;
  const ids = ctx.defs;
  const corners = groupContentCorners(group);
  const bounds = boundsOf(corners);
  // A group's inner shadow is not resolved (`ResolvedGroup.effects`): the body is its soft edge.
  const bodyFx = pro && effects.softEdge ? pro.body(effects, bounds, ids) : null;
  if (bodyFx) effectDefs(ctx.defs, bodyFx.defs);
  const blurFx = pro?.blur && effects.blur ? pro.blur(effects, bounds, ids) : null;
  if (blurFx) effectDefs(ctx.defs, blurFx.defs);
  const content = svg(doc, 'g', {}, [
    bodyFx ? softEdged(doc, group, inner, bodyFx.filter, ctx) : inner,
  ]);
  const contentId = ctx.defs.id('content');
  content.setAttribute('id', contentId);
  const shown = blurFx ? svg(doc, 'g', { filter: `url(#${blurFx.filter})` }, [content]) : content;
  const source: EffectSource = {
    contentId,
    matrix: group.matrix,
    width: group.element.xfrm.cx,
    height: group.element.xfrm.cy,
    corners,
  };
  const reflection = pro && effects.reflection ? pro.reflection(effects, source, ids) : null;
  if (reflection) {
    effectDefs(ctx.defs, reflection.defs);
    outer.appendChild(effectElement(doc, reflection.node));
  }
  // A preset shadow's second copy lies under its first.
  const second = shadowCopy(ctx.defs, effects, source, 'secondShadow');
  if (second) outer.appendChild(second);
  const shadow = shadowCopy(ctx.defs, effects, source);
  if (shadow) outer.appendChild(shadow);
  const glow = pro && effects.glow ? pro.glow(effects, bounds, ids, true) : null;
  if (glow) {
    effectDefs(ctx.defs, glow.defs);
    let glowSource = contentId;
    if (shown !== content) {
      glowSource = ctx.defs.id('shown');
      shown.setAttribute('id', glowSource);
    }
    outer.appendChild(
      svg(doc, 'g', { filter: `url(#${glow.filter})`, 'data-effect': 'glow' }, [
        svg(doc, 'use', { href: `#${glowSource}` }),
      ])
    );
  }
  outer.appendChild(shown);
}

/** The custom property that hides the text of shapes in the filtered copy (`softEdged`). */
const COPY_TEXT = '--pptx-group-text';

/**
 * The shape of each painted text body, for the soft edge of a group that holds it. `paintItem`
 * notes each body with `ownText`, and `softEdged` reads it to clip the text that a later body
 * covers (`groupTextCovers`).
 */
const textOwners = new WeakMap<Element, LayoutItem>();

/** Note that `item` painted the text body `node`. Returns `node`. @internal */
export function ownText(item: LayoutItem, node: SVGElement): SVGElement {
  textOwners.set(node, item);
  return node;
}

/**
 * Clip the text body `body` to each of `paths` (`groupTextCovers`), in the slide space of its
 * parent: one group with a clip path for each, around the body in its place.
 */
function clipText(
  body: Element,
  paths: readonly (readonly PathCommand[])[],
  ctx: PaintContext
): void {
  const { doc } = ctx.defs;
  const parent = body.parentNode;
  if (!parent) return;
  const next = body.nextSibling;
  let node = body as SVGElement;
  for (const path of paths) {
    const id = ctx.defs.id('cover');
    ctx.defs.add(
      svg(doc, 'clipPath', { id, clipPathUnits: 'userSpaceOnUse' }, [
        svg(doc, 'path', { d: commandData(path), 'clip-rule': 'evenodd' }),
      ])
    );
    node = svg(doc, 'g', { 'clip-path': `url(#${id})` }, [node]);
  }
  parent.insertBefore(node, next);
}

/**
 * The soft edge of a group (`filter`) over its painted children (`inner`), with the text of its
 * shapes painted once, unfaded, over the faded picture. The filter takes a `<use>` copy of the
 * children; in the copy, the text of shapes is hidden (`COPY_TEXT`, inherited from the copy).
 * The children themselves keep their nodes and their text, with everything but that text hidden,
 * so the text stays in its shape's node and paints once at its own alpha. A table's text fades
 * with the picture, as the PDF export draws it. A soft-edged group inside the children hides its
 * own text in the copy in the same way. Text that a later body in the group covers paints clipped
 * to the outside of that body (`groupTextCovers`), so the body still shows over it.
 */
function softEdged(
  doc: Document,
  group: Extract<LayoutItem, { kind: 'group' }>,
  inner: SVGElement,
  filter: string,
  ctx: PaintContext
): SVGElement {
  const { clips } = groupTextCovers(group);
  for (const body of inner.querySelectorAll('[data-pptx-text]')) {
    if (body.closest('[data-kind="table"]')) continue;
    const style = body.getAttribute('style') ?? '';
    if (!style.includes(COPY_TEXT))
      body.setAttribute('style', `${style};visibility:var(${COPY_TEXT},visible)`);
    const owner = textOwners.get(body);
    const paths = owner && clips.get(owner);
    if (paths) clipText(body, paths, ctx);
  }
  const innerId = ctx.defs.id('soft');
  inner.setAttribute('id', innerId);
  return svg(doc, 'g', {}, [
    svg(doc, 'g', { filter: `url(#${filter})` }, [
      svg(doc, 'use', { href: `#${innerId}`, style: `${COPY_TEXT}:hidden` }),
    ]),
    svg(doc, 'g', { visibility: 'hidden' }, [inner]),
  ]);
}
