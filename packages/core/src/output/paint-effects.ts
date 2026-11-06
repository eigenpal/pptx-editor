// Shape effects as SVG (design D1 of `effects-fidelity`).
//
// A shape paints as one content group (geometry and text). An outer shadow is a `<use>` copy of
// that group under the effect transform, inside an untransformed group that holds the filter, so
// blur and color always work in slide space.
//
// Glow, inner shadow, reflection, and soft edge come from the installed effects renderer
// (`drawing-renderers.ts`, `@pptx-editor.dev/pro/drawing`) as plain data, which `effectDefs`
// and `effectElement` build here. Without the renderer, a shape paints without them.
//
// Approximations are listed on the public fidelity page (`docs/site/content/pptx-fidelity.mdx`)
// and in the harness report: the blur kernel (`BLUR_SIGMA_PER_RADIUS`, measured).

import { toCss, type ResolvedEffects } from '../cascade/index.ts';
import { apply, type EffectCopySource, type EffectNode, type Matrix } from '../layout/index.ts';
import { blurSigma, boundsOf, effectMatrix, grow, PX, type Bounds } from './effect-transforms.ts';
import { matrixAttr, num, svg, type Defs } from './svg.ts';

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
