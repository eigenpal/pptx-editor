// The picture presets of the Picture Format tab, as the DrawingML each one writes: the
// Recolor presets of the Color gallery (ECMA-376 Part 1, 20.1.8 blip effects `a:grayscl`,
// `a:duotone`, `a:biLevel`) and the Picture Styles gallery (`a:ln`, `a:effectLst`, and the
// geometry of `p:spPr`). The values are the preset definitions in `reference/quick-ref/drawingml.md`.
// They are not yet checked against reference PDFs (listed as an approximation on `docs/site/content/pptx-fidelity.mdx`).
//
// A preset is matched back from a file by a structural signature, so the chrome can show the
// current choice of a picture that another app styled.

import { NS, specOf, type Element, type NodeSpec } from '../store/index.ts';
import type { ImageRecolor, ImageStyleId } from '../contracts/index.ts';
import { specFromXml } from '../automation/commands/shape-commands.ts';

/** Every blip effect of CT_Blip (ECMA-376 Part 1, 20.1.8.13), the extension list excluded. */
export const BLIP_EFFECTS: ReadonlySet<string> = new Set([
  'alphaBiLevel',
  'alphaCeiling',
  'alphaFloor',
  'alphaInv',
  'alphaMod',
  'alphaModFix',
  'alphaRepl',
  'biLevel',
  'blur',
  'clrChange',
  'clrRepl',
  'duotone',
  'fillOverlay',
  'grayscl',
  'hsl',
  'lum',
  'tint',
]);

/** The blip effects a Recolor preset writes. */
export const RECOLOR_EFFECTS: ReadonlySet<string> = new Set(['grayscl', 'duotone', 'biLevel']);

/** The `a:ext/@uri` of the 2010 picture layer (artistic effects, `a14:imgProps`). */
export const IMG_PROPS_EXT = '{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}';
/** The `a:ext/@uri` of an SVG alternative (`asvg:svgBlip`). */
export const SVG_BLIP_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

const ACCENTS = [1, 2, 3, 4, 5, 6] as const;

/** The Recolor presets in gallery order. */
export const RECOLOR_PRESETS: readonly ImageRecolor[] = [
  'none',
  'grayscale',
  'sepia',
  'blackWhite25',
  'blackWhite50',
  'blackWhite75',
  ...ACCENTS.map((n) => `accent${n}Dark` as const),
  ...ACCENTS.map((n) => `accent${n}Light` as const),
];

/** The blip effect a Recolor preset writes, or null for `none`. */
export function recolorXml(recolor: ImageRecolor): string | null {
  switch (recolor) {
    case 'none':
      return null;
    case 'grayscale':
      return '<a:grayscl/>';
    case 'sepia':
      return '<a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone>';
    case 'blackWhite25':
      return '<a:biLevel thresh="25000"/>';
    case 'blackWhite50':
      return '<a:biLevel thresh="50000"/>';
    case 'blackWhite75':
      return '<a:biLevel thresh="75000"/>';
  }
  const m = /^accent([1-6])(Dark|Light)$/.exec(recolor);
  if (!m) return null;
  const accent = `accent${m[1]}`;
  // Dark variations map black to a shade of the accent; light ones map white to a tint of it.
  return m[2] === 'Dark'
    ? `<a:duotone><a:schemeClr val="${accent}"><a:shade val="45000"/><a:satMod val="135000"/></a:schemeClr><a:prstClr val="white"/></a:duotone>`
    : `<a:duotone><a:prstClr val="black"/><a:schemeClr val="${accent}"><a:tint val="45000"/><a:satMod val="400000"/></a:schemeClr></a:duotone>`;
}

/** What a picture style writes into `p:spPr`. */
export interface PictureStyleXml {
  /** The geometry element (`a:prstGeom`). */
  readonly geometry: string;
  /** The outline, or null for none. */
  readonly ln: string | null;
  /** The effect list, or null for none. */
  readonly effectLst: string | null;
}

const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
const FRAME_SHADOW =
  '<a:effectLst><a:outerShdw blurRad="55000" dist="18000" dir="5400000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>';
const frame = (color: string): string =>
  `<a:ln w="88900" cap="sq"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:miter lim="800000"/></a:ln>`;

/** The picture styles, in the relative order of the gallery. */
export const PICTURE_STYLES: Readonly<Record<ImageStyleId, PictureStyleXml>> = {
  simpleFrameWhite: { geometry: RECT, ln: frame('FFFFFF'), effectLst: FRAME_SHADOW },
  reflectedRoundedRectangle: {
    geometry:
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8594"/></a:avLst></a:prstGeom>',
    ln: null,
    effectLst:
      '<a:effectLst><a:reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>',
  },
  softEdgeRectangle: {
    geometry: RECT,
    ln: null,
    effectLst: '<a:effectLst><a:softEdge rad="112500"/></a:effectLst>',
  },
  simpleFrameBlack: { geometry: RECT, ln: frame('000000'), effectLst: FRAME_SHADOW },
  softEdgeOval: {
    geometry: '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
    ln: null,
    effectLst: '<a:effectLst><a:softEdge rad="112500"/></a:effectLst>',
  },
};

/** The picture style ids, in gallery order. */
export const PICTURE_STYLE_IDS = Object.keys(PICTURE_STYLES) as ImageStyleId[];

/**
 * A structural signature of an element: expanded names, attributes in name order, and child
 * elements. Prefixes, namespace declarations, and text are left out, so a preset matches the
 * same XML whatever prefixes a file binds.
 */
export function signature(spec: NodeSpec): string {
  const attrs = (spec.attrs ?? [])
    .map((a) => `${a.ns}|${a.local}=${a.value}`)
    .sort()
    .join(',');
  const kids = (spec.children ?? [])
    .filter((c): c is NodeSpec => 'local' in c)
    .map(signature)
    .join('');
  return `<${spec.ns}|${spec.local}(${attrs})${kids}>`;
}

const recolorSignatures = new Map<string, ImageRecolor>();
function recolorBySignature(): ReadonlyMap<string, ImageRecolor> {
  if (recolorSignatures.size === 0)
    for (const preset of RECOLOR_PRESETS) {
      const xml = recolorXml(preset);
      if (xml) recolorSignatures.set(signature(specFromXml(xml)), preset);
    }
  return recolorSignatures;
}

/** The Recolor preset of a blip's recolor effects: `none`, a preset, or `other`. */
export function recolorOf(blip: Element | undefined): ImageRecolor | 'other' {
  const effects = (blip?.children ?? []).filter(
    (c): c is Element => c.type === 'element' && c.ns === NS.a && RECOLOR_EFFECTS.has(c.local)
  );
  if (effects.length === 0) return 'none';
  if (effects.length > 1) return 'other';
  return recolorBySignature().get(signature(specOf(effects[0]))) ?? 'other';
}

let styleSignatures: Map<string, ImageStyleId> | null = null;

/** The signature of what a style writes: geometry, outline, and effects. */
function styleSignature(geometry: NodeSpec | undefined, ln?: NodeSpec, fx?: NodeSpec): string {
  return [geometry, ln, fx].map((s) => (s ? signature(s) : '-')).join('');
}

/** The picture style a `p:spPr` matches, if any. */
export function pictureStyleOf(spPr: Element | undefined): ImageStyleId | undefined {
  if (!spPr) return undefined;
  if (!styleSignatures) {
    styleSignatures = new Map();
    for (const id of PICTURE_STYLE_IDS) {
      const s = PICTURE_STYLES[id];
      const parse = (xml: string | null): NodeSpec | undefined =>
        xml ? specFromXml(xml) : undefined;
      styleSignatures.set(styleSignature(parse(s.geometry), parse(s.ln), parse(s.effectLst)), id);
    }
  }
  const of = (local: string): NodeSpec | undefined => {
    const el = spPr.children.find(
      (c): c is Element => c.type === 'element' && c.ns === NS.a && c.local === local
    );
    return el ? specOf(el) : undefined;
  };
  return styleSignatures.get(styleSignature(of('prstGeom'), of('ln'), of('effectLst')));
}
