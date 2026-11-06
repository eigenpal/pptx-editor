// The `output` lane: the painter over layout records. Browser-only (it builds SVG DOM).
// Never re-derives geometry or resolves properties.

export { paintSlide, pathData, type AltTextPort, type PaintOptions } from './paint.ts';
export type { ImageSizePort, ImageUrlPort, PaintContext } from './paint-fill.ts';
export { STORED_PIXELS_STYLE } from './paint-fill.ts';
// Internal: the pure rules that `pptx-to-pdf` shares with this paint.
export {
  adjustForPathMode,
  blurSigma,
  circleGradientGeometry,
  effectMatrix,
  glowSpread,
  gradientPaintStops,
  innerShadowShift,
  linearGradientEnds,
  lumTransfer,
} from './painter-math.ts';
export type { SpreadBlur } from './painter-math.ts';
export { registerFontFaces, type FontFaceHandle } from './font-faces.ts';
export { SVG_NS } from './svg.ts';
// Internal: the DOM-free document that a host without a DOM paints into.
export { createSvgDocument, elementsWithAttribute, serializeSvgElement } from './svg-document.ts';
