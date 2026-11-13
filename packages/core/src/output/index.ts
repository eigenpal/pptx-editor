// The `output` lane: the painter over layout records. Browser-only (it builds SVG DOM).
// Never re-derives geometry or resolves properties.

export {
  paintSlide,
  MAX_OVERLAY_PATHS,
  paintWeight,
  pathData,
  type AltTextPort,
  type PaintOptions,
  type ViewRect,
} from './paint.ts';
export { SlidePainter } from './slide-painter.ts';
export type { ImageSizePort, ImageUrlPort, PaintContext } from './paint-fill.ts';
export { STORED_PIXELS_STYLE } from './paint-fill.ts';
export { metafilePng, MAX_METAFILE_RASTER_EDGE } from './metafile-png.ts';
// Internal: the pure rules that `pptx-to-pdf` and `pro/drawing` share with this paint.
export {
  adjustForPathMode,
  applyBlipColorSteps,
  blipColorSteps,
  blurSigma,
  circleGradientGeometry,
  effectBoxCorners,
  effectMatrix,
  finiteFacet,
  gradientPaintStops,
  linearGradientEnds,
  LUMA_WEIGHTS,
  lumTransfer,
  rectGradientGeometry,
  shapeBandBudget,
  shapeGradientBands,
  threeDColor,
  threeDFaceItem,
  tileGeometry,
  type BlipColorStep,
  type TileGeometry,
} from './painter-math.ts';
export { groupTextCovers, type GroupTextCovers } from './group-text-covers.ts';
export { PATTERN_PIXEL, patternBitmap } from './preset-patterns.ts';
export {
  compoundBandLines,
  compoundPlan,
  type Band as CompoundBand,
  type CompoundPlan,
  type CompoundBandLine,
  type Point as CompoundPoint,
  type Polyline as CompoundPolyline,
} from './compound-lines.ts';
export {
  registerFontFaces,
  type FontFaceHandle,
  type RegisterFontFacesOptions,
} from './font-faces.ts';
export { SVG_NS } from './svg.ts';
// Internal: the transform format that `pro/drawing` writes, to the same precision.
export { matrixAttr } from './svg.ts';
// Internal: the DOM-free document that a host without a DOM paints into.
export { createSvgDocument, elementsWithAttribute, serializeSvgElement } from './svg-document.ts';
