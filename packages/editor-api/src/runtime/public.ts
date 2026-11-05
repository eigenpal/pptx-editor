/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public surface both entries share: the runtime, the request context, the error model,
// and the presentation object model. Nothing here reaches the DOM.

export { ClientObject } from './client-object.ts';
export { ClientResult } from './client-result.ts';
export {
  isPptxEditorError,
  PptxEditorError,
  type PptxEditorDebugInfo,
  type PptxEditorErrorCode,
  type PptxEditorErrorInit,
  type PptxEditorInnerError,
} from './errors.ts';
export type { LoadOption, LoadPropertyOptions, LoadQueryOptions } from './load-options.ts';
export {
  RequestContext,
  type ClientRequestContext,
  type RequestContextDebugInfo,
} from './request-context.ts';
export type {
  PptxEditorRuntime,
  PptxEditorServerRuntime,
  PresentationCapabilities,
  RunCallback,
} from './runtime.ts';
export type { CreateServerOptions } from './server.ts';
export { TrackedObjects } from './tracked-objects.ts';

export {
  BindingType,
  BulletStyle,
  BulletType,
  ConnectorType,
  DocumentPropertyType,
  HyperlinkType,
  ParagraphHorizontalAlignment,
  PlaceholderType,
  ShapeAutoSize,
  ShapeFillType,
  ShapeFontUnderlineStyle,
  ShapeLineDashStyle,
  ShapeLineStyle,
  ShapeType,
  ShapeZOrder,
  SlideBackgroundFillType,
  SlideBackgroundGradientFillType,
  SlideBackgroundPatternFillType,
  InsertSlideFormatting,
  ShapeGetImageFormatType,
  SlideLayoutType,
  TableStyle,
  TextVerticalAlignment,
  ThemeColor,
} from '../model/enums.ts';
export { Adjustments } from '../model/adjustments.ts';
export { Application } from '../model/application.ts';
export { Binding, BindingCollection } from '../model/bindings.ts';
export {
  SlideBackground,
  SlideBackgroundFill,
  SlideBackgroundGradientFill,
  SlideBackgroundPatternFill,
  SlideBackgroundPictureOrTextureFill,
  SlideBackgroundSolidFill,
  SlideLayoutBackground,
  SlideMasterBackground,
  type SlideBackgroundGradientFillOptions,
  type SlideBackgroundPatternFillOptions,
  type SlideBackgroundPictureOrTextureFillOptions,
  type SlideBackgroundSolidFillOptions,
} from '../model/backgrounds.ts';
export { CustomProperty, CustomPropertyCollection } from '../model/custom-properties.ts';
export {
  CustomXmlPart,
  CustomXmlPartCollection,
  CustomXmlPartScopedCollection,
} from '../model/custom-xml.ts';
export { ShapeFill, ShapeLineFormat } from '../model/fill-line.ts';
export type { ShapeGetImageOptions, SlideGetImageOptions } from '../model/images.ts';
export { GeometricShapeType } from '../model/geometry.ts';
export {
  Hyperlink,
  HyperlinkCollection,
  HyperlinkScopedCollection,
  type HyperlinkAddOptions,
} from '../model/hyperlinks.ts';
export {
  DocumentProperties,
  PageSetup,
  Presentation,
  type InsertSlideOptions,
} from '../model/presentation.ts';
export {
  PlaceholderFormat,
  Shape,
  ShapeCollection,
  ShapeGroup,
  ShapeScopedCollection,
  type ShapeAddOptions,
} from '../model/shapes.ts';
export {
  Slide,
  SlideCollection,
  SlideLayout,
  SlideLayoutCollection,
  SlideMaster,
  SlideMasterCollection,
  SlideScopedCollection,
  type AddSlideOptions,
} from '../model/slides.ts';
export {
  Table,
  TableCellCollection,
  TableColumn,
  TableColumnCollection,
  TableRow,
  TableRowCollection,
  type TableClearOptions,
} from '../model/tables.ts';
export type {
  BorderProperties,
  FillProperties,
  TableAddOptions,
  TableCellBorders,
  TableCellMargins,
  TableCellProperties,
  TableColumnProperties,
  TableMergedAreaProperties,
  TableRowProperties,
} from '../model/table-options.ts';
export { TableCell, type FontProperties, type TextRun } from '../model/table-cell.ts';
export { Border, Borders, Margins, TableStyleSettings } from '../model/table-format.ts';
export { Tag, TagCollection } from '../model/tags.ts';
export { ThemeColorScheme } from '../model/theme.ts';
export { BulletFormat, ParagraphFormat, ShapeFont, TextFrame, TextRange } from '../model/text.ts';
