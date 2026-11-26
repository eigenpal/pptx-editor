/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
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
export type { CreateBrowserOptions } from './browser.ts';
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
  ShapeArrowheadStyle,
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
  ParagraphLineSpacingRule,
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
export {
  Comment,
  CommentCollection,
  CommentReply,
  CommentReplyCollection,
  type CommentAddOptions,
  type CommentAnchor,
  type CommentAuthorOptions,
} from '../model/comments.ts';
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
export type { SearchOptions } from '../model/search.ts';
export { Section, SectionCollection, type SectionDeleteOptions } from '../model/sections.ts';
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
export {
  BulletFormat,
  ParagraphFormat,
  ShapeFont,
  TextFrame,
  TextRange,
  type TextFrameFit,
} from '../model/text.ts';
export { TextRangeCollection } from '../model/text-ranges.ts';
export { SlideHeaderFooter, type HeaderFooterSettings } from '../model/header-footer.ts';

export { SlideTransitionDirection, SlideTransitionEffect } from '../model/transition-enums.ts';
export { SlideTransition, type SlideTransitionOptions } from '../model/transitions.ts';

export type {
  LineEndType,
  ShapeFormatPatch,
  ShapeFormatProperties,
  GradientDetail,
  GradientStopDetail,
  ShadowDetail,
  ReflectionDetail,
} from '@pptx-editor.dev/core/contracts';
export {
  Chart,
  ChartReadOnlyReason,
  ChartType,
  ChartWorkbookState,
  type ChartData,
  type ChartSeriesData,
} from '../model/charts.ts';
