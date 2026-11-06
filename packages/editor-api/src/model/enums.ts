/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The Office.js enum values of the presentation object model, as string literal types, and
// their ECMA-376 values. The automation protocol speaks ECMA-376 (`ST_TextAnchoringType`,
// `ST_PresetLineDashVal`, …); this module is the only place that translates.
//
// Each map is a null-prototype record, so an enum value from a consumer can never select an
// inherited key.

import { failArgument, nearest, oneOf } from '../runtime/errors.ts';

/** @public */
export enum ShapeType {
  unsupported = 'Unsupported',
  image = 'Image',
  geometricShape = 'GeometricShape',
  group = 'Group',
  line = 'Line',
  table = 'Table',
  callout = 'Callout',
  chart = 'Chart',
  contentApp = 'ContentApp',
  diagram = 'Diagram',
  freeform = 'Freeform',
  graphic = 'Graphic',
  ink = 'Ink',
  media = 'Media',
  model3D = 'Model3D',
  ole = 'Ole',
  placeholder = 'Placeholder',
  textBox = 'TextBox',
}

/** @public */
export enum BindingType {
  shape = 'Shape',
}

/** @public */
export enum ConnectorType {
  straight = 'Straight',
  elbow = 'Elbow',
  curve = 'Curve',
}

/** @public */
export enum HyperlinkType {
  textRange = 'TextRange',
  shape = 'Shape',
}

/** @public */
export enum ShapeZOrder {
  bringForward = 'BringForward',
  bringToFront = 'BringToFront',
  sendBackward = 'SendBackward',
  sendToBack = 'SendToBack',
}

/** @public */
export enum ShapeFillType {
  noFill = 'NoFill',
  solid = 'Solid',
  gradient = 'Gradient',
  pattern = 'Pattern',
  pictureAndTexture = 'PictureAndTexture',
  slideBackground = 'SlideBackground',
}

/** @public */
export enum ShapeLineDashStyle {
  dash = 'Dash',
  dashDot = 'DashDot',
  dashDotDot = 'DashDotDot',
  longDash = 'LongDash',
  longDashDot = 'LongDashDot',
  roundDot = 'RoundDot',
  solid = 'Solid',
  squareDot = 'SquareDot',
  longDashDotDot = 'LongDashDotDot',
  systemDash = 'SystemDash',
  systemDot = 'SystemDot',
  systemDashDot = 'SystemDashDot',
}

/** @public */
export enum ShapeLineStyle {
  single = 'Single',
  thickBetweenThin = 'ThickBetweenThin',
  thickThin = 'ThickThin',
  thinThick = 'ThinThick',
  thinThin = 'ThinThin',
}

/** @public */
export enum ShapeAutoSize {
  autoSizeNone = 'AutoSizeNone',
  autoSizeTextToFitShape = 'AutoSizeTextToFitShape',
  autoSizeShapeToFitText = 'AutoSizeShapeToFitText',
  autoSizeMixed = 'AutoSizeMixed',
}

/** @public */
export enum TextVerticalAlignment {
  top = 'Top',
  middle = 'Middle',
  bottom = 'Bottom',
  topCentered = 'TopCentered',
  middleCentered = 'MiddleCentered',
  bottomCentered = 'BottomCentered',
}

/** @public */
export enum ParagraphHorizontalAlignment {
  left = 'Left',
  center = 'Center',
  right = 'Right',
  justify = 'Justify',
  justifyLow = 'JustifyLow',
  distributed = 'Distributed',
  thaiDistributed = 'ThaiDistributed',
}

/** @public */
export enum ShapeFontUnderlineStyle {
  none = 'None',
  single = 'Single',
  double = 'Double',
  heavy = 'Heavy',
  dotted = 'Dotted',
  dottedHeavy = 'DottedHeavy',
  dash = 'Dash',
  dashHeavy = 'DashHeavy',
  dashLong = 'DashLong',
  dashLongHeavy = 'DashLongHeavy',
  dotDash = 'DotDash',
  dotDashHeavy = 'DotDashHeavy',
  dotDotDash = 'DotDotDash',
  dotDotDashHeavy = 'DotDotDashHeavy',
  wavy = 'Wavy',
  wavyHeavy = 'WavyHeavy',
  wavyDouble = 'WavyDouble',
}

/** @public */
export enum PlaceholderType {
  unsupported = 'Unsupported',
  date = 'Date',
  slideNumber = 'SlideNumber',
  footer = 'Footer',
  header = 'Header',
  title = 'Title',
  body = 'Body',
  centerTitle = 'CenterTitle',
  subtitle = 'Subtitle',
  verticalTitle = 'VerticalTitle',
  verticalBody = 'VerticalBody',
  content = 'Content',
  chart = 'Chart',
  table = 'Table',
  onlinePicture = 'OnlinePicture',
  media = 'Media',
  verticalContent = 'VerticalContent',
  picture = 'Picture',
  cameo = 'Cameo',
}

/** @public */
export enum SlideLayoutType {
  blank = 'Blank',
  chart = 'Chart',
  chartAndText = 'ChartAndText',
  clipArtAndText = 'ClipArtAndText',
  clipArtAndVerticalText = 'ClipArtAndVerticalText',
  comparison = 'Comparison',
  contentWithCaption = 'ContentWithCaption',
  custom = 'Custom',
  fourObjects = 'FourObjects',
  largeObject = 'LargeObject',
  mediaClipAndText = 'MediaClipAndText',
  mixed = 'Mixed',
  object = 'Object',
  objectAndText = 'ObjectAndText',
  objectAndTwoObjects = 'ObjectAndTwoObjects',
  objectOverText = 'ObjectOverText',
  organizationChart = 'OrganizationChart',
  pictureWithCaption = 'PictureWithCaption',
  sectionHeader = 'SectionHeader',
  table = 'Table',
  text = 'Text',
  textAndChart = 'TextAndChart',
  textAndClipArt = 'TextAndClipArt',
  textAndMediaClip = 'TextAndMediaClip',
  textAndObject = 'TextAndObject',
  textAndTwoObjects = 'TextAndTwoObjects',
  textOverObject = 'TextOverObject',
  title = 'Title',
  titleOnly = 'TitleOnly',
  twoColumnText = 'TwoColumnText',
  twoObjects = 'TwoObjects',
  twoObjectsAndObject = 'TwoObjectsAndObject',
  twoObjectsAndText = 'TwoObjectsAndText',
  twoObjectsOverText = 'TwoObjectsOverText',
  verticalText = 'VerticalText',
  verticalTitleAndText = 'VerticalTitleAndText',
  verticalTitleAndTextOverChart = 'VerticalTitleAndTextOverChart',
}

/** @public */
export enum BulletType {
  unsupported = 'Unsupported',
  none = 'None',
  numbered = 'Numbered',
  unnumbered = 'Unnumbered',
}

/** @public */
export enum BulletStyle {
  unsupported = 'Unsupported',
  alphabetLowercaseParenthesesBoth = 'AlphabetLowercaseParenthesesBoth',
  alphabetUppercaseParenthesesBoth = 'AlphabetUppercaseParenthesesBoth',
  alphabetLowercaseParenthesisRight = 'AlphabetLowercaseParenthesisRight',
  alphabetUppercaseParenthesisRight = 'AlphabetUppercaseParenthesisRight',
  alphabetLowercasePeriod = 'AlphabetLowercasePeriod',
  alphabetUppercasePeriod = 'AlphabetUppercasePeriod',
  arabicNumeralParenthesesBoth = 'ArabicNumeralParenthesesBoth',
  arabicNumeralParenthesisRight = 'ArabicNumeralParenthesisRight',
  arabicNumeralPeriod = 'ArabicNumeralPeriod',
  arabicNumeralPlain = 'ArabicNumeralPlain',
  romanLowercaseParenthesesBoth = 'RomanLowercaseParenthesesBoth',
  romanUppercaseParenthesesBoth = 'RomanUppercaseParenthesesBoth',
  romanLowercaseParenthesisRight = 'RomanLowercaseParenthesisRight',
  romanUppercaseParenthesisRight = 'RomanUppercaseParenthesisRight',
  romanLowercasePeriod = 'RomanLowercasePeriod',
  romanUppercasePeriod = 'RomanUppercasePeriod',
  circleNumberDoubleBytePlain = 'CircleNumberDoubleBytePlain',
  circleNumberWideDoubleByteBlackPlain = 'CircleNumberWideDoubleByteBlackPlain',
  circleNumberWideDoubleByteWhitePlain = 'CircleNumberWideDoubleByteWhitePlain',
  arabicDoubleBytePeriod = 'ArabicDoubleBytePeriod',
  arabicDoubleBytePlain = 'ArabicDoubleBytePlain',
  simplifiedChinesePeriod = 'SimplifiedChinesePeriod',
  simplifiedChinesePlain = 'SimplifiedChinesePlain',
  traditionalChinesePeriod = 'TraditionalChinesePeriod',
  traditionalChinesePlain = 'TraditionalChinesePlain',
  kanjiSimplifiedChineseDoubleBytePeriod = 'KanjiSimplifiedChineseDoubleBytePeriod',
  kanjiKoreanPlain = 'KanjiKoreanPlain',
  kanjiKoreanPeriod = 'KanjiKoreanPeriod',
  arabicAlphabetDash = 'ArabicAlphabetDash',
  arabicAbjadDash = 'ArabicAbjadDash',
  hebrewAlphabetDash = 'HebrewAlphabetDash',
  thaiAlphabetPeriod = 'ThaiAlphabetPeriod',
  thaiAlphabetParenthesisRight = 'ThaiAlphabetParenthesisRight',
  thaiAlphabetParenthesesBoth = 'ThaiAlphabetParenthesesBoth',
  thaiNumeralPeriod = 'ThaiNumeralPeriod',
  thaiNumeralParenthesisRight = 'ThaiNumeralParenthesisRight',
  thaiNumeralParenthesesBoth = 'ThaiNumeralParenthesesBoth',
  hindiAlphabetPeriod = 'HindiAlphabetPeriod',
  hindiNumeralPeriod = 'HindiNumeralPeriod',
  hindiNumeralParenthesisRight = 'HindiNumeralParenthesisRight',
  hindiAlphabet1Period = 'HindiAlphabet1Period',
}

/** @public */
export enum SlideBackgroundFillType {
  gradient = 'Gradient',
  pattern = 'Pattern',
  pictureOrTexture = 'PictureOrTexture',
  solid = 'Solid',
  unsupported = 'Unsupported',
}

/** @public */
export enum SlideBackgroundGradientFillType {
  linear = 'Linear',
  path = 'Path',
  radial = 'Radial',
  rectangular = 'Rectangular',
  shadeFromTitle = 'ShadeFromTitle',
  unsupported = 'Unsupported',
}

/** @public */
export enum SlideBackgroundPatternFillType {
  cross = 'Cross',
  darkDownwardDiagonal = 'DarkDownwardDiagonal',
  darkHorizontal = 'DarkHorizontal',
  darkUpwardDiagonal = 'DarkUpwardDiagonal',
  darkVertical = 'DarkVertical',
  dashedDownwardDiagonal = 'DashedDownwardDiagonal',
  dashedHorizontal = 'DashedHorizontal',
  dashedUpwardDiagonal = 'DashedUpwardDiagonal',
  dashedVertical = 'DashedVertical',
  diagonalBrick = 'DiagonalBrick',
  diagonalCross = 'DiagonalCross',
  divot = 'Divot',
  dottedDiamond = 'DottedDiamond',
  dottedGrid = 'DottedGrid',
  downwardDiagonal = 'DownwardDiagonal',
  horizontal = 'Horizontal',
  horizontalBrick = 'HorizontalBrick',
  largeCheckerBoard = 'LargeCheckerBoard',
  largeConfetti = 'LargeConfetti',
  largeGrid = 'LargeGrid',
  lightDownwardDiagonal = 'LightDownwardDiagonal',
  lightHorizontal = 'LightHorizontal',
  lightUpwardDiagonal = 'LightUpwardDiagonal',
  lightVertical = 'LightVertical',
  narrowHorizontal = 'NarrowHorizontal',
  narrowVertical = 'NarrowVertical',
  outlinedDiamond = 'OutlinedDiamond',
  percent10 = 'Percent10',
  percent20 = 'Percent20',
  percent25 = 'Percent25',
  percent30 = 'Percent30',
  percent40 = 'Percent40',
  percent5 = 'Percent5',
  percent50 = 'Percent50',
  percent60 = 'Percent60',
  percent70 = 'Percent70',
  percent75 = 'Percent75',
  percent80 = 'Percent80',
  percent90 = 'Percent90',
  plaid = 'Plaid',
  shingle = 'Shingle',
  smallCheckerBoard = 'SmallCheckerBoard',
  smallConfetti = 'SmallConfetti',
  smallGrid = 'SmallGrid',
  solidDiamond = 'SolidDiamond',
  sphere = 'Sphere',
  trellis = 'Trellis',
  unsupported = 'Unsupported',
  upwardDiagonal = 'UpwardDiagonal',
  vertical = 'Vertical',
  wave = 'Wave',
  weave = 'Weave',
  wideDownwardDiagonal = 'WideDownwardDiagonal',
  wideUpwardDiagonal = 'WideUpwardDiagonal',
  zigZag = 'ZigZag',
}

/** @public */
export enum ThemeColor {
  accent1 = 'Accent1',
  accent2 = 'Accent2',
  accent3 = 'Accent3',
  accent4 = 'Accent4',
  accent5 = 'Accent5',
  accent6 = 'Accent6',
  dark1 = 'Dark1',
  dark2 = 'Dark2',
  followedHyperlink = 'FollowedHyperlink',
  hyperlink = 'Hyperlink',
  light1 = 'Light1',
  light2 = 'Light2',
  mixed = 'Mixed',
  none = 'None',
}

/** @public */
export enum TableStyle {
  darkStyle1 = 'DarkStyle1',
  darkStyle1Accent1 = 'DarkStyle1Accent1',
  darkStyle1Accent2 = 'DarkStyle1Accent2',
  darkStyle1Accent3 = 'DarkStyle1Accent3',
  darkStyle1Accent4 = 'DarkStyle1Accent4',
  darkStyle1Accent5 = 'DarkStyle1Accent5',
  darkStyle1Accent6 = 'DarkStyle1Accent6',
  darkStyle2 = 'DarkStyle2',
  darkStyle2Accent1 = 'DarkStyle2Accent1',
  darkStyle2Accent2 = 'DarkStyle2Accent2',
  darkStyle2Accent3 = 'DarkStyle2Accent3',
  lightStyle1 = 'LightStyle1',
  lightStyle1Accent1 = 'LightStyle1Accent1',
  lightStyle1Accent2 = 'LightStyle1Accent2',
  lightStyle1Accent3 = 'LightStyle1Accent3',
  lightStyle1Accent4 = 'LightStyle1Accent4',
  lightStyle1Accent5 = 'LightStyle1Accent5',
  lightStyle1Accent6 = 'LightStyle1Accent6',
  lightStyle2 = 'LightStyle2',
  lightStyle2Accent1 = 'LightStyle2Accent1',
  lightStyle2Accent2 = 'LightStyle2Accent2',
  lightStyle2Accent3 = 'LightStyle2Accent3',
  lightStyle2Accent4 = 'LightStyle2Accent4',
  lightStyle2Accent5 = 'LightStyle2Accent5',
  lightStyle2Accent6 = 'LightStyle2Accent6',
  lightStyle3 = 'LightStyle3',
  lightStyle3Accent1 = 'LightStyle3Accent1',
  lightStyle3Accent2 = 'LightStyle3Accent2',
  lightStyle3Accent3 = 'LightStyle3Accent3',
  lightStyle3Accent4 = 'LightStyle3Accent4',
  lightStyle3Accent5 = 'LightStyle3Accent5',
  lightStyle3Accent6 = 'LightStyle3Accent6',
  mediumStyle1 = 'MediumStyle1',
  mediumStyle1Accent1 = 'MediumStyle1Accent1',
  mediumStyle1Accent2 = 'MediumStyle1Accent2',
  mediumStyle1Accent3 = 'MediumStyle1Accent3',
  mediumStyle1Accent4 = 'MediumStyle1Accent4',
  mediumStyle1Accent5 = 'MediumStyle1Accent5',
  mediumStyle1Accent6 = 'MediumStyle1Accent6',
  mediumStyle2 = 'MediumStyle2',
  mediumStyle2Accent1 = 'MediumStyle2Accent1',
  mediumStyle2Accent2 = 'MediumStyle2Accent2',
  mediumStyle2Accent3 = 'MediumStyle2Accent3',
  mediumStyle2Accent4 = 'MediumStyle2Accent4',
  mediumStyle2Accent5 = 'MediumStyle2Accent5',
  mediumStyle2Accent6 = 'MediumStyle2Accent6',
  mediumStyle3 = 'MediumStyle3',
  mediumStyle3Accent1 = 'MediumStyle3Accent1',
  mediumStyle3Accent2 = 'MediumStyle3Accent2',
  mediumStyle3Accent3 = 'MediumStyle3Accent3',
  mediumStyle3Accent4 = 'MediumStyle3Accent4',
  mediumStyle3Accent5 = 'MediumStyle3Accent5',
  mediumStyle3Accent6 = 'MediumStyle3Accent6',
  mediumStyle4 = 'MediumStyle4',
  mediumStyle4Accent1 = 'MediumStyle4Accent1',
  mediumStyle4Accent2 = 'MediumStyle4Accent2',
  mediumStyle4Accent3 = 'MediumStyle4Accent3',
  mediumStyle4Accent4 = 'MediumStyle4Accent4',
  mediumStyle4Accent5 = 'MediumStyle4Accent5',
  mediumStyle4Accent6 = 'MediumStyle4Accent6',
  noStyleNoGrid = 'NoStyleNoGrid',
  noStyleTableGrid = 'NoStyleTableGrid',
  themedStyle1Accent1 = 'ThemedStyle1Accent1',
  themedStyle1Accent2 = 'ThemedStyle1Accent2',
  themedStyle1Accent3 = 'ThemedStyle1Accent3',
  themedStyle1Accent4 = 'ThemedStyle1Accent4',
  themedStyle1Accent5 = 'ThemedStyle1Accent5',
  themedStyle1Accent6 = 'ThemedStyle1Accent6',
  themedStyle2Accent1 = 'ThemedStyle2Accent1',
  themedStyle2Accent2 = 'ThemedStyle2Accent2',
  themedStyle2Accent3 = 'ThemedStyle2Accent3',
  themedStyle2Accent4 = 'ThemedStyle2Accent4',
  themedStyle2Accent5 = 'ThemedStyle2Accent5',
  themedStyle2Accent6 = 'ThemedStyle2Accent6',
}

/** @public */
export enum InsertSlideFormatting {
  keepSourceFormatting = 'KeepSourceFormatting',
  useDestinationTheme = 'UseDestinationTheme',
}

/** The format of a shape picture from `Shape.getImageAsBase64`. @public */
export enum ShapeGetImageFormatType {
  png = 'Png',
}

/** The type of a custom document property value. @public */
export enum DocumentPropertyType {
  boolean = 'Boolean',
  date = 'Date',
  number = 'Number',
  string = 'String',
}

/** A two-way table between Office.js values and ECMA-376 values. */
export interface EnumMap<O extends string> {
  /** The ECMA-376 value of an Office.js value, or `InvalidArgument` naming `target`. */
  toFile(value: unknown, target: string): string;
  /** The Office.js value of an ECMA-376 value, or `fallback`. */
  fromFile(value: string | null | undefined): O | null;
}

function enumMap<O extends string>(
  pairs: readonly (readonly [O, string])[],
  options: { readonly writeOnly?: readonly (readonly [O, string])[]; readonly fallback?: O } = {}
): EnumMap<O> {
  const toFile: Record<string, string> = Object.create(null);
  const fromFile: Record<string, O> = Object.create(null);
  for (const [o, f] of pairs) {
    toFile[o] = f;
    fromFile[f] ??= o;
  }
  for (const [o, f] of options.writeOnly ?? []) toFile[o] = f;
  return {
    toFile(value, target) {
      if (typeof value !== 'string' || !(value in toFile))
        failArgument(target, oneOf(Object.keys(toFile)), nearest(value, Object.keys(toFile)));
      return toFile[value]!;
    },
    fromFile(value) {
      if (value === null || value === undefined) return null;
      return fromFile[value] ?? options.fallback ?? null;
    },
  };
}

/** `ST_TextAnchoringType` with `anchorCtr` (ECMA-376 Part 1, 21.1.2.1.1). */
export const VERTICAL_ALIGNMENT = enumMap<`${TextVerticalAlignment}`>([
  ['Top', 't'],
  ['Middle', 'ctr'],
  ['Bottom', 'b'],
  ['TopCentered', 'tCentered'],
  ['MiddleCentered', 'ctrCentered'],
  ['BottomCentered', 'bCentered'],
]);

/** `ST_TextAlignType` (ECMA-376 Part 1, 20.1.10.59). */
export const HORIZONTAL_ALIGNMENT = enumMap<`${ParagraphHorizontalAlignment}`>([
  ['Left', 'l'],
  ['Center', 'ctr'],
  ['Right', 'r'],
  ['Justify', 'just'],
  ['JustifyLow', 'justLow'],
  ['Distributed', 'dist'],
  ['ThaiDistributed', 'thaiDist'],
]);

/** The autofit choice of `a:bodyPr` (ECMA-376 Part 1, 21.1.2.1.1). */
export const AUTO_SIZE = enumMap<`${ShapeAutoSize}`>([
  ['AutoSizeNone', 'none'],
  ['AutoSizeTextToFitShape', 'normal'],
  ['AutoSizeShapeToFitText', 'shape'],
]);

/** `ST_TextUnderlineType` (ECMA-376 Part 1, 20.1.10.82). `words` reads as `Single`. */
export const UNDERLINE = enumMap<`${ShapeFontUnderlineStyle}`>([
  ['None', 'none'],
  ['Single', 'sng'],
  ['Single', 'words'],
  ['Double', 'dbl'],
  ['Heavy', 'heavy'],
  ['Dotted', 'dotted'],
  ['DottedHeavy', 'dottedHeavy'],
  ['Dash', 'dash'],
  ['DashHeavy', 'dashHeavy'],
  ['DashLong', 'dashLong'],
  ['DashLongHeavy', 'dashLongHeavy'],
  ['DotDash', 'dotDash'],
  ['DotDashHeavy', 'dotDashHeavy'],
  ['DotDotDash', 'dotDotDash'],
  ['DotDotDashHeavy', 'dotDotDashHeavy'],
  ['Wavy', 'wavy'],
  ['WavyHeavy', 'wavyHeavy'],
  ['WavyDouble', 'wavyDbl'],
]);

/**
 * `ST_PresetLineDashVal` (ECMA-376 Part 1, 20.1.10.48). The file has eleven dash presets and
 * the object model twelve: `SquareDot` writes `sysDot`, which reads back as `SystemDot`.
 */
export const DASH_STYLE = enumMap<`${ShapeLineDashStyle}`>(
  [
    ['Solid', 'solid'],
    ['RoundDot', 'dot'],
    ['Dash', 'dash'],
    ['LongDash', 'lgDash'],
    ['DashDot', 'dashDot'],
    ['LongDashDot', 'lgDashDot'],
    ['LongDashDotDot', 'lgDashDotDot'],
    ['SystemDash', 'sysDash'],
    ['SystemDot', 'sysDot'],
    ['SystemDashDot', 'sysDashDot'],
    ['DashDotDot', 'sysDashDotDot'],
  ],
  { writeOnly: [['SquareDot', 'sysDot']] }
);

/** `ST_CompoundLine` (ECMA-376 Part 1, 20.1.10.15). */
export const LINE_STYLE = enumMap<`${ShapeLineStyle}`>([
  ['Single', 'sng'],
  ['ThinThin', 'dbl'],
  ['ThickThin', 'thickThin'],
  ['ThinThick', 'thinThick'],
  ['ThickBetweenThin', 'tri'],
]);

/** `ST_PlaceholderType` (ECMA-376 Part 1, 19.7.10). A diagram placeholder reads as `Unsupported`. */
export const PLACEHOLDER_TYPE = enumMap<`${PlaceholderType}`>(
  [
    ['Title', 'title'],
    ['Body', 'body'],
    ['CenterTitle', 'ctrTitle'],
    ['Subtitle', 'subTitle'],
    ['Date', 'dt'],
    ['SlideNumber', 'sldNum'],
    ['Footer', 'ftr'],
    ['Header', 'hdr'],
    ['Content', 'obj'],
    ['Chart', 'chart'],
    ['Table', 'tbl'],
    ['OnlinePicture', 'clipArt'],
    ['Media', 'media'],
    ['Picture', 'pic'],
  ],
  { fallback: 'Unsupported' }
);

/** `ST_SlideLayoutType` (ECMA-376 Part 1, 19.7.15). */
export const LAYOUT_TYPE = enumMap<`${SlideLayoutType}`>(
  [
    ['Title', 'title'],
    ['Text', 'tx'],
    ['TwoColumnText', 'twoColTx'],
    ['Table', 'tbl'],
    ['TextAndChart', 'txAndChart'],
    ['ChartAndText', 'chartAndTx'],
    ['OrganizationChart', 'dgm'],
    ['Chart', 'chart'],
    ['TextAndClipArt', 'txAndClipArt'],
    ['ClipArtAndText', 'clipArtAndTx'],
    ['TitleOnly', 'titleOnly'],
    ['Blank', 'blank'],
    ['TextAndObject', 'txAndObj'],
    ['ObjectAndText', 'objAndTx'],
    ['LargeObject', 'objOnly'],
    ['Object', 'obj'],
    ['TextAndMediaClip', 'txAndMedia'],
    ['MediaClipAndText', 'mediaAndTx'],
    ['ObjectOverText', 'objOverTx'],
    ['TextOverObject', 'txOverObj'],
    ['TextAndTwoObjects', 'txAndTwoObj'],
    ['TwoObjectsAndText', 'twoObjAndTx'],
    ['TwoObjectsOverText', 'twoObjOverTx'],
    ['FourObjects', 'fourObj'],
    ['VerticalText', 'vertTx'],
    ['ClipArtAndVerticalText', 'clipArtAndVertTx'],
    ['VerticalTitleAndText', 'vertTitleAndTx'],
    ['VerticalTitleAndTextOverChart', 'vertTitleAndTxOverChart'],
    ['TwoObjects', 'twoObj'],
    ['ObjectAndTwoObjects', 'objAndTwoObj'],
    ['TwoObjectsAndObject', 'twoObjAndObj'],
    ['Custom', 'cust'],
    ['SectionHeader', 'secHead'],
    ['Comparison', 'twoTxTwoObj'],
    ['ContentWithCaption', 'objTx'],
    ['PictureWithCaption', 'picTx'],
  ],
  { fallback: 'Custom' }
);

/** The fill kinds of `a:spPr` (ECMA-376 Part 1, 20.1.8). */
export const FILL_TYPE = enumMap<`${ShapeFillType}`>([
  ['NoFill', 'noFill'],
  ['Solid', 'solid'],
  ['Gradient', 'gradient'],
  ['Pattern', 'pattern'],
  ['PictureAndTexture', 'pictureAndTexture'],
]);

/** The connector kinds of `addLine` (`p:cxnSp` presets, ECMA-376 Part 1, 20.1.10.56). */
export const CONNECTOR = enumMap<`${ConnectorType}`>([
  ['Straight', 'straight'],
  ['Elbow', 'elbow'],
  ['Curve', 'curve'],
]);

/** The z-order moves of `setZOrder`. */
export const Z_ORDER = enumMap<`${ShapeZOrder}`>([
  ['BringForward', 'bringForward'],
  ['BringToFront', 'bringToFront'],
  ['SendBackward', 'sendBackward'],
  ['SendToBack', 'sendToBack'],
]);

/** `ST_TextAutonumberScheme` (ECMA-376 Part 1, 20.1.10.61) for the bullet styles that have one. */
export const BULLET_STYLE = enumMap<`${BulletStyle}`>(
  [
    ['AlphabetLowercaseParenthesesBoth', 'alphaLcParenBoth'],
    ['AlphabetUppercaseParenthesesBoth', 'alphaUcParenBoth'],
    ['AlphabetLowercaseParenthesisRight', 'alphaLcParenR'],
    ['AlphabetUppercaseParenthesisRight', 'alphaUcParenR'],
    ['AlphabetLowercasePeriod', 'alphaLcPeriod'],
    ['AlphabetUppercasePeriod', 'alphaUcPeriod'],
    ['ArabicNumeralParenthesesBoth', 'arabicParenBoth'],
    ['ArabicNumeralParenthesisRight', 'arabicParenR'],
    ['ArabicNumeralPeriod', 'arabicPeriod'],
    ['ArabicNumeralPlain', 'arabicPlain'],
    ['RomanLowercaseParenthesesBoth', 'romanLcParenBoth'],
    ['RomanUppercaseParenthesesBoth', 'romanUcParenBoth'],
    ['RomanLowercaseParenthesisRight', 'romanLcParenR'],
    ['RomanUppercaseParenthesisRight', 'romanUcParenR'],
    ['RomanLowercasePeriod', 'romanLcPeriod'],
    ['RomanUppercasePeriod', 'romanUcPeriod'],
    ['CircleNumberDoubleBytePlain', 'circleNumDbPlain'],
    ['CircleNumberWideDoubleByteBlackPlain', 'circleNumWdBlackPlain'],
    ['CircleNumberWideDoubleByteWhitePlain', 'circleNumWdWhitePlain'],
    ['ArabicDoubleBytePeriod', 'arabicDbPeriod'],
    ['ArabicDoubleBytePlain', 'arabicDbPlain'],
    ['SimplifiedChinesePeriod', 'ea1ChsPeriod'],
    ['SimplifiedChinesePlain', 'ea1ChsPlain'],
    ['TraditionalChinesePeriod', 'ea1ChtPeriod'],
    ['TraditionalChinesePlain', 'ea1ChtPlain'],
    ['KanjiSimplifiedChineseDoubleBytePeriod', 'ea1JpnChsDbPeriod'],
    ['KanjiKoreanPlain', 'ea1JpnKorPlain'],
    ['KanjiKoreanPeriod', 'ea1JpnKorPeriod'],
    ['ArabicAlphabetDash', 'arabic1Minus'],
    ['ArabicAbjadDash', 'arabic2Minus'],
    ['HebrewAlphabetDash', 'hebrew2Minus'],
    ['ThaiAlphabetPeriod', 'thaiAlphaPeriod'],
    ['ThaiAlphabetParenthesisRight', 'thaiAlphaParenR'],
    ['ThaiAlphabetParenthesesBoth', 'thaiAlphaParenBoth'],
    ['ThaiNumeralPeriod', 'thaiNumPeriod'],
    ['ThaiNumeralParenthesisRight', 'thaiNumParenR'],
    ['ThaiNumeralParenthesesBoth', 'thaiNumParenBoth'],
    ['HindiAlphabetPeriod', 'hindiAlphaPeriod'],
    ['HindiNumeralPeriod', 'hindiNumPeriod'],
    ['HindiNumeralParenthesisRight', 'hindiNumParenR'],
    ['HindiAlphabet1Period', 'hindiAlpha1Period'],
  ],
  { fallback: 'Unsupported' }
);

/** The bullet kinds of a paragraph: `a:buNone`, `a:buAutoNum`, `a:buChar` or `a:buBlip`. */
export const BULLET_TYPE = enumMap<`${BulletType}`>(
  [
    ['None', 'none'],
    ['Numbered', 'autonum'],
    ['Unnumbered', 'char'],
  ],
  { fallback: 'Unsupported' }
);

/** The shape type of a protocol read (`geometricShape`) as the Office.js value. */
export function shapeTypeOf(type: string): `${ShapeType}` {
  const known: readonly `${ShapeType}`[] = [
    'Image',
    'GeometricShape',
    'Group',
    'Line',
    'Table',
    'Chart',
    'Diagram',
    'Freeform',
    'Media',
    'Ole',
    'Placeholder',
    'TextBox',
  ];
  const value = (type.charAt(0).toUpperCase() + type.slice(1)) as ShapeType;
  return known.includes(value) ? value : 'Unsupported';
}

/** The background fill kinds (EG_FillProperties of `p:bgPr`, ECMA-376 Part 1, 19.3.1.2). */
export const BACKGROUND_FILL_TYPE = enumMap<`${SlideBackgroundFillType}`>(
  [
    ['Solid', 'solid'],
    ['Gradient', 'gradient'],
    ['Pattern', 'pattern'],
    ['PictureOrTexture', 'picture'],
  ],
  { fallback: 'Unsupported' }
);

/**
 * The gradient shadings of a background: `a:lin`, or an `a:path` (ECMA-376 Part 1, 20.1.8.46).
 * `ShadeFromTitle` (`p:bgPr/@shadeToTitle`) is a read state only.
 */
export const BACKGROUND_GRADIENT_TYPE = enumMap<`${SlideBackgroundGradientFillType}`>(
  [
    ['Linear', 'linear'],
    ['Radial', 'circle'],
    ['Rectangular', 'rect'],
    ['Path', 'shape'],
  ],
  { fallback: 'Unsupported' }
);

/** `ST_PresetPatternVal` (ECMA-376 Part 1, 20.1.10.51). */
export const BACKGROUND_PATTERN = enumMap<`${SlideBackgroundPatternFillType}`>(
  [
    ['Cross', 'cross'],
    ['DarkDownwardDiagonal', 'dkDnDiag'],
    ['DarkHorizontal', 'dkHorz'],
    ['DarkUpwardDiagonal', 'dkUpDiag'],
    ['DarkVertical', 'dkVert'],
    ['DashedDownwardDiagonal', 'dashDnDiag'],
    ['DashedHorizontal', 'dashHorz'],
    ['DashedUpwardDiagonal', 'dashUpDiag'],
    ['DashedVertical', 'dashVert'],
    ['DiagonalBrick', 'diagBrick'],
    ['DiagonalCross', 'diagCross'],
    ['Divot', 'divot'],
    ['DottedDiamond', 'dotDmnd'],
    ['DottedGrid', 'dotGrid'],
    ['DownwardDiagonal', 'dnDiag'],
    ['Horizontal', 'horz'],
    ['HorizontalBrick', 'horzBrick'],
    ['LargeCheckerBoard', 'lgCheck'],
    ['LargeConfetti', 'lgConfetti'],
    ['LargeGrid', 'lgGrid'],
    ['LightDownwardDiagonal', 'ltDnDiag'],
    ['LightHorizontal', 'ltHorz'],
    ['LightUpwardDiagonal', 'ltUpDiag'],
    ['LightVertical', 'ltVert'],
    ['NarrowHorizontal', 'narHorz'],
    ['NarrowVertical', 'narVert'],
    ['OutlinedDiamond', 'openDmnd'],
    ['Percent10', 'pct10'],
    ['Percent20', 'pct20'],
    ['Percent25', 'pct25'],
    ['Percent30', 'pct30'],
    ['Percent40', 'pct40'],
    ['Percent5', 'pct5'],
    ['Percent50', 'pct50'],
    ['Percent60', 'pct60'],
    ['Percent70', 'pct70'],
    ['Percent75', 'pct75'],
    ['Percent80', 'pct80'],
    ['Percent90', 'pct90'],
    ['Plaid', 'plaid'],
    ['Shingle', 'shingle'],
    ['SmallCheckerBoard', 'smCheck'],
    ['SmallConfetti', 'smConfetti'],
    ['SmallGrid', 'smGrid'],
    ['SolidDiamond', 'solidDmnd'],
    ['Sphere', 'sphere'],
    ['Trellis', 'trellis'],
    ['UpwardDiagonal', 'upDiag'],
    ['Vertical', 'vert'],
    ['Wave', 'wave'],
    ['Weave', 'weave'],
    ['WideDownwardDiagonal', 'wdDnDiag'],
    ['WideUpwardDiagonal', 'wdUpDiag'],
    ['ZigZag', 'zigZag'],
  ],
  { fallback: 'Unsupported' }
);

/** The colors of a theme's `a:clrScheme` (CT_ColorScheme, ECMA-376 Part 1, 20.1.6.2). */
export const THEME_COLOR = enumMap<`${ThemeColor}`>([
  ['Dark1', 'dk1'],
  ['Light1', 'lt1'],
  ['Dark2', 'dk2'],
  ['Light2', 'lt2'],
  ['Accent1', 'accent1'],
  ['Accent2', 'accent2'],
  ['Accent3', 'accent3'],
  ['Accent4', 'accent4'],
  ['Accent5', 'accent5'],
  ['Accent6', 'accent6'],
  ['Hyperlink', 'hlink'],
  ['FollowedHyperlink', 'folHlink'],
]);

/**
 * The `styleName` of a table style (`a:tblStyle/@styleName`, ECMA-376 Part 1, 20.1.4.2.26) for
 * each `TableStyle` value: `MediumStyle2Accent1` is "Medium Style 2 - Accent 1". `NoStyleNoGrid`
 * is a table with no style choice, so it has no name.
 */
export function tableStyleName(value: unknown, target: string): string | null {
  // No guess here: a style that is not defined in the presentation refuses at the sync, so a
  // suggested style could fail too.
  if (typeof value !== 'string' || !TABLE_STYLES.has(value))
    failArgument(target, oneOf([...TABLE_STYLES]));
  if (value === 'NoStyleNoGrid') return null;
  if (value === 'NoStyleTableGrid') return 'No Style, Table Grid';
  const m = /^(Dark|Light|Medium|Themed)Style(\d)(?:Accent(\d))?$/.exec(value)!;
  return `${m[1]} Style ${m[2]}${m[3] ? ` - Accent ${m[3]}` : ''}`;
}

/** The `TableStyle` value of a table style name, or null for a name it has no value for. */
export function tableStyleOf(name: string): `${TableStyle}` | null {
  if (name === 'No Style, No Grid') return 'NoStyleNoGrid';
  if (name === 'No Style, Table Grid') return 'NoStyleTableGrid';
  const m = /^(Dark|Light|Medium|Themed) Style (\d)(?: - Accent (\d))?$/.exec(name);
  const value = m ? `${m[1]}Style${m[2]}${m[3] ? `Accent${m[3]}` : ''}` : '';
  return TABLE_STYLES.has(value) ? (value as `${TableStyle}`) : null;
}

const TABLE_STYLES: ReadonlySet<string> = new Set(Object.values(TableStyle));
