/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The hint of each argument that the engine refuses with `invalid-argument`, by the name the
// engine gives (the detail of the refusal). Every refusal of an argument says what to pass, so
// no argument refusal falls back to a generic text. `scripts/__tests__/refusal-reasons.test.ts`
// fails when the engine names an argument that has no row here.

const WHOLE_FROM_ZERO = 'Use a whole number of 0 or more that is less than the number of items.';
const BOOLEAN = 'Use true or false.';
const TEXT = 'Use text of at most 4096 characters, with no control characters.';
const POINTS = 'Use a number of points from 0 to 4032.';
const MARGIN = 'Use a margin of 0 to 4032 points.';

/** `[engine detail, hint]`. A hint says what to pass. */
const ROWS: readonly (readonly [string, string])[] = [
  // Colors and fills.
  ['color', "Use '#RRGGBB' (for example '#1F4E79') or a named color such as 'red'."],
  ['foregroundColor', "Use '#RRGGBB' (for example '#1F4E79') or a named color such as 'red'."],
  ['backgroundColor', "Use '#RRGGBB' (for example '#1F4E79') or a named color such as 'red'."],
  ['transparency', 'Use a fraction from 0 (opaque) to 1 (clear).'],
  ['fill', 'Use a fill object: a solid color, a gradient, a pattern, a picture, or no fill.'],
  ['gradientType', 'Use a value of SlideBackgroundGradientFillType.'],
  ['pattern', 'Use a value of SlideBackgroundPatternFillType.'],
  ['imageBase64', 'Use the Base64 bytes of a PNG, JPEG, GIF, or BMP picture.'],
  ['base64', 'Use the Base64 bytes of a PNG, JPEG, GIF, or BMP picture.'],
  ['image-size', 'Use a width and a height greater than 0, at most 8192 pixels.'],
  ['background', 'Use a fill write that this slide can take; a layout or master cannot inherit.'],
  ['areBackgroundGraphicsHidden', BOOLEAN],
  // Lines and borders.
  ['line', 'Use a line object with color, weight, dashStyle, style, transparency, or visible.'],
  ['weight', 'Use a line weight from 0 to 1584 points.'],
  ['dashStyle', 'Use a value of ShapeLineDashStyle.'],
  ['style', 'Use a value of ShapeLineStyle, or a table style the presentation defines.'],
  ['visible', BOOLEAN],
  ['side', 'Use a border side: left, right, top, bottom, diagonalDown, or diagonalUp.'],
  // Shapes and their boxes.
  ['shape', 'Use a shape kind that this collection can add, with its options.'],
  ['shapes', 'Use an array of shapes of one slide or one group, at most 10 000.'],
  ['geometry', 'Use a value of GeometricShapeType.'],
  ['connector', 'Use a value of ConnectorType.'],
  ['left', 'Use a number of points from -169 093 to 169 093.'],
  ['top', 'Use a number of points from -169 093 to 169 093.'],
  ['width', 'Use a number of points from 0 to 169 093.'],
  ['height', 'Use a number of points from 0 to 169 093.'],
  ['rotation', 'Use a rotation in degrees, a finite number.'],
  ['position', 'Use a value of ShapeZOrder.'],
  ['props', 'Use an object of the properties to write.'],
  ['isDecorative', BOOLEAN],
  ['title', TEXT],
  ['descr', TEXT],
  ['altTextTitle', TEXT],
  ['altTextDescription', TEXT],
  ['name', 'Use a name of 1 to 255 characters, with no control characters.'],
  ['value', 'Use a value of the type this property takes, in its range.'],
  ['index', WHOLE_FROM_ZERO],
  ['group', 'Use a shape that is a group.'],
  ['group-count', 'Use two or more different shapes.'],
  ['group-parent', 'Use shapes of one slide, or of one group.'],
  // Text.
  ['text', 'Use text of at most 1 000 000 characters, with no control characters.'],
  ['span', 'Use a start and a length inside the text, around some text.'],
  ['start', 'Use a start of 0 or more, at most the length of the text.'],
  ['length', 'Use a length of 0 or more that ends inside the text.'],
  ['font', 'Use an object of font properties.'],
  ['size', 'Use a font size from 1 to 4000 points.'],
  ['underline', 'Use a value of ShapeFontUnderlineStyle.'],
  ['bold', BOOLEAN],
  ['italic', BOOLEAN],
  ['allCaps', BOOLEAN],
  ['smallCaps', BOOLEAN],
  ['strikethrough', BOOLEAN],
  ['doubleStrikethrough', BOOLEAN],
  ['subscript', BOOLEAN],
  ['superscript', BOOLEAN],
  ['paragraph', 'Use an object of paragraph format properties.'],
  ['alignment', 'Use a value of ParagraphHorizontalAlignment.'],
  ['indentLevel', 'Use a whole number from 1 to 9.'],
  ['bulletStyle', 'Use a value of ShapeBulletStyle.'],
  ['verticalAlignment', 'Use a value of TextVerticalAlignment.'],
  ['wordWrap', BOOLEAN],
  ['autoSize', 'Use a value of ShapeAutoSize.'],
  ['leftMargin', MARGIN],
  ['rightMargin', MARGIN],
  ['topMargin', MARGIN],
  ['bottomMargin', MARGIN],
  [
    'address',
    "Use an 'http', 'https', 'mailto', 'tel', or 'ftp' address of at most 2048 characters.",
  ],
  ['screenTip', 'Use a screen tip of at most 1024 characters, with no control characters.'],
  // Tables.
  ['values', 'Use an array of rows of strings, at most one per row and one per column.'],
  ['cell', 'Use a row and a column inside the table.'],
  ['row', 'Use a row index of 0 or more, less than the number of rows.'],
  ['column', 'Use a column index of 0 or more, less than the number of columns.'],
  ['rows', 'Use a whole number of rows from 1 to 75.'],
  ['columns', 'Use a whole number of columns from 1 to 75.'],
  ['rowCount', 'Use a whole number of rows of 1 or more that ends inside the table.'],
  ['columnCount', 'Use a whole number of columns of 1 or more that ends inside the table.'],
  ['count', 'Use a whole number of 1 or more that fits the table.'],
  ['lineSize', POINTS],
  ['anchor', 'Use a value of TableCellVerticalAlignment.'],
  ['anchorCenter', BOOLEAN],
  ['margins', 'Use an object of margins, each from 0 to 4032 points, or null.'],
  ['options', 'Use true or false for text and for format.'],
  ['styleName', 'Use the name of a table style as a string, or null.'],
  ['firstRow', BOOLEAN],
  ['lastRow', BOOLEAN],
  ['firstColumn', BOOLEAN],
  ['lastColumn', BOOLEAN],
  ['bandedRows', BOOLEAN],
  ['bandedColumns', BOOLEAN],
  // Presentation, properties, tags, custom XML, slides.
  ['properties', 'Use an object of document properties.'],
  ['revisionNumber', 'Use a whole number from 0 to 2147483647.'],
  ['subject', TEXT],
  ['author', TEXT],
  ['keywords', TEXT],
  ['comments', TEXT],
  ['category', TEXT],
  ['lastAuthor', TEXT],
  ['company', TEXT],
  ['manager', TEXT],
  ['slideWidth', 'Use a slide width from 72 to 4032 points (1 to 56 inches).'],
  ['slideHeight', 'Use a slide height from 72 to 4032 points (1 to 56 inches).'],
  [
    'key',
    'Use a key of 1 to 255 characters, with no control characters, on a presentation with fewer than 2 147 483 647 properties.',
  ],
  ['type', 'Use a value of DocumentPropertyType.'],
  ['id', 'Use an id of 1 to 255 characters.'],
  ['binding-limit', 'Delete bindings first: a runtime keeps at most 10 000.'],
  ['binding-selection', 'Select exactly one shape first.'],
  ['custom-xml', 'Use well-formed XML text with no DTD and one root element.'],
  ['custom-xml-id', 'Use an id that no other custom XML part has.'],
  ['xml', 'Delete custom XML parts first: the presentation has no free part name left.'],
  ['namespaceUri', 'Use a namespace URI as a string.'],
  ['slides', 'Use a non-empty array of slides of this presentation.'],
  ['layoutId', 'Use the id of a slide layout of this presentation as a string.'],
  ['slideMasterId', 'Use the id of a slide master of this presentation as a string.'],
  ['formatting', 'Use a value of InsertSlideFormatting.'],
  ['import-base64', 'Use the Base64 bytes of a PPTX file.'],
  ['sourceSlideIds', 'Use an array of different slide ids of the source file.'],
  ['operations', 'Make fewer calls in one sync.'],
];

/** The hints, in a record with no prototype, so a detail never reaches a built-in name. */
export const ARGUMENT_HINTS: Readonly<Record<string, string>> = Object.freeze(
  Object.assign(Object.create(null) as Record<string, string>, Object.fromEntries(ROWS))
);

/** The hint for an argument the engine refused, or undefined when the engine named none. */
export function argumentHint(detail: string | undefined): string | undefined {
  return detail === undefined ? undefined : ARGUMENT_HINTS[detail];
}
