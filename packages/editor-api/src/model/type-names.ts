/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The Office.js type name of each object of the model, for `context.loadRecursive`, whose
// options are keyed by type name. By class identity, never by `constructor.name`, so a bundler
// that renames classes changes nothing.

import type { ClientObject } from '../runtime/client-object.ts';
import { Adjustments } from './adjustments.ts';
import { Application } from './application.ts';
import {
  SlideBackground,
  SlideBackgroundFill,
  SlideBackgroundGradientFill,
  SlideBackgroundPatternFill,
  SlideBackgroundPictureOrTextureFill,
  SlideBackgroundSolidFill,
  SlideLayoutBackground,
  SlideMasterBackground,
} from './backgrounds.ts';
import { Binding, BindingCollection } from './bindings.ts';
import { CustomProperty, CustomPropertyCollection } from './custom-properties.ts';
import {
  CustomXmlPart,
  CustomXmlPartCollection,
  CustomXmlPartScopedCollection,
} from './custom-xml.ts';
import { ShapeFill, ShapeLineFormat } from './fill-line.ts';
import { Hyperlink, HyperlinkCollection, HyperlinkScopedCollection } from './hyperlinks.ts';
import { DocumentProperties, PageSetup, Presentation } from './presentation.ts';
import {
  PlaceholderFormat,
  Shape,
  ShapeCollection,
  ShapeGroup,
  ShapeScopedCollection,
} from './shapes.ts';
import {
  Slide,
  SlideCollection,
  SlideLayout,
  SlideLayoutCollection,
  SlideMaster,
  SlideMasterCollection,
  SlideScopedCollection,
} from './slides.ts';
import { TableCell } from './table-cell.ts';
import { Border, Borders, Margins, TableStyleSettings } from './table-format.ts';
import {
  Table,
  TableCellCollection,
  TableColumn,
  TableColumnCollection,
  TableRow,
  TableRowCollection,
} from './tables.ts';
import { Tag, TagCollection } from './tags.ts';
import { ThemeColorScheme } from './theme.ts';
import { BulletFormat, ParagraphFormat, ShapeFont, TextFrame, TextRange } from './text.ts';

/** Each class with its Office.js type name. */
const CLASSES: readonly (readonly [unknown, string])[] = [
  [Adjustments, 'Adjustments'],
  [Application, 'Application'],
  [Binding, 'Binding'],
  [BindingCollection, 'BindingCollection'],
  [Border, 'Border'],
  [Borders, 'Borders'],
  [BulletFormat, 'BulletFormat'],
  [CustomProperty, 'CustomProperty'],
  [CustomPropertyCollection, 'CustomPropertyCollection'],
  [CustomXmlPart, 'CustomXmlPart'],
  [CustomXmlPartCollection, 'CustomXmlPartCollection'],
  [CustomXmlPartScopedCollection, 'CustomXmlPartScopedCollection'],
  [DocumentProperties, 'DocumentProperties'],
  [Hyperlink, 'Hyperlink'],
  [HyperlinkCollection, 'HyperlinkCollection'],
  [HyperlinkScopedCollection, 'HyperlinkScopedCollection'],
  [Margins, 'Margins'],
  [PageSetup, 'PageSetup'],
  [ParagraphFormat, 'ParagraphFormat'],
  [PlaceholderFormat, 'PlaceholderFormat'],
  [Presentation, 'Presentation'],
  [Shape, 'Shape'],
  [ShapeCollection, 'ShapeCollection'],
  [ShapeFill, 'ShapeFill'],
  [ShapeFont, 'ShapeFont'],
  [ShapeGroup, 'ShapeGroup'],
  [ShapeLineFormat, 'ShapeLineFormat'],
  [ShapeScopedCollection, 'ShapeScopedCollection'],
  [Slide, 'Slide'],
  [SlideBackground, 'SlideBackground'],
  [SlideBackgroundFill, 'SlideBackgroundFill'],
  [SlideBackgroundGradientFill, 'SlideBackgroundGradientFill'],
  [SlideBackgroundPatternFill, 'SlideBackgroundPatternFill'],
  [SlideBackgroundPictureOrTextureFill, 'SlideBackgroundPictureOrTextureFill'],
  [SlideBackgroundSolidFill, 'SlideBackgroundSolidFill'],
  [SlideCollection, 'SlideCollection'],
  [SlideLayout, 'SlideLayout'],
  [SlideLayoutBackground, 'SlideLayoutBackground'],
  [SlideLayoutCollection, 'SlideLayoutCollection'],
  [SlideMaster, 'SlideMaster'],
  [SlideMasterBackground, 'SlideMasterBackground'],
  [SlideMasterCollection, 'SlideMasterCollection'],
  [SlideScopedCollection, 'SlideScopedCollection'],
  [Table, 'Table'],
  [TableCell, 'TableCell'],
  [TableCellCollection, 'TableCellCollection'],
  [TableColumn, 'TableColumn'],
  [TableColumnCollection, 'TableColumnCollection'],
  [TableRow, 'TableRow'],
  [TableRowCollection, 'TableRowCollection'],
  [TableStyleSettings, 'TableStyleSettings'],
  [Tag, 'Tag'],
  [TagCollection, 'TagCollection'],
  [TextFrame, 'TextFrame'],
  [TextRange, 'TextRange'],
  [ThemeColorScheme, 'ThemeColorScheme'],
];

const BY_CLASS = new Map<unknown, string>(CLASSES.map(([type, name]) => [type, name]));

/** Every type name, sorted. */
export const TYPE_NAMES: readonly string[] = Object.freeze(CLASSES.map(([, name]) => name).sort());

/** The Office.js type name of an object of the model, or undefined. */
export function typeNameOf(object: ClientObject): string | undefined {
  return BY_CLASS.get(Object.getPrototypeOf(object)?.constructor);
}
