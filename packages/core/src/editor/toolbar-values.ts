// The typed values of value slots, after DOCX Editor's
// `packages/core/src/editor/toolbar-values.ts`. `runToolbarCommand` and the adapters'
// `useEditorValueCommand` take these types, so a typo in a value does not compile. A slot that is
// not listed takes `unknown`, as before. DOCX Editor commit adf0e178.

import type {
  ColorChoice,
  FontFormatCommand,
  ParagraphFormatCommand,
  LineDash,
  LineEndType,
  LineKind,
  ObjectAlign,
  OrderMove,
  RotateChoice,
  ShapeEffectsValue,
  TextEffectsValue,
  ShapeStyleChoice,
  ShapeStyleValue,
  TextAlign,
  TextCase,
  TextFitting,
  TextOrientation,
  VerticalAlign,
} from '../contracts/index.ts';
import type { ChromeSlotId } from './chrome-controls.ts';

/** The keys of the Bullets gallery (Home > Bullets), in its order. @public */
export type BulletGalleryKey =
  | 'none'
  | 'filledRound'
  | 'hollowRound'
  | 'filledSquare'
  | 'hollowSquare'
  | 'star'
  | 'arrow'
  | 'checkmark';

/** The keys of the Numbering gallery (Home > Numbering), in its order. @public */
export type NumberGalleryKey =
  | 'none'
  | 'arabicPeriod'
  | 'arabicParenR'
  | 'romanUcPeriod'
  | 'alphaUcPeriod'
  | 'alphaLcParenR'
  | 'alphaLcPeriod'
  | 'romanLcPeriod';

/** The value each typed value slot takes. @public */
export interface ToolbarValueMap {
  'font.family': string;
  /**
   * Points (`1` or more), or `{ step: 1 }` / `{ step: -1 }` for Increase and Decrease Font Size
   * (the size list).
   */
  'font.size': number | { readonly step: 1 | -1 };
  /** Points added after each character; 0 is Normal, negative condenses. */
  'text.characterSpacing': number;
  /** Home > Change Case. */
  'text.changeCase': TextCase;
  'text.color': ColorChoice | null;
  'text.highlight': ColorChoice | null;
  /** A URL, or null to remove the link. */
  'text.link': string | null;
  'paragraph.align': TextAlign;
  /** A line spacing multiple (`1.5`). */
  'list.lineSpacing': number;
  'paragraph.valign': VerticalAlign;
  /** Home > Text Direction: the orientations the editor sets. */
  'paragraph.textDirection': Exclude<TextOrientation, 'mongolianVertical' | 'wordArtVerticalRtl'>;
  /** How the text fits its shape: `none`, `shrink` (text on overflow), or `shape` (shape to text). */
  'paragraph.autofit': TextFitting;
  /** The number of text columns, 1 to 16. */
  'paragraph.columns': number;
  /** A Bullets gallery item; `none` turns the list off. */
  'list.bulletStyle': BulletGalleryKey;
  /** A Numbering gallery item; `none` turns the list off. */
  'list.numberStyle': NumberGalleryKey;
  /** The Paragraph dialog's fields, run as one `setParagraphFormat` without the dialog. */
  'paragraph.dialog': ParagraphFormatCommand;
  /** The Font dialog's fields, run as one `setFontFormat` without the dialog. */
  'text.dialog': FontFormatCommand;
  /** A preset geometry name (`rect`). */
  'insert.shape': string;
  'insert.line': LineKind;
  'shape.fill': ColorChoice | null;
  'shape.lineColor': ColorChoice | null;
  /** Points. */
  'shape.lineWidth': number;
  'shape.lineStart': LineEndType;
  'shape.lineEnd': LineEndType;
  'shape.lineDash': LineDash;
  'arrange.order': OrderMove;
  'arrange.align': ObjectAlign | 'distributeHorizontal' | 'distributeVertical';
  'shape.styles': ShapeStyleValue | ShapeStyleChoice;
  'shape.effects': ShapeEffectsValue;
  'shape.textEffects': TextEffectsValue;
  'arrange.rotate': RotateChoice;
  /** A layout id from `snapshot().layouts`. */
  'slide.add': string;
  /** A layout id from `snapshot().layouts`. */
  'slide.layout': string;
  'view.presentationView': 'normal' | 'slideSorter';
}

/** The value a slot takes: its `ToolbarValueMap` type, or `unknown` for a slot not listed. @public */
export type ToolbarSlotValue<K extends ChromeSlotId> = K extends keyof ToolbarValueMap
  ? ToolbarValueMap[K]
  : unknown;

/**
 * The value a typed slot shows: its `ToolbarValueMap` type, except `shape.effects`, which shows
 * one choice for each effect kind, separated by spaces (`shadow.outerTop reflection.none …`),
 * `shape.styles`, which shows the `row.color` string, `font.size`, which shows points, and
 * `paragraph.textDirection`, which also shows the two orientations the editor does not set.
 * @public
 */
export type ToolbarShownValue<K extends keyof ToolbarValueMap> = K extends 'shape.effects'
  ? string
  : K extends 'shape.styles'
    ? ShapeStyleValue
    : K extends 'font.size'
      ? number
      : K extends 'paragraph.textDirection'
        ? TextOrientation
        : ToolbarValueMap[K];
