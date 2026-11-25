// The operation vocabulary, after DOCX Editor's `packages/core/src/automation/operations.ts`. DOCX
// Editor commit adf0e178.
//
// Each operation is plain data. A query answers from the working package; a command lowers to
// store ops through the edits in `commands/`. Lengths are in points and angles in degrees, as
// the object model shows them; the host converts to EMU (12 700 per point) at this boundary.

import type {
  FindOptions,
  LineEndType,
  ShapeFormatPatch,
  TransitionKind,
} from '../contracts/neutral.ts';
import type { TransitionDirection } from './commands/transition-commands.ts';
import type { AutomationRef } from './protocol.ts';

/** How a collection item is found: by position, or by the id the object model shows. */
export type AutomationItemKey = { readonly index: number } | { readonly id: string };

/** A box in points. Each field is optional on writes. */
export interface AutomationBox {
  readonly left?: number;
  readonly top?: number;
  readonly width?: number;
  readonly height?: number;
}

/** A stretch of a shape's text: a UTF-16 start and length. Absent means the whole text. */
export interface AutomationTextSpan {
  readonly start: number;
  readonly length: number;
}

/**
 * A table cell whose text body a text operation addresses instead of the shape's own: the
 * cell at this grid position, or the anchor of the merged block that covers it.
 */
export interface AutomationCellAddress {
  readonly row: number;
  readonly column: number;
}

/** What a new shape is. */
export type AutomationNewShape =
  /** A preset geometry (`ST_ShapeType`, ECMA-376 Part 1, 20.1.10.56). */
  | { readonly kind: 'geometric'; readonly geometry: string }
  /** A connector: `straight`, `elbow`, or `curve`. */
  | { readonly kind: 'line'; readonly connector: 'straight' | 'elbow' | 'curve' }
  | { readonly kind: 'textBox'; readonly text: string }
  | {
      readonly kind: 'table';
      readonly rows: number;
      readonly columns: number;
      /** Row-major cell text. */
      readonly values?: readonly (readonly string[])[];
    }
  /**
   * A picture (`p:pic`, `CT_Picture`, ECMA-376 Part 1, §19.3.1.37) over a new Image part
   * (§15.2.14): the Base64 bytes of a PNG, JPEG, or GIF image of at most 50 MiB.
   */
  | { readonly kind: 'picture'; readonly base64: string };

/** Shape properties a write sets. */
export interface AutomationShapeWrite extends AutomationBox {
  readonly name?: string;
  readonly rotation?: number;
  readonly visible?: boolean;
  readonly altTextTitle?: string;
  readonly altTextDescription?: string;
  readonly isDecorative?: boolean;
}

/** A fill write. Colors are `#RRGGBB` or a preset color name; transparency is 0 to 1. */
export type AutomationFillWrite =
  | { readonly kind: 'solid'; readonly color: string; readonly transparency?: number }
  | { readonly kind: 'clear' }
  /** Change the transparency of the current solid fill. */
  | { readonly kind: 'transparency'; readonly transparency: number }
  /** A picture fill (`a:blipFill`) from PNG, JPEG, or GIF bytes in base64, stretched. */
  | { readonly kind: 'image'; readonly base64: string };

/**
 * A background write (`p:cSld/p:bg`, ECMA-376 Part 1, 19.3.1.1). A fill write gives the part an
 * own `p:bgPr`; a field it leaves out keeps the current value, or takes the default.
 */
export type AutomationBackgroundWrite =
  /** Remove the own `p:bg`: show the background of the layout or master. */
  | { readonly kind: 'follow' }
  /** Copy the background the part shows into an own `p:bg`. */
  | { readonly kind: 'own' }
  /** `showMasterSp`: hide or show the shapes of the level above. */
  | { readonly kind: 'graphicsHidden'; readonly hidden: boolean }
  | { readonly kind: 'solid'; readonly color?: string; readonly transparency?: number }
  /** The gradient shading: `linear`, or an `a:path` of `circle`, `rect`, or `shape`. */
  | { readonly kind: 'gradient'; readonly gradientType?: 'linear' | 'circle' | 'rect' | 'shape' }
  | {
      readonly kind: 'pattern';
      /** `ST_PresetPatternVal`. */
      readonly pattern?: string;
      readonly foregroundColor?: string;
      readonly backgroundColor?: string;
    }
  /** A stretched picture from PNG, JPEG, or GIF bytes in base64, or the current picture. */
  | { readonly kind: 'picture'; readonly base64?: string; readonly transparency?: number };

/** Outline properties a write sets. */
export interface AutomationLineWrite {
  /** Start/end arrowhead type. Other endpoint properties remain unchanged. */
  readonly head?: LineEndType;
  readonly tail?: LineEndType;
  readonly visible?: boolean;
  readonly color?: string;
  /** Points. */
  readonly weight?: number;
  /** `ST_PresetLineDashVal`. */
  readonly dashStyle?: string;
  /** `ST_CompoundLine`. */
  readonly style?: string;
  readonly transparency?: number;
}

/** Text frame properties a write sets. Margins in points. */
export interface AutomationTextFrameWrite {
  readonly autoSize?: 'none' | 'normal' | 'shape';
  readonly leftMargin?: number;
  readonly rightMargin?: number;
  readonly topMargin?: number;
  readonly bottomMargin?: number;
  /** `t`, `ctr`, `b`, or with a centered column `tCentered`, `ctrCentered`, `bCentered`. */
  readonly verticalAlignment?: string;
  readonly wordWrap?: boolean;
}

/** Font properties a write sets. */
export interface AutomationFontWrite {
  readonly name?: string;
  /** Points. */
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** `ST_TextUnderlineType`. */
  readonly underline?: string;
  readonly color?: string;
  readonly strikethrough?: boolean;
  readonly doubleStrikethrough?: boolean;
  readonly allCaps?: boolean;
  readonly smallCaps?: boolean;
  readonly superscript?: boolean;
  readonly subscript?: boolean;
  /**
   * The run highlight (`a:highlight`, CT_Color, ECMA-376 Part 1, 21.1.2.3.4): `#RRGGBB` or a
   * preset color name sets it, null removes it.
   */
  readonly highlightColor?: string | null;
}

/** Paragraph properties a write sets. */
export interface AutomationParagraphWrite {
  /** `ST_TextAlignType`. */
  readonly alignment?: string;
  readonly bulletVisible?: boolean;
  /** `autonum` with a scheme, or `char`. */
  readonly bulletType?: 'none' | 'char' | 'autonum';
  readonly bulletStyle?: string;
  /** 1-based list level, 1 to 9. */
  readonly indentLevel?: number;
  /**
   * Line spacing in points (`a:lnSpc`, CT_TextSpacing, ECMA-376 Part 1, 21.1.2.2.5), 0.01 to
   * 1584. Each paragraph keeps its rule. An exact spacing (`a:spcPts`) takes the points. A
   * multiple (`a:spcPct`) takes the points over 12, so 12 points is one line.
   */
  readonly lineSpacing?: number;
  /**
   * The line spacing kind: `multiple` (`a:spcPct`) or `exactly` (`a:spcPts`). Each paragraph
   * keeps the points of its spacing, so `exactly` on a 1.5-line paragraph gives 18 points.
   * Applied before `lineSpacing` when both are given.
   */
  readonly lineSpacingRule?: 'multiple' | 'exactly';
  /** Space before in points (`a:spcBef/a:spcPts`, ECMA-376 Part 1, 21.1.2.2.10), 0 to 1584. */
  readonly spaceBefore?: number;
  /** Space after in points (`a:spcAft/a:spcPts`, ECMA-376 Part 1, 21.1.2.2.9), 0 to 1584. */
  readonly spaceAfter?: number;
  /** Space before in lines (`a:spcBef/a:spcPct`), 0 to 132. Not with `spaceBefore`. */
  readonly lineUnitBefore?: number;
  /** Space after in lines (`a:spcAft/a:spcPct`), 0 to 132. Not with `spaceAfter`. */
  readonly lineUnitAfter?: number;
  /** The left margin in points (`a:pPr/@marL`, ST_TextMargin), 0 to 4032. */
  readonly leftIndent?: number;
  /** The first line indent in points (`a:pPr/@indent`, ST_TextIndent), -4032 to 4032. */
  readonly firstLineIndent?: number;
}

/** Document properties a write sets. */
export interface AutomationPropertiesWrite {
  readonly title?: string;
  readonly subject?: string;
  readonly author?: string;
  readonly keywords?: string;
  readonly comments?: string;
  readonly category?: string;
  readonly revisionNumber?: number;
  readonly company?: string;
  readonly manager?: string;
}

/** A custom property value a write sets. A Date is an ISO 8601 text. */
export type AutomationCustomPropertyWrite =
  | { readonly type: 'String'; readonly value: string }
  | { readonly type: 'Number'; readonly value: number }
  | { readonly type: 'Boolean'; readonly value: boolean }
  | { readonly type: 'Date'; readonly value: string };

/** Options of a slide import (Office.js `InsertSlideOptions`). */
export interface AutomationInsertSlidesOptions {
  /** Keep the source layouts, masters, and themes (default), or use this presentation's. */
  readonly formatting?: 'keepSource' | 'useDestinationTheme';
  /** The source slides, by slide id, in the order they go in. Absent: every slide. */
  readonly sourceSlideIds?: readonly string[];
  /** The slide the new slides follow, by id. Absent: they go first. */
  readonly targetSlideId?: string;
}

/** A line of a table cell: an edge, or a diagonal (`a:lnTlToBr` down, `a:lnBlToTr` up). */
export type AutomationCellBorderSide =
  'left' | 'right' | 'top' | 'bottom' | 'diagonalDown' | 'diagonalUp';

/** Cell margins a write sets, in points. Null removes one, back to its default. */
export interface AutomationCellMarginsWrite {
  readonly left?: number | null;
  readonly right?: number | null;
  readonly top?: number | null;
  readonly bottom?: number | null;
}

/** Table style options (`a:tblPr` flags, ECMA-376 Part 1, 21.1.3.15) and the style a write sets. */
export interface AutomationTableStyleWrite {
  readonly firstRow?: boolean;
  readonly lastRow?: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
  readonly bandedRows?: boolean;
  readonly bandedColumns?: boolean;
  /**
   * The style by its `styleName`: one the presentation defines in `ppt/tableStyles.xml`, or the
   * built-in default. Null removes the style choice.
   */
  readonly styleName?: string | null;
}

/**
 * What a caller read of a comment, for a write to it. The host refuses the write when the
 * comment at the id has another author or date, so a write never reaches another comment.
 */
export interface AutomationCommentExpect {
  readonly author: string;
  readonly date?: string;
  readonly text: string;
}

/** `ShapeZOrder` moves. */
export type AutomationZOrder = 'bringForward' | 'bringToFront' | 'sendBackward' | 'sendToBack';

/**
 * The slide transition settings a write sets (`p:transition`, `CT_SlideTransition`, ECMA-376
 * Part 1, §19.3.1.50). Absent fields keep their values.
 */
export interface AutomationTransitionWrite {
  /**
   * The effect element to write. On a transition with an effect, `none` removes the effect, the
   * sound, and the extensions, and keeps the advance settings. A slide with only a sound keeps it.
   */
  readonly kind?: TransitionKind | 'none';
  /** 10 to 59,000 ms: `p14:dur`, and the nearest `spd` (`ST_TransitionSpeed`, §19.7.54). */
  readonly durationMs?: number;
  /** `advClick`. */
  readonly advanceOnClick?: boolean;
  /** `advTm`: 0 to 3,600,000 ms, or null to remove it. */
  readonly advanceAfterMs?: number | null;
  /** The side of `push` and `wipe`, or the orientation and direction of `split`. */
  readonly direction?: TransitionDirection;
  /** `thruBlk` of `fade` and `cut` (`CT_OptionalBlackTransition`). */
  readonly throughBlack?: boolean;
  /** `true` alone removes the whole `p:transition`, with its sound and extensions. */
  readonly remove?: true;
}

/**
 * The slide number, footer, and date that a `setHeaderFooter` write shows or hides, through the
 * `sldNum`, `ftr`, and `dt` placeholders (`p:ph`, `ST_PlaceholderType`, ECMA-376 Part 1,
 * 19.3.1.36 and 19.7.10). A field the write leaves out stays as it is.
 */
export interface AutomationHeaderFooterWrite {
  /** Show or hide the slide number: a `sldNum` placeholder with a `slidenum` field. */
  readonly slideNumberVisible?: boolean;
  /** False hides the footer. True without `footerText` keeps or spreads the deck's footer. */
  readonly footerVisible?: boolean;
  /** One line of 1 to 1024 characters. Shows the footer. Only '' with `footerVisible: false`. */
  readonly footerText?: string;
  /** False hides the date. True without other date fields keeps or spreads the deck's date. */
  readonly dateVisible?: boolean;
  /**
   * One line of 1 to 1024 characters. Shows the date: fixed text, or a `datetime1` field that
   * holds it with `dateFixed: false`. Only '' with `dateVisible: false`.
   */
  readonly dateText?: string;
  /**
   * True turns date fields into fixed text; false turns fixed text into today's date field.
   * Ignored with `dateVisible: false`.
   */
  readonly dateFixed?: boolean;
  /** Slides on a `title` layout (`ST_SlideLayoutType`, 19.7.15) hide what the others show. */
  readonly hideOnTitleSlides?: boolean;
}

/** Every operation a host understands. */
export type AutomationOperation =
  // Presentation
  | { readonly op: 'getPresentation' }
  | { readonly op: 'readPresentation' }
  | { readonly op: 'setPageSetup'; readonly slideWidth?: number; readonly slideHeight?: number }
  | { readonly op: 'getProperties' }
  | { readonly op: 'setProperties'; readonly properties: AutomationPropertiesWrite }
  // Custom properties: `owner` is the presentation.
  | { readonly op: 'readCustomProperties'; readonly owner: AutomationRef }
  | {
      readonly op: 'setCustomProperty';
      readonly owner: AutomationRef;
      readonly key: string;
      readonly value: AutomationCustomPropertyWrite;
    }
  | { readonly op: 'deleteCustomProperty'; readonly owner: AutomationRef; readonly key: string }
  | { readonly op: 'deleteCustomProperties'; readonly owner: AutomationRef }
  // Custom XML parts of a presentation, slide, layout, master, or shape. `id` is the item id.
  | {
      readonly op: 'listCustomXmlParts';
      readonly owner: AutomationRef;
      /** Only the parts whose root has this namespace URI. */
      readonly namespaceUri?: string;
    }
  | { readonly op: 'readCustomXml'; readonly owner: AutomationRef; readonly id: string }
  | {
      readonly op: 'addCustomXmlPart';
      readonly owner: AutomationRef;
      readonly xml: string;
      /** The item id of the new part, a GUID. The host draws one when it is absent. */
      readonly id?: string;
    }
  | {
      readonly op: 'setCustomXml';
      readonly owner: AutomationRef;
      readonly id: string;
      readonly xml: string;
    }
  | { readonly op: 'deleteCustomXmlPart'; readonly owner: AutomationRef; readonly id: string }
  // Slides
  | { readonly op: 'listSlides' }
  | { readonly op: 'getSlide'; readonly key: AutomationItemKey; readonly orNull?: boolean }
  | { readonly op: 'readSlide'; readonly slide: AutomationRef }
  | {
      readonly op: 'addSlide';
      /** Zero-based position; the end when absent. */
      readonly index?: number;
      readonly layout?: AutomationRef;
      readonly master?: AutomationRef;
      /** A layout by `p:sldLayoutId/@id`, as `AddSlideOptions.layoutId` names it. */
      readonly layoutId?: string;
      /** A master by `p:sldMasterId/@id`, as `AddSlideOptions.slideMasterId` names it. */
      readonly masterId?: string;
    }
  | { readonly op: 'getSlideLayout'; readonly slide: AutomationRef }
  | { readonly op: 'getSlideMaster'; readonly owner: AutomationRef }
  | { readonly op: 'deleteSlide'; readonly slide: AutomationRef }
  | {
      /**
       * Copy a slide, with its notes and the parts it owns, directly after it (`p:sldIdLst`,
       * ECMA-376 Part 1, 19.2.1.34). Answers the handle of the copy.
       */
      readonly op: 'duplicateSlide';
      readonly slide: AutomationRef;
    }
  | { readonly op: 'moveSlide'; readonly slide: AutomationRef; readonly index: number }
  | {
      /** Hide or show a slide: `p:sld/@show` (ECMA-376 Part 1, §19.3.1.38). */
      readonly op: 'setSlideHidden';
      readonly slide: AutomationRef;
      readonly hidden: boolean;
    }
  | { readonly op: 'applyLayout'; readonly slide: AutomationRef; readonly layout: AutomationRef }
  | { readonly op: 'exportSlides'; readonly slides: readonly AutomationRef[] }
  | {
      /**
       * Check a picture of a slide, or of a shape on a slide, and answer its handle for
       * {@link AutomationHost.renderImage}. Needs the `images` capability.
       */
      readonly op: 'prepareImage';
      readonly target: AutomationRef;
      readonly width?: number;
      readonly height?: number;
    }
  | {
      /** Import the slides of another `.pptx`, given as base64. */
      readonly op: 'insertSlides';
      readonly base64: string;
      readonly options?: AutomationInsertSlidesOptions;
    }
  | { readonly op: 'listHyperlinks'; readonly slide: AutomationRef }
  | {
      /**
       * Replace every match of `query` in the slide shapes, table cells, and speaker notes.
       * Answers the number of replaced matches. A match stays in one paragraph and never covers
       * a field.
       */
      readonly op: 'replaceText';
      readonly query: string;
      readonly replacement: string;
      readonly options?: FindOptions;
      /** Only this slide: its shapes, its table cells, and its speaker notes. */
      readonly slide?: AutomationRef;
    }
  // Bindings: the host's named links to shapes, kept in memory, never in the package. The owner
  // is the presentation.
  | { readonly op: 'listBindings'; readonly owner: AutomationRef }
  | {
      /** Bind a shape under `id`; a binding with that id moves to the shape. */
      readonly op: 'addBinding';
      readonly owner: AutomationRef;
      readonly shape: AutomationRef;
      readonly id: string;
    }
  | {
      /** Bind the one selected shape under `id`. Needs the `selection` capability. */
      readonly op: 'addBindingFromSelection';
      readonly owner: AutomationRef;
      readonly id: string;
    }
  | { readonly op: 'deleteBinding'; readonly owner: AutomationRef; readonly id: string }
  | {
      /** The handle of the shape a binding names; `item-not-found` when the shape is gone. */
      readonly op: 'getBindingShape';
      readonly owner: AutomationRef;
      readonly id: string;
    }
  // Masters and layouts
  | { readonly op: 'listMasters' }
  | { readonly op: 'getMaster'; readonly key: AutomationItemKey; readonly orNull?: boolean }
  | { readonly op: 'readMaster'; readonly master: AutomationRef }
  | { readonly op: 'listLayouts'; readonly master: AutomationRef }
  | {
      readonly op: 'getLayout';
      readonly master: AutomationRef;
      readonly key: AutomationItemKey;
      readonly orNull?: boolean;
    }
  | { readonly op: 'readLayout'; readonly layout: AutomationRef }
  // Shapes. A container is a slide, a layout, a master, or a group shape.
  | {
      readonly op: 'listShapes';
      readonly container: AutomationRef;
      /** A container shape that is not a group answers null instead of refusing. */
      readonly orNull?: boolean;
    }
  | {
      readonly op: 'getShape';
      readonly container: AutomationRef;
      readonly key: AutomationItemKey;
      readonly orNull?: boolean;
    }
  | {
      readonly op: 'readShape';
      readonly shape: AutomationRef;
      /** Read the shape as a group: a shape that is not a group refuses (`not-group`). */
      readonly group?: boolean;
      /** With `group`, a shape that is not a group answers null instead of refusing. */
      readonly orNull?: boolean;
    }
  | {
      readonly op: 'addShape';
      readonly container: AutomationRef;
      readonly shape: AutomationNewShape;
      readonly box?: AutomationBox;
    }
  | { readonly op: 'deleteShape'; readonly shape: AutomationRef }
  | { readonly op: 'setShape'; readonly shape: AutomationRef; readonly props: AutomationShapeWrite }
  | { readonly op: 'setZOrder'; readonly shape: AutomationRef; readonly position: AutomationZOrder }
  | { readonly op: 'group'; readonly shapes: readonly AutomationRef[] }
  | { readonly op: 'ungroup'; readonly group: AutomationRef }
  | { readonly op: 'getParentGroup'; readonly shape: AutomationRef; readonly orNull?: boolean }
  | {
      readonly op: 'setShapeHyperlink';
      readonly shape: AutomationRef;
      readonly address: string | null;
      readonly screenTip?: string;
    }
  | {
      readonly op: 'parentOf';
      readonly shape: AutomationRef;
      readonly kind: 'slide' | 'slideLayout' | 'slideMaster';
      readonly orNull?: boolean;
    }
  // Fill and line
  | { readonly op: 'readShapeFormat'; readonly shape: AutomationRef }
  | {
      readonly op: 'patchShapeFormat';
      readonly shape: AutomationRef;
      readonly patch: ShapeFormatPatch;
    }
  | {
      readonly op: 'readFill';
      readonly shape: AutomationRef;
      /** A table cell's fill (`a:tcPr`) instead of the shape's. */
      readonly cell?: AutomationCellAddress;
    }
  | {
      readonly op: 'setFill';
      readonly shape: AutomationRef;
      readonly fill: AutomationFillWrite;
      readonly cell?: AutomationCellAddress;
    }
  // Backgrounds and theme colors. The owner is a slide, a layout, or a master.
  | { readonly op: 'readBackground'; readonly owner: AutomationRef }
  | {
      readonly op: 'setBackground';
      readonly owner: AutomationRef;
      readonly background: AutomationBackgroundWrite;
    }
  /** A color of the theme's `a:clrScheme`: `dk1`, `lt1`, … `accent6`, `hlink`, `folHlink`. */
  | { readonly op: 'readThemeColor'; readonly owner: AutomationRef; readonly slot: string }
  | {
      readonly op: 'setThemeColor';
      readonly owner: AutomationRef;
      readonly slot: string;
      readonly color: string;
    }
  | { readonly op: 'readLine'; readonly shape: AutomationRef }
  | { readonly op: 'setLine'; readonly shape: AutomationRef; readonly line: AutomationLineWrite }
  // Text
  | {
      readonly op: 'readTextFrame';
      readonly shape: AutomationRef;
      /** Answer null, not a refusal, when the shape cannot hold text (a picture, a group). */
      readonly orNull?: boolean;
    }
  // Whether the text fits its text area, from the layout that the shape's picture draws.
  | { readonly op: 'readTextFit'; readonly shape: AutomationRef }
  | {
      readonly op: 'setTextFrame';
      readonly shape: AutomationRef;
      readonly props: AutomationTextFrameWrite;
    }
  | {
      readonly op: 'readText';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly cell?: AutomationCellAddress;
      /** Answer null, not a refusal, when the shape cannot hold text (a picture, a group). */
      readonly orNull?: boolean;
    }
  | {
      readonly op: 'setText';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly text: string;
      readonly cell?: AutomationCellAddress;
    }
  | {
      readonly op: 'readFont';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly cell?: AutomationCellAddress;
      /** Answer null, not a refusal, when the shape cannot hold text (a picture, a group). */
      readonly orNull?: boolean;
    }
  | {
      /** The runs of a text body, each with its own format: `TableCell.textRuns`. */
      readonly op: 'readTextRuns';
      readonly shape: AutomationRef;
      readonly cell?: AutomationCellAddress;
    }
  | {
      readonly op: 'setFont';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly font: AutomationFontWrite;
      readonly cell?: AutomationCellAddress;
    }
  | {
      readonly op: 'readParagraph';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly cell?: AutomationCellAddress;
      /** Answer null, not a refusal, when the shape cannot hold text (a picture, a group). */
      readonly orNull?: boolean;
    }
  | {
      readonly op: 'setParagraph';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      readonly paragraph: AutomationParagraphWrite;
      readonly cell?: AutomationCellAddress;
    }
  | {
      readonly op: 'setHyperlink';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
      /** Null removes the link. */
      readonly address: string | null;
      readonly screenTip?: string;
    }
  | {
      /**
       * The text links of a shape's text that overlap a span, or of all its text without one.
       * An empty span answers the link that holds the point.
       */
      readonly op: 'listTextHyperlinks';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
    }
  // Tags. The owner is the presentation, a slide, or a shape.
  | { readonly op: 'readTags'; readonly owner: AutomationRef }
  | {
      readonly op: 'setTag';
      readonly owner: AutomationRef;
      readonly key: string;
      readonly value: string;
    }
  | { readonly op: 'deleteTag'; readonly owner: AutomationRef; readonly key: string }
  // Tables
  | { readonly op: 'readTable'; readonly shape: AutomationRef }
  | {
      readonly op: 'readTableCell';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
    }
  | {
      readonly op: 'setTableCell';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
      readonly text: string;
    }
  | {
      readonly op: 'insertTableRows';
      readonly shape: AutomationRef;
      /** The row the new rows go before; the end when absent. */
      readonly index?: number;
      readonly count: number;
    }
  | {
      readonly op: 'insertTableColumns';
      readonly shape: AutomationRef;
      readonly index?: number;
      readonly count: number;
    }
  | {
      readonly op: 'deleteTableRows';
      readonly shape: AutomationRef;
      readonly index: number;
      readonly count: number;
    }
  | {
      readonly op: 'deleteTableColumns';
      readonly shape: AutomationRef;
      readonly index: number;
      readonly count: number;
    }
  | {
      readonly op: 'mergeTableCells';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
      readonly rowCount: number;
      readonly columnCount: number;
    }
  | {
      /** `a:tcPr/@anchor` and `@anchorCtr` of the cell's block (ECMA-376 Part 1, 21.1.3.17). */
      readonly op: 'setTableCellAnchor';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
      /** `ST_TextAnchoringType`: `t`, `ctr`, or `b`. */
      readonly anchor: string;
      readonly anchorCenter: boolean;
    }
  | {
      /** Split the cell's block into a grid of `rowCount` by `columnCount` cells. */
      readonly op: 'splitTableCell';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
      readonly rowCount: number;
      readonly columnCount: number;
    }
  | {
      /** Make the cell's block span `rowCount` rows and `columnCount` columns from its anchor. */
      readonly op: 'resizeTableCell';
      readonly shape: AutomationRef;
      readonly row: number;
      readonly column: number;
      readonly rowCount: number;
      readonly columnCount: number;
    }
  | {
      /** Remove the text, the direct format, or both, of every cell. */
      readonly op: 'clearTable';
      readonly shape: AutomationRef;
      readonly text: boolean;
      readonly format: boolean;
    }
  | {
      /** The borders and margins of the cell's block. */
      readonly op: 'readTableCellFormat';
      readonly shape: AutomationRef;
      readonly cell: AutomationCellAddress;
    }
  | {
      /**
       * One line of the cell's block. A field the write leaves out keeps the line the cell
       * shows. An outside edge is also written on the cells across it.
       */
      readonly op: 'setTableCellBorder';
      readonly shape: AutomationRef;
      readonly cell: AutomationCellAddress;
      readonly side: AutomationCellBorderSide;
      readonly line: AutomationLineWrite;
    }
  | {
      readonly op: 'setTableCellMargins';
      readonly shape: AutomationRef;
      readonly cell: AutomationCellAddress;
      readonly margins: AutomationCellMarginsWrite;
    }
  | {
      readonly op: 'setTableStyle';
      readonly shape: AutomationRef;
      readonly style: AutomationTableStyleWrite;
    }
  // Preset geometry adjust values. The shape is a shape, a picture, or a connector.
  | { readonly op: 'readAdjustments'; readonly shape: AutomationRef }
  | {
      /** One adjust value by its zero-based position, in the units `readAdjustments` gives. */
      readonly op: 'setAdjustment';
      readonly shape: AutomationRef;
      readonly index: number;
      readonly value: number;
    }
  | {
      readonly op: 'setTableSize';
      readonly shape: AutomationRef;
      /** A row height or a column width, in points. */
      readonly target: 'row' | 'column';
      readonly index: number;
      readonly size: number;
    }
  // Comments of a slide: both formats (`plan-comments.ts`). `comment` is an id key that
  // `readComments` answered.
  | { readonly op: 'readComments'; readonly slide: AutomationRef }
  | {
      /**
       * Add a thread in the format the deck uses: threaded, unless the deck has only legacy
       * comments (`p:cmLst`, ECMA-376 Part 1, 19.4.4). Answers the new comment's id key.
       */
      readonly op: 'addComment';
      readonly slide: AutomationRef;
      readonly text: string;
      /** A shape on the top level of the slide that the thread points at. */
      readonly shape?: AutomationRef;
      /** With `shape`: the stretch of its text the thread points at (threaded format only). */
      readonly span?: AutomationTextSpan;
      /** The author's name. Absent: the host's author. */
      readonly author?: string;
      readonly initials?: string;
    }
  | {
      /** Reply to the thread of `comment`. Answers the reply's id key. */
      readonly op: 'replyToComment';
      readonly slide: AutomationRef;
      readonly comment: string;
      readonly text: string;
      readonly author?: string;
      readonly initials?: string;
      readonly expect?: AutomationCommentExpect;
    }
  | {
      readonly op: 'setCommentText';
      readonly slide: AutomationRef;
      readonly comment: string;
      readonly text: string;
      readonly expect?: AutomationCommentExpect;
    }
  | {
      /** Resolve or reopen a thread. The legacy format has no state and refuses. */
      readonly op: 'setCommentResolved';
      readonly slide: AutomationRef;
      readonly comment: string;
      readonly resolved: boolean;
      readonly expect?: AutomationCommentExpect;
    }
  | {
      /** Delete a reply, or a thread with its replies. */
      readonly op: 'deleteComment';
      readonly slide: AutomationRef;
      readonly comment: string;
      readonly expect?: AutomationCommentExpect;
    }
  // Selection (browser host only)
  | { readonly op: 'readSelection' }
  | { readonly op: 'selectSlides'; readonly slides: readonly AutomationRef[] }
  | { readonly op: 'selectShapes'; readonly shapes: readonly AutomationRef[] }
  | {
      readonly op: 'selectText';
      readonly shape: AutomationRef;
      readonly span?: AutomationTextSpan;
    }
  // Speaker notes: the body placeholder of the slide's notes slide (`p:notes`, ECMA-376 Part 1,
  // 19.3.1.26). Paragraphs and line breaks read as `\n`.
  | { readonly op: 'readSlideNotes'; readonly slide: AutomationRef }
  | {
      /**
       * Set the notes text, one paragraph for each line. A slide with no notes slide gets one
       * from the notes master (Part 1, 13.3.5 and 13.3.4) when the text is not empty.
       */
      readonly op: 'setSlideNotes';
      readonly slide: AutomationRef;
      readonly text: string;
    }
  // Sections: the `p14:sectionLst` in the presentation's `p:extLst` (ECMA-376 Part 1, 19.2.1.12).
  | { readonly op: 'listSections' }
  | { readonly op: 'getSection'; readonly key: AutomationItemKey; readonly orNull?: boolean }
  | { readonly op: 'readSection'; readonly section: AutomationRef }
  | {
      /**
       * Add a section that starts at `slide` and runs to the end of the section that holds it.
       * In a presentation without sections, the slides before `slide` go into a first section
       * named "Default Section". Answers the handle of the new section.
       */
      readonly op: 'addSection';
      readonly name: string;
      readonly slide: AutomationRef;
    }
  | { readonly op: 'renameSection'; readonly section: AutomationRef; readonly name: string }
  | {
      /** Move a section with its slides to a zero-based position in the section list. */
      readonly op: 'moveSection';
      readonly section: AutomationRef;
      readonly index: number;
    }
  | {
      /**
       * Remove a section. Its slides join the section before it (after it, for the first
       * section), or are deleted with `deleteSlides`.
       */
      readonly op: 'deleteSection';
      readonly section: AutomationRef;
      readonly deleteSlides?: boolean;
    }
  | {
      /** Remove every section. The slides stay. */
      readonly op: 'deleteAllSections';
    }
  // The transition of a slide (`p:transition`, ECMA-376 Part 1, §19.3.1.50).
  | { readonly op: 'readTransition'; readonly slide: AutomationRef }
  | {
      readonly op: 'setTransition';
      readonly slide: AutomationRef;
      readonly transition: AutomationTransitionWrite;
    }
  // Slide number, footer, and date (the placeholders `sldNum`, `ftr`, and `dt`).
  | { readonly op: 'readHeaderFooter'; readonly slide: AutomationRef }
  | {
      readonly op: 'setHeaderFooter';
      /** The slides to change. Every slide when absent. */
      readonly slides?: readonly AutomationRef[];
      readonly settings: AutomationHeaderFooterWrite;
    };

/** The operations that write. Every other operation only reads. */
export const AUTOMATION_COMMANDS: ReadonlySet<AutomationOperation['op']> = new Set([
  'setPageSetup',
  'setProperties',
  'setCustomProperty',
  'deleteCustomProperty',
  'deleteCustomProperties',
  'addCustomXmlPart',
  'setCustomXml',
  'deleteCustomXmlPart',
  'addSlide',
  'deleteSlide',
  'duplicateSlide',
  'moveSlide',
  'setSlideHidden',
  'applyLayout',
  'insertSlides',
  'replaceText',
  'addBinding',
  'addBindingFromSelection',
  'deleteBinding',
  'addShape',
  'deleteShape',
  'setShape',
  'setZOrder',
  'group',
  'ungroup',
  'setFill',
  'patchShapeFormat',
  'setLine',
  'setBackground',
  'setThemeColor',
  'setTextFrame',
  'setText',
  'setFont',
  'setParagraph',
  'setHyperlink',
  'setShapeHyperlink',
  'setTag',
  'deleteTag',
  'setTableCell',
  'insertTableRows',
  'insertTableColumns',
  'deleteTableRows',
  'deleteTableColumns',
  'mergeTableCells',
  'setTableCellAnchor',
  'splitTableCell',
  'resizeTableCell',
  'clearTable',
  'setTableSize',
  'setTableCellBorder',
  'setTableCellMargins',
  'setTableStyle',
  'setAdjustment',
  'setSlideNotes',
  'addSection',
  'renameSection',
  'moveSection',
  'deleteSection',
  'deleteAllSections',
  'setTransition',
  'setHeaderFooter',
  'addComment',
  'replyToComment',
  'setCommentText',
  'setCommentResolved',
  'deleteComment',
]);

/** The operations that need a reader's selection. */
export const AUTOMATION_SELECTION_OPERATIONS: ReadonlySet<AutomationOperation['op']> = new Set([
  'readSelection',
  'addBindingFromSelection',
  'selectSlides',
  'selectShapes',
  'selectText',
]);
