// The DOM-free part of the `contracts` lane: declaration-only public types shared by the
// editor, its adapters, and the automation lane. Nothing here names a DOM type, so the DOM-free
// automation lane imports this module, never `index.ts`, and `check:lane-boundaries` compiles it
// without the DOM library. The one DOM-typed contract, `Editor`, lives in `editor.ts`.

export type {
  EditorFontErrorCode,
  FontConfiguration,
  FontFaceRequest,
  FontSource,
  FontSourceSubstitution,
} from './editor-fonts.ts';

export type { EditorExecOptions, HistoryGroup, HistoryGroupOptions } from './history-group.ts';
import type { EditorRejection, EditorRemoval } from './notices.ts';
export type { EditorRejection, EditorRemoval } from './notices.ts';

import type {
  ObjectAxis,
  ShapeEffectChoice,
  ShapeStyleChoice,
  ShapeEffectsState,
  ShapeSummary,
} from './shape-format.ts';

import type {
  TextFormatCommand,
  BulletFormatRead,
  TextOrientation,
  TextColumns,
  FontFormatRead,
  ParagraphFormatRead,
} from './text-format.ts';

export type {
  TextFormatCommand,
  AutonumberScheme,
  BulletFormat,
  BulletFormatRead,
  TextOrientation,
  TextColumns,
  TextCase,
  UnderlineStyle,
  StrikeStyle,
  CapsStyle,
  FontFormatCommand,
  FontFormatRead,
  ParagraphAlignment,
  ParagraphLineSpacing,
  ParagraphFormatCommand,
  ParagraphFormatRead,
} from './text-format.ts';

import type { TableBorderPen, TableCellAddress, TableContext, TableCommand } from './table.ts';

export type {
  TableCellAddress,
  TableBorderTarget,
  TableBorderEdgeTarget,
  TableBorderPen,
  TableBorderSpec,
  TableStyleOption,
  TableStyleChoice,
  TableContext,
  TableCommand,
} from './table.ts';

export type {
  ObjectAxis,
  ShadowPreset,
  ReflectionPreset,
  ShapeEffectChoice,
  ShapeStyleRow,
  ShapeStyleColor,
  ShapeStyleValue,
  ShapeEffectsValue,
  RotateChoice,
  ShapeStyleChoice,
  ShapeEffectsState,
  ShapeSummary,
} from './shape-format.ts';

import type { PictureCommand, SelectedImageState, AltText } from './picture.ts';
import type { FillCommand, GradientFill, PatternFill, SlideBackground } from './fill.ts';
import type { SlideBackgroundState } from './fill.ts';

export type {
  FillCommand,
  GradientFill,
  PatternFill,
  PatternPreset,
  SlideBackground,
  SlideBackgroundState,
} from './fill.ts';

export type {
  ImageCropPercent,
  ImageCrop,
  ImageRecolor,
  ImageProperties,
  ImageAdjustments,
  ImageStyleId,
  SelectedImageState,
  AltText,
  PictureCommand,
} from './picture.ts';

/** The editing mode, matching the toolbar's mode pill. @public */
export type EditorMode = 'edit' | 'view';

/** Paragraph alignment values the chrome can set. @public */
export type TextAlign = 'l' | 'ctr' | 'r' | 'just';

/** Object alignment targets. @public */
export type ObjectAlign = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';

/** Z-order moves. @public */
export type OrderMove = 'forward' | 'backward' | 'front' | 'back';

/** Every editing command. Each compiles to store ops or is refused with a reason. @public */
export type EditorCommand =
  | { readonly type: 'undo' }
  | { readonly type: 'redo' }
  | { readonly type: 'toggleBold' }
  | { readonly type: 'toggleItalic' }
  | { readonly type: 'toggleUnderline' }
  | { readonly type: 'toggleStrike' }
  | { readonly type: 'setFontSize'; readonly points: number }
  | { readonly type: 'stepFontSize'; readonly direction: 1 | -1 }
  | { readonly type: 'setFontFamily'; readonly family: string }
  /** Text color: a color choice, or null for the inherited color. */
  | { readonly type: 'setTextColor'; readonly color: ColorChoice | null }
  | { readonly type: 'setHighlight'; readonly color: ColorChoice | null }
  | { readonly type: 'setAlign'; readonly align: TextAlign }
  | { readonly type: 'toggleBullets' }
  | { readonly type: 'toggleNumbering' }
  | { readonly type: 'indent'; readonly direction: 1 | -1 }
  /**
   * Backspace at the start of a list item: remove the bullet or number of every list item the
   * selection touches. The item keeps its level, and its text moves to where the bullet was.
   */
  | { readonly type: 'removeBullet' }
  | { readonly type: 'setLineSpacing'; readonly multiple: number }
  | { readonly type: 'insertTextBox' }
  | { readonly type: 'insertShape'; readonly preset: string }
  | {
      readonly type: 'insertImage';
      readonly bytes: Uint8Array;
      readonly mime: string;
      readonly width: number;
      readonly height: number;
    }
  | { readonly type: 'order'; readonly move: OrderMove }
  | { readonly type: 'alignObjects'; readonly align: ObjectAlign }
  | { readonly type: 'duplicate' }
  | { readonly type: 'delete' }
  /**
   * Home > New Slide: add a slide after the current one. `layoutId` (`SlideLayoutSummary.id`)
   * picks its layout. Without it, the slide takes the current slide's layout, or Title and
   * Content after a title slide.
   */
  | { readonly type: 'addSlide'; readonly layoutId?: string }
  | { readonly type: 'duplicateSlide'; readonly index: number }
  | { readonly type: 'deleteSlide'; readonly index: number }
  /**
   * Move slide `from` to position `to`. It joins the section of the slide before it, or
   * `section`, by the one rule of `moveSlides`. Behavior change: a slide moved up to the first
   * place of a section joined that section before; now it joins the section before it. Pass
   * that section's id as `section` to get the earlier result.
   */
  | {
      readonly type: 'moveSlide';
      readonly from: number;
      readonly to: number;
      readonly section?: string;
    }
  /**
   * Insert > Slide numbers: show or hide the slide number on every slide, through the layout's
   * slide-number placeholder. The change applies to all slides, as Header and Footer > Apply to All.
   */
  | { readonly type: 'setSlideNumbers'; readonly show: boolean }
  /** Delete the selected slides (`EditorSnapshot.selectedSlides`) as one undo step. */
  | { readonly type: 'deleteSlides'; readonly indices: readonly number[] }
  /** Duplicate the selected slides; the copies follow the last one, in order. */
  | { readonly type: 'duplicateSlides'; readonly indices: readonly number[] }
  /**
   * Move the selected slides as one block, in order, so the block starts at `to`, counted
   * among the slides that do not move. The block joins the section of the slide before it.
   * `section` (a `SlideSection.id`) puts it into that section instead, for a drop below a
   * section header or a step across one in the slide pane: the section must be the
   * section of the slide before the block, of the slide after it, or an empty section between
   * them; otherwise the command is refused. With `section`, a block that stays in place only
   * changes its section.
   */
  | {
      readonly type: 'moveSlides';
      readonly indices: readonly number[];
      readonly to: number;
      readonly section?: string;
    }
  /**
   * Speaker notes: write `text` as slide `index`'s notes. With `paragraphs` (the same text as
   * formatted runs, as a notes field shows it), each character also takes the bold, italic,
   * underline, and strikethrough it has there, in the same undo step.
   */
  | {
      readonly type: 'setNotes';
      readonly index: number;
      readonly text: string;
      readonly paragraphs?: readonly NotesParagraph[];
    }
  /**
   * Speaker notes: bold, italic, underline, or strikethrough on the plain-text range
   * `[start, end)` of slide `index`'s notes (offsets as in `EditorSnapshot.notes`), as one undo
   * step. `value` absent toggles: on, unless the whole range already has the format.
   */
  | {
      readonly type: 'formatNotes';
      readonly index: number;
      readonly start: number;
      readonly end: number;
      readonly format: NotesFormat;
      readonly value?: boolean;
    }
  /**
   * Home > Section > Add Section: a section that starts at slide `index` (the current slide when
   * absent) and takes the rest of the section it splits. `name` defaults to the translated
   * "Untitled Section" (`sections.untitled`). In a deck without sections, the slides before it go
   * into a first section named `defaultSectionName`, by default "Default Section" (`sections.default`).
   */
  | {
      readonly type: 'addSection';
      readonly name?: string;
      readonly index?: number;
      readonly defaultSectionName?: string;
    }
  /** Home > Section > Remove All Sections: the slides stay, the section list goes. */
  | { readonly type: 'removeAllSections' }
  /** Home > Section > Rename Section. */
  | { readonly type: 'renameSection'; readonly sectionId: string; readonly name: string }
  /**
   * Remove Section: its slides join the section before it (after it, for the first). With
   * `withSlides`, Remove Section & Slides deletes them. The last section removes the list.
   * `sectionId` absent: the section of the current slide.
   */
  | { readonly type: 'removeSection'; readonly sectionId?: string; readonly withSlides?: boolean }
  /** Move Section Up (`-1`) or Down (`1`), with its slides. */
  | { readonly type: 'moveSection'; readonly sectionId: string; readonly direction: -1 | 1 }
  /**
   * Home > Layout: put the slides on a layout of the deck (`SlideLayoutSummary.id`) as one undo
   * step. Placeholders re-map by type and index, and their content is kept.
   * `indices` absent: the selected slides (`EditorSnapshot.selectedSlides`).
   */
  | {
      readonly type: 'applyLayout';
      readonly layoutId: string;
      readonly indices?: readonly number[];
    }
  /**
   * Home > Reset: re-apply each slide's layout. Text placeholders go back to the layout's
   * position, size, and formatting, and deleted placeholders come back; their text stays.
   * Pictures, tables, and charts in placeholders keep their own box. `indices` absent: the
   * selected slides.
   */
  | { readonly type: 'resetSlides'; readonly indices?: readonly number[] }
  /**
   * Slide Show > Hide Slide: hide the slides (`p:sld/@show="0"`) or show them again. `indices`
   * absent: the selected slides. `hidden` absent: the toggle, which hides them unless every one
   * of them is hidden already.
   */
  | {
      readonly type: 'setSlidesHidden';
      readonly indices?: readonly number[];
      readonly hidden?: boolean;
    }
  /** View > Smart guides: drags snap to and show alignment guides. A view setting, not an edit. */
  | { readonly type: 'toggleSmartGuides' }
  /**
   * Set the table border pen (`snapshot.tableBorderPen`) that the Borders menu draws with. A
   * setting of this editor, not an edit: no undo step, never saved.
   */
  | { readonly type: 'setTableBorderPen'; readonly pen: TableBorderPen }
  /** View > Normal or Slide Sorter. A view setting, not an edit: no undo step, never saved. */
  | { readonly type: 'setPresentationView'; readonly view: PresentationView }
  /**
   * Collapse or expand sections in the slide pane and the slide sorter (Collapse All, Expand All
   * when `sectionIds` is absent). A view setting, not an edit: no undo step, never saved.
   */
  | {
      readonly type: 'setSectionsCollapsed';
      readonly collapsed: boolean;
      readonly sectionIds?: readonly string[];
    }
  /**
   * Replace one find match with `replacement`. The new text takes the formatting of the first
   * replaced character. Refused when the matched text is no longer there.
   */
  | { readonly type: 'replace'; readonly match: FindMatch; readonly replacement: string }
  /** Replace every match of `query` in the deck (slides, tables, notes) as one undo step. */
  | {
      readonly type: 'replaceAll';
      readonly query: string;
      readonly replacement: string;
      readonly options?: FindOptions;
    }
  | {
      readonly type: 'setTransform';
      readonly nodeId: string;
      readonly x: number;
      readonly y: number;
      readonly cx: number;
      readonly cy: number;
      readonly rot: number;
    }
  | { readonly type: 'nudge'; readonly dx: number; readonly dy: number }
  /** Shape fill: a color choice, or null for no fill. */
  | { readonly type: 'setShapeFill'; readonly color: ColorChoice | null }
  /** Shape fill: a linear or radial gradient. */
  | { readonly type: 'setShapeGradient'; readonly gradient: GradientFill }
  /** Outline color: a color choice, or null for no outline. */
  | { readonly type: 'setLineColor'; readonly color: ColorChoice | null }
  /** Outline width in points. */
  | { readonly type: 'setLineWidth'; readonly points: number }
  | { readonly type: 'setLineDash'; readonly dash: LineDash }
  /** Link the selected text to a URL (`http`, `https`, `mailto`, `tel`, `ftp`); null unlinks. */
  | { readonly type: 'setLink'; readonly url: string | null }
  /** Remove character formatting from the selected text; links stay. */
  | { readonly type: 'clearFormatting' }
  /** Vertical alignment of text in its shape (`a:bodyPr/@anchor`). */
  | { readonly type: 'setVerticalAlign'; readonly anchor: VerticalAlign }
  /** Arrowheads of a line: `head` at its start, `tail` at its end. Absent keeps the current end. */
  | { readonly type: 'setLineEnds'; readonly head?: LineEndType; readonly tail?: LineEndType }
  /**
   * The slide background: a color, a gradient, or a picture; null removes the slide's own
   * background so the layout's shows. `allSlides` applies it to every slide.
   */
  | {
      readonly type: 'setSlideBackground';
      readonly background: SlideBackground | null;
      readonly allSlides?: boolean;
    }
  /** Slide size in EMU, 1 to 56 in a side; `content: 'scale'` (default) fits the content. */
  | {
      readonly type: 'setSlideSize';
      readonly width: number;
      readonly height: number;
      readonly content?: 'scale' | 'keep';
    }
  /** Adjust values of one preset shape by guide name, in guide units (see `a:gd`). */
  | { readonly type: 'setAdjustments'; readonly values: Readonly<Record<string, number>> }
  /** Insert a line, an arrow, or an elbow or curved connector. */
  | { readonly type: 'insertLine'; readonly kind: LineKind }
  /** Text fitting: none, shrink text on overflow, or resize the shape to fit its text. */
  | { readonly type: 'setAutofit'; readonly mode: TextFitting }
  | {
      /** Several transforms in one transaction. Values are in each element's file coordinates. */
      readonly type: 'setTransforms';
      readonly items: readonly TransformItem[];
    }
  | { readonly type: 'group' }
  | { readonly type: 'ungroup' }
  | { readonly type: 'paste'; readonly payload: ClipboardPayload }
  /**
   * The Format Painter button. A press copies the formatting of the selection and arms the
   * painter for one paint; a second press within the double-click time keeps it on until
   * Escape (`sticky` does that at once); a press while it is on turns it off.
   */
  | { readonly type: 'toggleFormatPainter'; readonly sticky?: boolean }
  /** Copy the formatting of the selection without arming the painter (Ctrl/Cmd+Shift+C). */
  | { readonly type: 'copyFormat' }
  /** Apply the copied formatting to the selection as one undo step (Ctrl/Cmd+Shift+V). */
  | { readonly type: 'pasteFormat' }
  | TextFormatCommand
  | PictureCommand
  | FillCommand
  /**
   * Shape Effects: set or remove one effect of the selected shapes, connectors, and pictures.
   * A shape whose effects come from its theme style (`a:effectRef`) gets that style's effects
   * written into it first, so its other effects stay.
   */
  | { readonly type: 'setShapeEffect'; readonly effect: ShapeEffectChoice }
  /**
   * Shape Styles: apply a theme style of the Shape Styles gallery to the selected shapes. It
   * writes `p:style` and removes the shape's own fill, outline, and effects.
   */
  | { readonly type: 'setShapeStyle'; readonly style: ShapeStyleChoice }
  /**
   * Distribute the selected elements so the gaps between them are equal. One element centers
   * on the slide; two are refused; three or more keep the outer two in place.
   */
  | { readonly type: 'distributeObjects'; readonly axis: ObjectAxis }
  /** Flip each selected element about its own center (`a:xfrm/@flipH`, `@flipV`). */
  | { readonly type: 'flipObjects'; readonly axis: ObjectAxis }
  /** Turn each selected element by `degrees` about its center; 90 is Rotate Right 90°. */
  | { readonly type: 'rotateObjects'; readonly degrees: number }
  /** Lock or unlock the aspect ratio of the selected elements (`noChangeAspect`). */
  | { readonly type: 'setLockAspectRatio'; readonly lock: boolean }
  /** Selection pane: show or hide elements of the current slide (`p:cNvPr/@hidden`). */
  | {
      readonly type: 'setShapesVisible';
      readonly nodeIds: readonly string[];
      readonly visible: boolean;
    }
  /** Selection pane: rename an element of the current slide (`p:cNvPr/@name`). */
  | { readonly type: 'renameShape'; readonly nodeId: string; readonly name: string }
  | TableCommand;

/**
 * The format painter: off, armed for one paint (`once`), or on until Escape or a second press
 * (`sticky`). @public
 */
export type FormatPainterMode = 'off' | 'once' | 'sticky';

/** How find compares text. @public */
export interface FindOptions {
  /** Match upper and lower case exactly. */
  readonly matchCase?: boolean;
  /** Match whole words only: no letter, digit, or underscore touches either end. */
  readonly wholeWord?: boolean;
}

/** One match of a find. Offsets are UTF-16 offsets in the paragraph. @public */
export interface FindMatch {
  /** The slide index. */
  readonly slide: number;
  /** `slide` for text on the slide (shapes, groups, table cells); `notes` for speaker notes. */
  readonly where: 'slide' | 'notes';
  /** The part that holds the text: the slide part, or the slide's notes part. */
  readonly part: string;
  /** The element the text belongs to: a `p:sp`, or the `p:graphicFrame` of a table. */
  readonly nodeId: string;
  /** The text body: a shape's `p:txBody` or a table cell's `a:txBody`. */
  readonly txBodyId: string;
  readonly paragraph: number;
  readonly start: number;
  readonly end: number;
  /** The text the match covers. */
  readonly text: string;
}

/**
 * A theme color as the color menus name it: the color map's names, so `bg1` is the
 * color the slide's `p:clrMap` maps to Background 1. @public
 */
export type ThemeColorSlot =
  | 'bg1'
  | 'tx1'
  | 'bg2'
  | 'tx2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6';

/**
 * A theme color pick. It is written as `a:schemeClr`, so the color follows the theme when the
 * theme changes. `shift` is the luminance variant of {@link ThemeSwatch}. @public
 */
export interface ThemeColorChoice {
  readonly theme: ThemeColorSlot;
  /** 0 or absent for the theme color; −99 to 99 for "Darker n%" or "Lighter n%". */
  readonly shift?: number;
}

/** A color the chrome writes: `#RRGGBB` (`a:srgbClr`) or a theme color (`a:schemeClr`). @public */
export type ColorChoice = string | ThemeColorChoice;

/** One swatch of the theme color grid. @public */
export interface ThemeSwatch {
  readonly slot: ThemeColorSlot;
  /** The color as a slide renders it, `#RRGGBB`. */
  readonly color: string;
  /**
   * The luminance variant: 0 is the theme color itself, a positive number is "Lighter n%"
   * (`a:lumMod` 100−n, `a:lumOff` n), and a negative number is "Darker n%" (`a:lumMod` 100−n).
   */
  readonly shift: number;
  /**
   * The theme color the slot names through the slide's color map (`p:clrMap`): `lt1` for
   * `bg1` with the default map, `dk1` on a dark map. A `p:style` reference to `dk1` or `lt1`
   * names this theme color directly, not through the map.
   */
  readonly scheme?: string;
}

/** Vertical text alignment: top, middle, bottom. @public */
export type VerticalAlign = 't' | 'ctr' | 'b';

/** Line end (arrowhead) types (`a:headEnd/@type`, `a:tailEnd/@type`). @public */
export type LineEndType = 'none' | 'triangle' | 'arrow' | 'stealth' | 'oval' | 'diamond';

/** Lines the chrome inserts. @public */
export type LineKind = 'line' | 'arrow' | 'elbow' | 'curved';

/** Outline dash presets offered by the chrome (`a:prstDash` values). @public */
export type LineDash = 'solid' | 'sysDot' | 'sysDash' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot';

/** Text fitting modes (`a:noAutofit`, `a:normAutofit`, `a:spAutoFit`). @public */
export type TextFitting = 'none' | 'shrink' | 'shape';

/** One element's new box, in its file coordinates (its parent's child space). @public */
export interface TransformItem {
  readonly nodeId: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
  /** Flips; absent keeps the current ones. A line's direction is its flips. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/** A serialized element, structurally the store's `NodeSpec`. @public */
export interface ClipboardNode {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly attrs?: readonly { ns: string; local: string; prefix?: string; value: string }[];
  readonly nsDecls?: readonly { prefix: string; uri: string }[];
  readonly children?: readonly (ClipboardNode | { readonly text: string })[];
}

/** A relationship a copied shape uses, with the bytes of its target. @public */
export interface ClipboardRelationship {
  readonly id: string;
  readonly type: string;
  readonly part: string;
  readonly contentType: string;
  readonly base64: string;
}

/** The clipboard format for shapes (`application/x-pptx-editor+json`). @public */
export interface ClipboardPayload {
  readonly version: 1;
  /** The session the shapes came from; pasting into it reuses its parts. */
  readonly document: string;
  /** The slide part the shapes came from; pasting onto it offsets them. */
  readonly slidePart: string;
  readonly shapes: readonly ClipboardNode[];
  readonly relationships: readonly ClipboardRelationship[];
}

/** The command type names. @public */
export type EditorCommandType = EditorCommand['type'];

/** Whether a command can run now; when it cannot, the i18n key of the reason. @public */
export type CanResult = { readonly ok: true } | { readonly ok: false; readonly reason: string };

/**
 * The outcome of a command. A refusal with `noop` means the command had nothing to do (nothing
 * to undo, nothing selected): nothing changed, and nothing went wrong. @public
 */
export type CommandResult =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly reason: string;
      readonly detail?: string;
      readonly noop?: boolean;
    };

/** What is selected. @public */
export type EditorSelection =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'shapes';
      readonly nodeIds: readonly string[];
      /** The group whose children are selected; absent for the slide's top level. */
      readonly group?: string;
    }
  | {
      readonly kind: 'text';
      readonly nodeId: string;
      readonly anchor: TextPoint;
      readonly focus: TextPoint;
      /** The table cell whose text is edited; `nodeId` is then the table's graphic frame. */
      readonly cell?: TableCellAddress;
    }
  | {
      /** A range of table cells, from the cell the drag started in to the cell it is over. */
      readonly kind: 'cells';
      /** The table's graphic frame. */
      readonly nodeId: string;
      readonly anchor: TableCellAddress;
      readonly focus: TableCellAddress;
    };

/** A caret position: paragraph index and UTF-16 offset in the paragraph. @public */
export interface TextPoint {
  readonly paragraph: number;
  readonly offset: number;
}

/** Formatting at the selection, for toolbar state. @public */
export interface SelectionFormat {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly fontSize?: number;
  /** `fontSize` as drawn under `a:normAutofit/@fontScale`; null when the shapes differ. */
  readonly renderedFontSize?: number | null;
  /** The family the text is drawn in; a theme font reads as the theme's family name. */
  readonly fontFamily?: string;
  /**
   * The theme font the family comes from: `major` for `+mj-lt` (headings), `minor` for `+mn-lt`
   * (body). Absent when the text names a literal family. ECMA-376 Part 1, 20.1.4.1.17 `latin`.
   */
  readonly fontFamilyTheme?: 'major' | 'minor';
  readonly color?: string;
  readonly align?: TextAlign;
  /** Vertical alignment of the text in its shape. */
  readonly verticalAlign?: VerticalAlign;
  /** The link URL of the text at the selection. */
  readonly link?: string;
  readonly bullets: boolean;
  readonly numbering: boolean;
  /** The selected paragraphs, for the Paragraph dialog and `useParagraphFormat`. */
  readonly paragraph?: ParagraphFormatRead;
  /** The text is raised (`baseline` above 0). */
  readonly superscript?: boolean;
  /** The text is lowered (`baseline` below 0). */
  readonly subscript?: boolean;
  /** The selected text, for the Font dialog and the character spacing menu. */
  readonly font?: FontFormatRead;
  /**
   * The text direction of the selected shapes with text (`vert`); absent when they differ. A
   * file value outside `ST_TextVerticalType` reads as `horizontal`, as layout draws it.
   */
  readonly textDirection?: TextOrientation;
  /** The text columns of the selected shapes with text; absent when they differ. */
  readonly columns?: TextColumns;
  /**
   * The bullet format of the selected paragraphs: null when they have no list, absent when they
   * differ (so no gallery row is checked).
   */
  readonly bulletFormat?: BulletFormatRead | null;
}

/** The selected element's geometry, for the Format panel. EMU and degrees. @public */
export interface SelectionGeometry {
  readonly kind: 'shape' | 'picture' | 'group' | 'connector' | 'graphicFrame';
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
  /** The element's node id; position and size are in its file coordinates. */
  readonly nodeId: string;
  /** Solid fill as `#RRGGBB`, `none`, `gradient`, `pattern`, `picture`, or `other`. Shapes only. */
  readonly fill?: string;
  /** The gradient, when `fill` is `gradient`. */
  readonly gradient?: GradientFill;
  /** The pattern, when `fill` is `pattern`. */
  readonly pattern?: PatternFill;
  /** Outline color as `#RRGGBB`, `none`, or `other`. */
  readonly lineColor?: string;
  /** Outline width in points. */
  readonly lineWidth?: number;
  readonly lineDash?: string;
  /** Arrowheads, for lines and connectors. */
  readonly lineHead?: LineEndType;
  readonly lineTail?: LineEndType;
  /** Text fitting, for shapes with a text body. */
  readonly textFitting?: TextFitting;
  /** The element's alt text. */
  readonly altText?: AltText;
  /** Flips of the element's transform. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /**
   * The aspect ratio is locked (`noChangeAspect`): a drag of any resize handle on the canvas
   * keeps the ratio.
   */
  readonly lockAspectRatio?: boolean;
  /** Shadow, reflection, glow, and soft edges. Shapes, connectors, and pictures only. */
  readonly effects?: ShapeEffectsState;
  /** The Shape Styles gallery cell the shape's `p:style` matches. Shapes only. */
  readonly shapeStyle?: ShapeStyleChoice;
}

/** One slide in the rail. @public */
export interface SlideSummary {
  readonly index: number;
  /**
   * The slide's identity in the open document: it stays the same when slides move, and a
   * deleted slide's id is never given to another slide.
   */
  readonly id: string;
  readonly hidden: boolean;
  /** Changes whenever the slide's rendering may change, for thumbnail caching. */
  readonly key: string;
  /** The slide's layout (`SlideLayoutSummary.id`); absent when the slide names none. */
  readonly layoutId?: string;
  /**
   * The slide is still opening (`EditorSnapshot.pendingSlides`): its thumbnail shows a
   * placeholder until it is ready. Absent once it is.
   */
  readonly pending?: true;
}

/**
 * How the editor shows the deck: `normal` (the slide with the rail and notes) or `slideSorter`
 * (every slide as a thumbnail, by section), as the View tab names them
 * (`ppViewNormal`, `ppViewSlideSorter`). @public
 */
export type PresentationView = 'normal' | 'slideSorter';

/** One section of the slide list (`p14:section`). @public */
export interface SlideSection {
  /** The section's identity: its GUID in the file. */
  readonly id: string;
  readonly name: string;
  /** The index of its first slide; for an empty section, the index its slides would take. */
  readonly firstSlide: number;
  readonly slideCount: number;
}

/**
 * What a person can do about a file that did not open, after DOCX Editor's reason groups:
 * `damaged` (the archive or its XML is broken, or the file is not a ZIP package at all),
 * `not-pptx` (the package holds no presentation, or the file is empty), `too-large` (the file
 * crosses a size or count limit), and `unknown`. @public
 */
export type DocumentOpenFailureReason = 'damaged' | 'not-pptx' | 'too-large' | 'unknown';

/** Options of `Editor.load`. @public */
export interface LoadOptions {
  /**
   * The slide to show first, zero-based; clamped to the deck. In a long deck it is the slide
   * that parses first. Default 0.
   */
  readonly startSlide?: number;
}

/**
 * A file the editor refused to open. It carries a stable code only, never the file's content
 * or name. Each refusal is a new object, so a host can tell a second refusal of the same file
 * from the first. @public
 */
export interface DocumentOpenFailure {
  /** The engine's reason, the same value as `EditorSnapshot.error` (`zip`, `xml`, `empty`). */
  readonly code: string;
  /** The group of `code` that picks the message. */
  readonly reason: DocumentOpenFailureReason;
}

/** Character formats the speaker notes pane sets. @public */
export type NotesFormat = 'bold' | 'italic' | 'underline' | 'strike';

/** A run of the speaker notes with its character formatting. @public */
export interface NotesRun {
  /** The run's text. A line break inside a paragraph is `\n`. */
  readonly text: string;
  /** Bold, from the run or the notes text style it inherits. */
  readonly bold: boolean;
  /** Italic, from the run or the notes text style it inherits. */
  readonly italic: boolean;
  /** Underlined (any underline style). */
  readonly underline: boolean;
  /** Struck through (single or double). */
  readonly strike: boolean;
}

/** One paragraph of the speaker notes. @public */
export interface NotesParagraph {
  /** The paragraph's runs; neighbors with the same formatting are joined. */
  readonly runs: readonly NotesRun[];
}

/**
 * The kind of a slide layout, with the file's own values: `ST_SlideLayoutType` (ECMA-376 Part 1,
 * §19.7.15). `cust` is a custom layout, and also the value for a file that names none or an
 * unknown kind. The Office.js `SlideLayoutType` enum names the same idea with other values
 * (`"TitleOnly"` for `titleOnly`); the automation API keeps that name for its own type. @public
 */
export type SlideLayoutKind =
  | 'title'
  | 'tx'
  | 'twoColTx'
  | 'tbl'
  | 'txAndChart'
  | 'chartAndTx'
  | 'dgm'
  | 'chart'
  | 'txAndClipArt'
  | 'clipArtAndTx'
  | 'titleOnly'
  | 'blank'
  | 'txAndObj'
  | 'objAndTx'
  | 'objOnly'
  | 'obj'
  | 'txAndMedia'
  | 'mediaAndTx'
  | 'objOverTx'
  | 'txOverObj'
  | 'txAndTwoObj'
  | 'twoObjAndTx'
  | 'twoObjOverTx'
  | 'fourObj'
  | 'vertTx'
  | 'clipArtAndVertTx'
  | 'vertTitleAndTx'
  | 'vertTitleAndTxOverChart'
  | 'twoObj'
  | 'objAndTwoObj'
  | 'twoObjAndObj'
  | 'cust'
  | 'secHead'
  | 'twoTxTwoObj'
  | 'objTx'
  | 'picTx';

/** One slide layout of the deck, as the Layout and New Slide galleries list it. @public */
export interface SlideLayoutSummary {
  /** The layout's identity in the open document. */
  readonly id: string;
  /** The layout's name (`p:cSld/@name`, for example "Title and Content"). */
  readonly name: string;
  /** `p:sldLayout/@type`. */
  readonly type: SlideLayoutKind;
  /** The slide master the layout belongs to; galleries group layouts by master. */
  readonly masterId: string;
  /**
   * The master's name, as the gallery's group header shows it: the master's own name, or the
   * name of its theme (`a:theme/@name`).
   */
  readonly masterName: string;
}

/** A family the deck names that renders with another face. @public */
export interface EditorFontSubstitution {
  readonly family: string;
  readonly substitute: string;
  /** True when the substitute keeps the line breaks of the original font (same advance widths). */
  readonly metricCompatible: boolean;
}

/** A reference-stable view of editor state (`useSyncExternalStore` contract). @public */
/** How a click in the slide rail changes the slide selection. @public */
export type SlideSelectMode = 'replace' | 'toggle' | 'range';

export interface EditorSnapshot {
  readonly status: 'empty' | 'loading' | 'ready' | 'error';
  readonly error?: string;
  /**
   * The file the editor refused to open, or absent. A deck that was open before stays open
   * (`status` stays `ready`); with no deck, `status` is `error`. It clears when the next file
   * starts to open.
   */
  readonly openFailure?: DocumentOpenFailure | null;
  /**
   * Why the open presentation is read-only, or `null`. `'damaged-slide'`: a long deck that
   * opened and holds edits had a slide or a notes slide that did not parse. It stays on screen
   * in view mode with its edits, and its other slides go on opening; commands, undo, `save()`,
   * and automation writes are refused. `openFailure` stays `null`: the file did open. `change`
   * fires once when it is set. It clears when the next file opens, and that file opens in the
   * mode the host set (`setMode`): only `snapshot().mode` shows view meanwhile. Optional and additive: absent reads as `null`.
   */
  readonly readOnlyReason?: 'damaged-slide' | null;
  readonly revision: number;
  readonly mode: EditorMode;
  readonly slides: readonly SlideSummary[];
  /**
   * How many slides are still opening. A long deck shows its first slide and takes input while
   * the other slides parse in the background; each one's summary carries `pending` until it is
   * ready. Edits, undo, search, and save work before then: a slide that is needed parses at
   * once, and `save` parses the rest first. Moves at most every 300 ms, and `selectionChange`
   * fires when it does. Optional and additive: absent reads as 0.
   */
  readonly pendingSlides?: number;
  /**
   * Whether a deck is still on its way to its first painted slide (`status` is `loading`).
   * Show an opening screen while it holds. It turns false when the first slide paints; that
   * slide takes input while the others open (`pendingSlides`). Optional and additive: absent
   * reads as `false`.
   */
  readonly isOpening?: boolean;
  readonly currentSlide: number;
  /**
   * The slides selected in the slide rail, ascending. Always holds `currentSlide`. Shift+click
   * selects a range and Ctrl/Cmd+click adds or removes one slide (`selectSlide`).
   */
  readonly selectedSlides: readonly number[];
  /** Whether the current slide shows its slide number (Insert > Slide numbers). */
  readonly slideNumbers: boolean;
  /** The current slide's background. */
  readonly slideBackground: SlideBackgroundState;
  /** Slide size in EMU. */
  readonly slideWidth: number;
  readonly slideHeight: number;
  readonly zoom: number;
  readonly fitZoom: boolean;
  /**
   * How many objects (shapes, lines, text boxes, pictures, tables) this editor inserted. A
   * ribbon opens the new object's Format tab when the count grows.
   */
  readonly insertions: number;
  /** Whether drags snap to smart guides (View > Smart guides). On by default. */
  readonly smartGuides: boolean;
  /**
   * View > Normal or Slide Sorter. Optional: absent means the implementation has only the
   * normal view.
   */
  readonly presentationView?: PresentationView;
  /** The ids of the collapsed sections (`SlideSection.id`). A view setting, never saved. */
  readonly collapsedSections?: readonly string[];
  readonly selection: EditorSelection;
  readonly format: SelectionFormat | null;
  readonly geometry: SelectionGeometry | null;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly notes: string;
  /**
   * The current slide's notes as paragraphs of formatted runs; their text equals `notes`.
   * Reference-stable while the notes do not change. Optional: absent means the implementation
   * does not report notes formatting.
   */
  readonly notesFormatted?: readonly NotesParagraph[];
  readonly embeddedFontFaces: number;
  /** Named families rendered with another face; reference-stable while unchanged. */
  readonly substitutedFonts: readonly EditorFontSubstitution[];
  /**
   * The current slide's theme colors in the layout of the color menus: the first row
   * is the ten theme colors (Background 1, Text 1, Background 2, Text 2, Accent 1–6), and each
   * later row one lighter or darker variant of every column. Empty with no document open;
   * reference-stable while the theme and color map do not change.
   */
  readonly themeColors: readonly (readonly ThemeSwatch[])[];
  /**
   * Why the last edit was refused, or null. Keyboard input, the clipboard, and commands set it
   * when the editor does not apply them; the next applied edit clears it. Optional: absent
   * means the implementation does not report refusals.
   */
  readonly lastRejection?: EditorRejection | null;
  /**
   * The last Delete of shapes or slides, while it is the newest edit: chrome offers to undo it.
   * The next applied edit, an undo or redo, and opening a document clear it. Cut sets none.
   * Optional: absent means the implementation does not report removals.
   */
  readonly lastRemoval?: EditorRemoval | null;
  /**
   * The format painter. The surface shows the paintbrush cursor and the next selection made
   * with the pointer takes the copied formatting while it is not `off`. Optional: absent means
   * the implementation has no format painter.
   */
  readonly formatPainter?: FormatPainterMode;
  /**
   * The table border pen: the line the Borders menu draws and the target it last drew
   * (`setTableBorderPen`). Optional: absent means the implementation has no border pen.
   */
  readonly tableBorderPen?: TableBorderPen;
  /**
   * The selected pictures, while every selected element is a picture; the picture controls
   * show then. Optional: absent means the implementation does not report pictures.
   */
  readonly image?: SelectedImageState | null;
  /**
   * The current slide's elements, top first, for the selection pane. Hidden elements are in
   * the list. Optional: absent means the implementation has no selection pane.
   */
  readonly shapes?: readonly ShapeSummary[];
  /**
   * The deck's slide layouts, masters in order and each master's layouts in order.
   * Reference-stable while they do not change. Optional: absent means the implementation does
   * not list layouts.
   */
  readonly layouts?: readonly SlideLayoutSummary[];
  /**
   * The deck's sections, in order; empty when it has none. Reference-stable while they do not
   * change. Optional: absent means the implementation does not report sections.
   */
  readonly sections?: readonly SlideSection[];
  /**
   * The table of the selection (a caret in a cell, a cell range, or the selected table frame),
   * or null. Optional: absent means the implementation does not edit tables.
   */
  readonly table?: TableContext | null;
}
