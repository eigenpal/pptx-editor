// The resolved slide model: what the cascade hands to layout. Every value here is final —
// inheritance, theme references, color maps, and style matrices are already applied — so
// layout and paint never consult the tree for a property again.

import type { GeometryPath } from '@eigenpal/ooxml-core';
import type { Element } from '../store/index.ts';
import type { ResolvedChartPart } from './chart-part.ts';
import type { Rgba } from './color.ts';
import type { ResolvedEffects, ResolvedFill, ResolvedLine } from './paint-props.ts';

/** Where an element comes from. Only `slide` elements are editable in v0. */
export type Layer = 'master' | 'layout' | 'slide';

/** A 2D transform in EMU, as stored in `a:xfrm`. `rot` is in degrees. */
export interface Xfrm {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/** Shape geometry: a preset with adjust values, or a custom geometry element. */
export type GeometrySpec =
  | {
      readonly kind: 'preset';
      readonly preset: string;
      readonly adjust: ReadonlyMap<string, number>;
    }
  | { readonly kind: 'custom'; readonly element: Element }
  /**
   * Paths already evaluated in the element's local box (EMU). Never read from a file: a drawing
   * renderer emits them, for example the bars and lines of a chart.
   */
  | { readonly kind: 'paths'; readonly paths: readonly GeometryPath[] };

/** Bullet of a paragraph. Sizes: `pct` of the first run's size, or `pts` in hundredths. */
export type ResolvedBullet =
  | {
      readonly type: 'char';
      readonly char: string;
      readonly font?: string;
      readonly size: { readonly pct: number } | { readonly pts: number };
      readonly color?: Rgba;
    }
  | {
      readonly type: 'autonum';
      readonly scheme: string;
      readonly startAt: number;
      readonly font?: string;
      readonly size: { readonly pct: number } | { readonly pts: number };
      readonly color?: Rgba;
    }
  | {
      /** A picture bullet (`a:buBlip`, ECMA-376 Part 1, 21.1.2.4.2). */
      readonly type: 'picture';
      /** The part whose relationships resolve `rId` and `svgRId`. */
      readonly part: string;
      readonly rId: string;
      /** An `asvg:svgBlip` alternative, when present. */
      readonly svgRId?: string;
      /** The picture's width over its height. */
      readonly aspect: number;
      readonly size: { readonly pct: number } | { readonly pts: number };
    };

/** Line spacing: a percentage of single spacing, or an exact size in hundredths of a point. */
export type Spacing = { readonly pct: number } | { readonly pts: number };

/** Fully resolved character properties. Sizes in hundredths of a point. */
export interface RunProps {
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: string;
  readonly strike: string;
  readonly caps: 'none' | 'all' | 'small';
  /** Character spacing, hundredths of a point. */
  readonly spacing: number;
  /** Kerning threshold in hundredths of a point (`kern`); kerning applies at or above it. */
  readonly kern?: number;
  /** Superscript/subscript offset, percent of size (positive raises). */
  readonly baseline: number;
  readonly latin: string;
  /**
   * The theme font `latin` comes from (`+mj-lt` is `major`, `+mn-lt` is `minor`); absent when a
   * level names a literal family. ECMA-376 Part 1, 20.1.4.1.17 `latin`.
   */
  readonly latinTheme?: 'major' | 'minor';
  readonly ea: string;
  readonly cs: string;
  readonly sym?: string;
  readonly fill: ResolvedFill;
  /** The underline's own fill (`a:uFill`); absent when the underline takes the text fill. */
  readonly underlineFill?: ResolvedFill;
  /**
   * The outline of the underline (`a:uLn`, or the text outline for `a:uLnTx`); absent when the
   * underline has no outline or the line has no fill.
   */
  readonly underlineLine?: ResolvedLine;
  readonly highlight?: Rgba;
  readonly outline?: ResolvedLine;
  /** The run's language (`lang`), a BCP 47 tag. */
  readonly lang?: string;
  /** The run's alternate language (`altLang`), a BCP 47 tag. */
  readonly altLang?: string;
  readonly hyperlink?: string;
  /** The run's link action (`a:hlinkClick/@action`), such as a jump to another slide. */
  readonly hyperlinkAction?: string;
}

import type { MathZone } from './math.ts';

/** One inline item of a paragraph. */
export interface ResolvedRun {
  readonly kind: 'text' | 'break' | 'field' | 'math';
  /** The `a:r` / `a:br` / `a:fld` / `a14:m` node. */
  readonly nodeId: string;
  /** The `a:t` node holding the text (text and field runs). */
  readonly textNodeId?: string;
  readonly text: string;
  readonly props: RunProps;
  readonly fieldType?: string;
  /**
   * The equation of a math run (`a14:m`, ECMA-376 Part 1, 22.1). Its `text` is one object
   * character (U+FFFC): the caret steps over the whole equation.
   */
  readonly math?: MathZone;
}

/**
 * Custom East Asian line breaking lists (`p:kinsoku`, ECMA-376 Part 1, 19.2.1.17): the
 * characters that may not start a line (`invalStChars`) and may not end one (`invalEndChars`).
 */
export interface KinsokuRules {
  /** The language the lists apply to (`lang`), a BCP 47 tag. No tag: every language. */
  readonly lang?: string;
  /** Characters that may not start a line (`invalStChars`). */
  readonly notStart: string;
  /** Characters that may not end a line (`invalEndChars`). */
  readonly notEnd: string;
}

/** A resolved paragraph. Lengths in EMU. */
export interface ResolvedParagraph {
  readonly nodeId: string;
  readonly level: number;
  readonly align: 'l' | 'ctr' | 'r' | 'just' | 'dist' | 'justLow' | 'thaiDist';
  readonly marL: number;
  readonly marR: number;
  readonly indent: number;
  readonly lineSpacing: Spacing;
  readonly spaceBefore: Spacing;
  readonly spaceAfter: Spacing;
  readonly bullet: ResolvedBullet | null;
  readonly defaultTab: number;
  readonly tabs: readonly { readonly pos: number; readonly align: string }[];
  readonly rtl: boolean;
  readonly eaLineBreak: boolean;
  readonly latinLineBreak: boolean;
  /** `hangingPunct`: East Asian commas and full stops may hang past the line end. */
  readonly hangingPunct: boolean;
  /**
   * The presentation's custom East Asian line breaking lists (`p:kinsoku`), when they apply to
   * the paragraph's language. They replace the built-in lists.
   */
  readonly kinsoku?: KinsokuRules;
  readonly runs: readonly ResolvedRun[];
  /** Properties of the paragraph mark (`a:endParaRPr`), sizing an empty paragraph. */
  readonly endProps: RunProps;
}

/** Autofit as stored. */
export type Autofit =
  | { readonly kind: 'none' }
  | { readonly kind: 'normal'; readonly fontScale: number; readonly lineSpaceReduction: number }
  | { readonly kind: 'shape' };

/** Resolved `a:bodyPr`. Insets in EMU. */
export interface BodyProps {
  readonly lIns: number;
  readonly tIns: number;
  readonly rIns: number;
  readonly bIns: number;
  readonly anchor: 't' | 'ctr' | 'b' | 'just' | 'dist';
  readonly anchorCtr: boolean;
  readonly wrap: 'square' | 'none';
  readonly vert: string;
  readonly rot: number;
  readonly upright: boolean;
  readonly numCol: number;
  readonly spcCol: number;
  readonly autofit: Autofit;
  readonly spcFirstLastPara: boolean;
  /**
   * The preset text warp (`a:prstTxWarp`, ECMA-376 Part 1, 20.1.9.19), when it draws one:
   * absent for `textNoShape` and for a name that `ST_TextShapeType` does not hold.
   */
  readonly warp?: TextWarp;
}

/** A preset text warp: its `ST_TextShapeType` name and its adjust values (`a:avLst`). */
export interface TextWarp {
  readonly preset: string;
  readonly adjust: ReadonlyMap<string, number>;
}

/** A resolved text body. */
export interface ResolvedTextBody {
  /** The `p:txBody` node, or undefined when the text comes only from an inherited prompt. */
  readonly nodeId?: string;
  readonly body: BodyProps;
  readonly paragraphs: readonly ResolvedParagraph[];
}

interface ResolvedBase {
  readonly nodeId: string;
  readonly partName: string;
  readonly layer: Layer;
  readonly name: string;
  readonly xfrm: Xfrm;
  readonly hidden: boolean;
  readonly placeholder?: { readonly type: string; readonly idx?: string };
  /**
   * The drawing id (`p:cNvPr/@id`) of a PresentationML shape, picture, connector, or frame:
   * the id that connector ends (`a:stCxn`, `a:endCxn`) name.
   */
  readonly drawingId?: string;
}

/** One end of a connector glued to a shape (`CT_Connection`, ECMA-376 Part 1, 20.1.2.2.4). */
export interface ConnectionRef {
  /** The drawing id (`p:cNvPr/@id`) of the shape. */
  readonly id: string;
  /** The index of the connection site in the shape's geometry (`a:cxnLst`). */
  readonly idx: number;
}

/** `p:sp` or `p:cxnSp`. */
export interface ResolvedShape extends ResolvedBase {
  readonly kind: 'shape' | 'connector';
  readonly geometry: GeometrySpec;
  /** A connector's glued ends (`p:cNvCxnSpPr/a:stCxn` and `a:endCxn`). */
  readonly connections?: { readonly start?: ConnectionRef; readonly end?: ConnectionRef };
  readonly fill: ResolvedFill;
  readonly line?: ResolvedLine;
  readonly effects?: ResolvedEffects;
  readonly text?: ResolvedTextBody;
  /** A text box (`p:cNvSpPr/@txBox`, ECMA-376 Part 1, 19.3.1.13): a press in it edits its text. */
  readonly textBox?: true;
  /**
   * The text rectangle of a diagram shape (`dsp:txXfrm`), in the same space as `xfrm`. Its
   * turn adds to the shape's turn. Text lays out in it, not in the geometry's text rectangle.
   */
  readonly textXfrm?: Xfrm;
  /**
   * The data model point (`dgm:pt/@modelId`) whose text a diagram shape shows, when it shows
   * the text of exactly one node. A diagram text edit writes this point and the shape.
   */
  readonly diagramNode?: string;
}

/** `p:pic`. */
export interface ResolvedPicture extends ResolvedBase {
  readonly kind: 'picture';
  readonly geometry: GeometrySpec;
  readonly fill: Extract<ResolvedFill, { type: 'picture' }> | { readonly type: 'none' };
  readonly line?: ResolvedLine;
  readonly effects?: ResolvedEffects;
  /** Video/audio or an OLE icon: render the poster, preserve the media. */
  readonly media?: string;
}

/** `p:grpSp`. Child coordinates are in the `chOff`/`chExt` space. */
export interface ResolvedGroup extends ResolvedBase {
  readonly kind: 'group';
  readonly childOffset: { readonly x: number; readonly y: number };
  readonly childExtent: { readonly cx: number; readonly cy: number };
  readonly children: readonly ResolvedElement[];
  readonly fill: ResolvedFill;
}

/** The borders of one table cell. Absent: no line. */
export interface CellBorders {
  readonly left?: ResolvedLine;
  readonly right?: ResolvedLine;
  readonly top?: ResolvedLine;
  readonly bottom?: ResolvedLine;
  /** Diagonal from the top-left to the bottom-right corner. */
  readonly tlToBr?: ResolvedLine;
  /** Diagonal from the bottom-left to the top-right corner. */
  readonly blToTr?: ResolvedLine;
}

/** One `a:tc`, at its grid position. Covered cells of a merge have `covered` set. */
export interface ResolvedTableCell {
  readonly nodeId: string;
  readonly row: number;
  readonly col: number;
  readonly gridSpan: number;
  readonly rowSpan: number;
  /**
   * Covered by another cell's `gridSpan` or `rowSpan`, or no `a:tc` at the position: not drawn.
   * A merge marker (`hMerge`, `vMerge`) that no span reaches does not cover its cell.
   */
  readonly covered: boolean;
  readonly fill: ResolvedFill;
  readonly borders: CellBorders;
  readonly text: ResolvedTextBody;
}

/** A resolved `a:tbl`. Lengths in EMU. */
export interface ResolvedTable {
  readonly columns: readonly number[];
  /** Minimum row heights (`a:tr/@h`). */
  readonly rows: readonly { readonly nodeId: string; readonly height: number }[];
  /** Row-major; `cells[r * columns.length + c]`. */
  readonly cells: readonly ResolvedTableCell[];
  /** The table background (`a:tblPr` fill, or the style's `tblBg`). */
  readonly background: ResolvedFill;
  readonly rtl: boolean;
  /** The style the table uses, when it was found (`styleId`). */
  readonly styleId?: string;
}

/**
 * `p:graphicFrame`: tables, charts, diagrams, OLE. Tables, OLE posters, and diagrams with a
 * saved drawing render, and charts render through an installed chart renderer; the rest show a
 * fallback.
 */
export interface ResolvedGraphicFrame extends ResolvedBase {
  readonly kind: 'graphicFrame';
  readonly uri: string;
  readonly frameKind: 'chart' | 'diagram' | 'table' | 'ole' | 'unknown';
  /** The table, for `frameKind: 'table'`. */
  readonly table?: ResolvedTable;
  /** @deprecated Never set. An OLE object's picture is {@link ResolvedGraphicFrame.poster}. */
  readonly fallbackRId?: string;
  /**
   * The picture that shows an OLE object (`p:oleObj/p:pic`; ECMA-376 Part 1, 19.3.1.36), placed
   * in the frame's box. It carries the frame's node id. Absent for other frames, and when the
   * picture is not an internal PNG, JPEG, GIF, or BMP part (the frame then shows a placeholder).
   */
  readonly poster?: ResolvedPicture;
  /**
   * The chart part, for `frameKind: 'chart'` (ECMA-376 Part 1, 21.2). Absent when the part is
   * missing or is not a chart. The engine draws a chart only through an installed chart
   * renderer (`installDrawingRenderers`); without one the frame shows a placeholder.
   */
  readonly chart?: ResolvedChartPart;
  /**
   * The saved drawing of a diagram (`dsp:drawing`; ECMA-376 Part 1, 21.4): its shapes and
   * groups, in the frame's local space, drawn by the normal shape renderer. Read-only. Absent
   * when the diagram has no saved drawing, or when the drawing does not resolve within the
   * limits; the frame then shows a placeholder.
   */
  readonly drawing?: readonly ResolvedElement[];
  /**
   * Why a diagram frame shows its placeholder: `no-drawing` (no saved drawing, or a part that
   * does not resolve), `unsupported` (an element outside the supported subset), or `limit` (past
   * the shape, depth, or slide limit). Absent when `drawing` is present and for other frames.
   */
  readonly drawingFallback?: DiagramFallback;
}

/** Why a diagram frame shows its placeholder (`ResolvedGraphicFrame.drawingFallback`). @public */
export type DiagramFallback = 'no-drawing' | 'unsupported' | 'limit';

export type ResolvedElement =
  ResolvedShape | ResolvedPicture | ResolvedGroup | ResolvedGraphicFrame;

/** One slide, fully resolved. */
export interface ResolvedSlide {
  readonly partName: string;
  readonly index: number;
  /** Slide size, EMU. */
  readonly width: number;
  readonly height: number;
  readonly background: ResolvedFill;
  /** The part whose relationships resolve the background picture. */
  readonly backgroundPart: string;
  /** In paint order: master, layout, slide. */
  readonly elements: readonly ResolvedElement[];
}
