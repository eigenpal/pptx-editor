// The automation host protocol, adapted from DOCX Editor's
// `packages/core/src/automation/protocol.ts`. DOCX Editor commit adf0e178.
//
// One interface, two implementations: a headless host that owns a package it opened from bytes,
// and a browser host that borrows a live editor's store. Both answer the same operations with
// the same results, because the operations are implemented once (`plan.ts`) over the canonical
// package, and neither host may write any other way.
//
// EVERYTHING HERE IS TRANSPORT-SHAPED. A request is data, a response is data, and a handle is a
// name, not a pointer, so the same host can sit behind a worker port, an HTTP boundary, or
// nothing at all. Nothing in this file names a tree node or a store: a value a consumer receives
// must not be a reference into the engine.

import type { LineEndType, ShapeFormatProperties } from '../contracts/neutral.ts';
import type { AutomationOperation } from './operations.ts';

/** What kind of presentation object a handle names. */
export type AutomationObjectKind =
  'presentation' | 'slide' | 'slideLayout' | 'slideMaster' | 'shape';

declare const AUTOMATION_HANDLE_BRAND: unique symbol;

/**
 * An opaque host-minted name for a presentation object.
 *
 * Branded so a consumer cannot invent one. Its content is meaningless outside the host that
 * minted it: each host draws a random token when it starts and stamps it into every ref, so a
 * ref passed to another host, even one open on the same bytes, is `invalid-handle`.
 */
export type AutomationHandleRef = string & { readonly [AUTOMATION_HANDLE_BRAND]: 'handle' };

/** A stable reference to one presentation object, valid for the life of its host. */
export interface AutomationHandle<K extends AutomationObjectKind = AutomationObjectKind> {
  readonly kind: K;
  readonly ref: AutomationHandleRef;
}

/**
 * The object an earlier operation of the same batch answered with, by its index.
 *
 * How one batch adds a slide and then writes to it: the slide has no handle until the add runs,
 * so the write names the add's result instead. A reference to an operation that answered null
 * (`getSlide` with `orNull`) makes a read answer `null` and a write fail with `invalid-handle`.
 */
export interface AutomationResultRef {
  readonly result: number;
}

/** A handle, or the handle an earlier operation of the batch answered with. */
export type AutomationRef = AutomationHandle | AutomationResultRef;

/**
 * What a host supports, fixed at construction and frozen.
 *
 * A headless host reports `selection` false: it has no reader to select for, and refuses the
 * selection operations rather than inventing a selection.
 */
export interface AutomationCapabilities {
  /** Reading and editing the presentation. Every operation but the selection ones needs it. */
  readonly document: boolean;
  /** Serializing the current presentation back to `.pptx` bytes. */
  readonly save: boolean;
  /** Change notification through {@link AutomationHost.subscribe}. */
  readonly events: boolean;
  /** A reader's slide, shape, or text selection exists and can be read and set. */
  readonly selection: boolean;
  /** Text writes re-fit text bodies by their autofit rules (a text measurer is present). */
  readonly autofit: boolean;
  /**
   * Pictures of slides and shapes through {@link AutomationHost.renderImage}. Drawing needs a
   * DOM and a canvas, so only a browser host reports it.
   */
  readonly images: boolean;
}

/**
 * Why an operation was refused, as a value to branch on.
 *
 * Distinguishable on purpose: the object model maps each code to a different error, and a
 * caller that retries `stale-revision` must not retry `invalid-handle`.
 */
export type AutomationErrorCode =
  /** Nothing was applied: `expectedRevision` did not match the host's revision. */
  | 'stale-revision'
  /** A handle this host never minted, one of another kind, or one whose object is gone. */
  | 'invalid-handle'
  /** An argument outside its range: an index, an offset, a length, a size. */
  | 'invalid-argument'
  /** The operation needs a capability this host reports false. */
  | 'unsupported-capability'
  /** The host has been disposed. Every later call fails this way. */
  | 'disposed'
  /** The store refused the transaction; nothing was applied. */
  | 'transaction-refused'
  /** The operation is not one this protocol version defines. */
  | 'unknown-operation'
  /** The target holds content this operation does not edit (a chart, a diagram, a picture). */
  | 'unsupported-content'
  /** A lookup by key or index found nothing, and the caller did not ask for null. */
  | 'item-not-found'
  /** The host is live but holds no presentation right now (an editor between mounts). */
  | 'document-unavailable';

/**
 * One refused operation. Plain data, not an `Error`: it crosses transports.
 *
 * `code` is the stable part. `detail` is for logs and for the object model's diagnostic code;
 * the object model never copies it into a message, because a store rejection reason would
 * become documented API the moment somebody matched on it.
 */
export interface AutomationError {
  readonly code: AutomationErrorCode;
  /** Human-readable, for a log. Never parsed. */
  readonly message: string;
  /** Machine-ish specifics: the store's rejection reason, the bad argument's name. */
  readonly detail?: string;
}

/** `ST_PlaceholderType` read as the object model names it. */
export interface AutomationPlaceholderRead {
  /** The `p:ph/@type` value (`obj` when absent, ECMA-376 Part 1, 19.7.10). */
  readonly type: string;
  /** What the shape holds now: its own kind, as `AutomationShapeRead.type` spells it. */
  readonly containedType: string | null;
}

/** One shape, as plain values. Lengths in points, the rotation in degrees. */
export interface AutomationShapeRead {
  /** `p:cNvPr/@id` as a string. */
  readonly id: string;
  readonly name: string;
  /**
   * `geometricShape`, `line`, `textBox`, `placeholder`, `image`, `group`, `table`, `chart`,
   * `diagram` (a `dgm:` frame), `media`, `ole`, `freeform`, or `unsupported`.
   */
  readonly type: string;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly visible: boolean;
  /** Position in the owner's z-order, from 0 at the back. */
  readonly zOrderPosition: number;
  /** Group depth: 0 for a top-level shape. */
  readonly level: number;
  /** The preset geometry (`a:prstGeom/@prst`), or null. */
  readonly geometry: string | null;
  readonly altTextTitle: string;
  readonly altTextDescription: string;
  /** The decorative flag of the `a16:decorative` extension. */
  readonly isDecorative: boolean;
  /** Whether the shape can hold text (`p:txBody`). */
  readonly hasTextFrame: boolean;
  readonly placeholder: AutomationPlaceholderRead | null;
  readonly parentGroup: AutomationHandle<'shape'> | null;
}

/** A shape fill. Colors are `#RRGGBB`; transparency is 0 to 1. */
export interface AutomationFillRead {
  /** `noFill`, `solid`, `gradient`, `pattern`, `pictureAndTexture`, or `slideBackground`. */
  readonly type: string;
  readonly foregroundColor: string | null;
  readonly transparency: number | null;
}

/**
 * The fill of a background (`p:bg`, ECMA-376 Part 1, 19.3.1.1): the own one of the part, or the
 * one it shows from the layout or master. Colors are `#RRGGBB`; transparency is 0 to 1.
 */
export interface AutomationBackgroundFillRead {
  /** `solid`, `gradient`, `pattern`, `picture`, or `unsupported` (`a:noFill`, `a:grpFill`). */
  readonly type: 'solid' | 'gradient' | 'pattern' | 'picture' | 'unsupported';
  /** The solid color. */
  readonly color: string | null;
  /** The transparency of a solid color (`a:alpha`) or a picture (`a:alphaModFix`). */
  readonly transparency: number | null;
  /**
   * The gradient shading: `linear` (`a:lin`), `circle`, `rect`, or `shape` (`a:path/@path`), or
   * `shadeToTitle` (`p:bgPr/@shadeToTitle`).
   */
  readonly gradientType: string | null;
  /** The pattern preset (`ST_PresetPatternVal`). */
  readonly pattern: string | null;
  /** The pattern colors (`a:fgClr`, `a:bgClr`). */
  readonly foregroundColor: string | null;
  readonly backgroundColor: string | null;
}

/** The background of a slide, a layout, or a master. */
export interface AutomationBackgroundRead {
  /** The part has no own `p:bg`, so it shows the background of its layout or master. */
  readonly followsMaster: boolean;
  /** `showMasterSp="0"`: the shapes of the level above do not show. False for a master. */
  readonly graphicsHidden: boolean;
  readonly fill: AutomationBackgroundFillRead;
}

/** A shape outline. The weight is in points. */
export interface AutomationLineRead {
  readonly head?: LineEndType;
  readonly tail?: LineEndType;
  readonly visible: boolean;
  readonly color: string | null;
  readonly weight: number | null;
  /** The `a:prstDash/@val` value, or null. */
  readonly dashStyle: string | null;
  /** The compound line (`a:ln/@cmpd`). */
  readonly style: string | null;
  readonly transparency: number | null;
}

/** A text body's frame. Margins in points. */
export interface AutomationTextFrameRead {
  /** `none`, `normal` (`a:normAutofit`), or `shape` (`a:spAutoFit`). */
  readonly autoSize: 'none' | 'normal' | 'shape';
  readonly leftMargin: number;
  readonly rightMargin: number;
  readonly topMargin: number;
  readonly bottomMargin: number;
  /** `a:bodyPr/@anchor`, with `anchorCtr` (`ctr` + centered reads as `ctrCentered`). */
  readonly verticalAlignment: string;
  readonly wordWrap: boolean;
  readonly hasText: boolean;
}

/** The effective font of a text range. A field the runs disagree on is null. */
export interface AutomationFontRead {
  readonly name: string | null;
  readonly size: number | null;
  readonly bold: boolean | null;
  readonly italic: boolean | null;
  /** `a:rPr/@u`, `none` without an underline. */
  readonly underline: string | null;
  readonly color: string | null;
  readonly strikethrough: boolean | null;
  readonly doubleStrikethrough: boolean | null;
  readonly allCaps: boolean | null;
  readonly smallCaps: boolean | null;
  readonly superscript: boolean | null;
  readonly subscript: boolean | null;
}

/** One run of a text body with its effective character format. */
export interface AutomationTextRunRead {
  /** The run text; a paragraph break reads as `\r` and a line break as `\v`. */
  readonly text: string;
  readonly font: AutomationFontRead;
}

/** The effective paragraph format of a text range. A field the paragraphs disagree on is null. */
export interface AutomationParagraphRead {
  /** `a:pPr/@algn`: `l`, `ctr`, `r`, `just`, `justLow`, `dist`, or `thaiDist`. */
  readonly alignment: string | null;
  readonly bulletVisible: boolean | null;
  /** `none`, `char`, or `autonum`. */
  readonly bulletType: string | null;
  /** The autonumber scheme (`ST_TextAutonumberScheme`) when numbered. */
  readonly bulletStyle: string | null;
  /** 1-based list level (`a:pPr/@lvl` + 1). */
  readonly indentLevel: number | null;
}

/** One key and value of a tag list (`p:tagLst`). */
export interface AutomationTag {
  readonly key: string;
  readonly value: string;
}

/** The core and extended document properties. */
export interface AutomationPropertiesRead {
  readonly title: string;
  readonly subject: string;
  readonly author: string;
  readonly keywords: string;
  readonly comments: string;
  readonly category: string;
  readonly lastAuthor: string;
  readonly revisionNumber: number;
  readonly creationDate: string | null;
  readonly company: string;
  readonly manager: string;
}

/**
 * One custom document property (`docProps/custom.xml`, ECMA-376 Part 1, 22.3). A Date is an ISO
 * 8601 text in UTC.
 */
export interface AutomationCustomPropertyRead {
  readonly key: string;
  readonly type: 'String' | 'Number' | 'Boolean' | 'Date';
  readonly value: string | number | boolean;
}

/** One custom XML part (ECMA-376 Part 1, 22.5): its item id and the namespace of its root. */
export interface AutomationCustomXmlPartRead {
  /** `ds:itemID`, braced and upper case. */
  readonly id: string;
  /** The namespace URI of the root element, or an empty string. */
  readonly namespaceUri: string;
}

/** One cell of a table, at its grid position. */
export interface AutomationTableCellRead {
  readonly text: string;
  readonly rowIndex: number;
  readonly columnIndex: number;
  /** The rows and columns the cell spans. */
  readonly rowCount: number;
  readonly columnCount: number;
  /** `ST_TextAnchoringType` of the cell (`a:tcPr/@anchor`, default `t`). */
  readonly anchor: string;
  /** `a:tcPr/@anchorCtr`. */
  readonly anchorCenter: boolean;
}

/** A table's grid and values. Widths and heights in points. */
export interface AutomationTableRead {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly values: readonly (readonly string[])[];
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
  /** Merged areas, top-left cell first. */
  readonly mergedAreas: readonly {
    readonly rowIndex: number;
    readonly columnIndex: number;
    readonly rowCount: number;
    readonly columnCount: number;
  }[];
  /** `a:tblPr` flags and the style id. */
  readonly style: {
    readonly styleId: string | null;
    /** The `styleName` of the style the table uses, or null without a style or a name. */
    readonly styleName: string | null;
    readonly firstRow: boolean;
    readonly lastRow: boolean;
    readonly firstColumn: boolean;
    readonly lastColumn: boolean;
    readonly bandedRows: boolean;
    readonly bandedColumns: boolean;
  };
}

/**
 * The borders and margins of one table cell (`a:tcPr`, ECMA-376 Part 1, 21.1.3.17). A border
 * reads as the cell shows it: its own line, else the table style's on a slide. Margins in points.
 */
export interface AutomationTableCellFormatRead {
  readonly borders: {
    readonly left: AutomationLineRead;
    readonly right: AutomationLineRead;
    readonly top: AutomationLineRead;
    readonly bottom: AutomationLineRead;
    /** `a:lnTlToBr`: from the top-left to the bottom-right corner. */
    readonly diagonalDown: AutomationLineRead;
    /** `a:lnBlToTr`: from the bottom-left to the top-right corner. */
    readonly diagonalUp: AutomationLineRead;
  };
  readonly margins: {
    readonly left: number;
    readonly right: number;
    readonly top: number;
    readonly bottom: number;
  };
}

/**
 * The adjust values of a preset geometry (`a:prstGeom/a:avLst`, ECMA-376 Part 1, 20.1.9.5), in
 * the order the preset defines them. A length is a fraction of the shape size (the guide value
 * / 100 000); an angle is in degrees (the guide value / 60 000).
 */
export interface AutomationAdjustmentsRead {
  /** `a:prstGeom/@prst`. */
  readonly preset: string;
  readonly values: readonly number[];
}

/** One hyperlink of a slide. */
export interface AutomationHyperlinkRead {
  readonly address: string | null;
  readonly screenTip: string | null;
  /** `textRange` or `shape`: what carries the link. */
  readonly type: 'textRange' | 'shape';
  readonly shape: AutomationHandle<'shape'>;
  /** For a text link: the range it covers in the shape's text. */
  readonly start: number | null;
  readonly length: number | null;
}

/** One binding of the host: an id that names a shape. Bindings are never file content. */
export interface AutomationBindingRead {
  readonly id: string;
  /** What the binding names. Only `shape`. */
  readonly type: 'shape';
}

/** A layout or master. */
export interface AutomationLayoutRead {
  /** `p:sldLayoutId/@id` or `p:sldMasterId/@id`, as a string. */
  readonly id: string;
  readonly name: string;
  /** For a layout: `p:sldLayout/@type` (`cust` when absent). Null for a master. */
  readonly type: string | null;
  readonly master: AutomationHandle<'slideMaster'> | null;
}

/** A slide. */
export interface AutomationSlideRead {
  /** `p:sldId/@id`, as a string. */
  readonly id: string;
  readonly index: number;
  readonly hidden: boolean;
  readonly layout: AutomationHandle<'slideLayout'> | null;
  readonly master: AutomationHandle<'slideMaster'> | null;
}

/** The presentation. Lengths in points. */
export interface AutomationPresentationRead {
  /** A stable id for the open presentation in this host. */
  readonly id: string;
  /** The core properties title. */
  readonly title: string;
  readonly slideWidth: number;
  readonly slideHeight: number;
}

/** A reader's selection (browser host). */
export interface AutomationSelectionRead {
  readonly slides: readonly AutomationHandle<'slide'>[];
  readonly shapes: readonly AutomationHandle<'shape'>[];
  /** The selected text: a shape and a range of its text, or null. */
  readonly text: {
    readonly shape: AutomationHandle<'shape'>;
    readonly start: number;
    readonly length: number;
  } | null;
}

/** What an operation answered with. */
export type AutomationValue =
  | { readonly kind: 'shapeFormat'; readonly format: ShapeFormatProperties }
  | { readonly kind: 'handle'; readonly handle: AutomationHandle }
  | { readonly kind: 'handles'; readonly handles: readonly AutomationHandle[] }
  /** A lookup that found nothing, asked for with `orNull`. */
  | { readonly kind: 'null' }
  | { readonly kind: 'text'; readonly text: string }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'flag'; readonly value: boolean }
  /** `.pptx` bytes as base64: an exported slide. */
  | { readonly kind: 'base64'; readonly base64: string }
  | { readonly kind: 'presentation'; readonly presentation: AutomationPresentationRead }
  | { readonly kind: 'slide'; readonly slide: AutomationSlideRead }
  | { readonly kind: 'layout'; readonly layout: AutomationLayoutRead }
  | { readonly kind: 'shape'; readonly shape: AutomationShapeRead }
  | { readonly kind: 'fill'; readonly fill: AutomationFillRead }
  | { readonly kind: 'background'; readonly background: AutomationBackgroundRead }
  | { readonly kind: 'line'; readonly line: AutomationLineRead }
  | { readonly kind: 'textFrame'; readonly textFrame: AutomationTextFrameRead }
  | { readonly kind: 'font'; readonly font: AutomationFontRead }
  | { readonly kind: 'textRuns'; readonly runs: readonly AutomationTextRunRead[] }
  | { readonly kind: 'paragraph'; readonly paragraph: AutomationParagraphRead }
  | { readonly kind: 'tags'; readonly tags: readonly AutomationTag[] }
  | { readonly kind: 'properties'; readonly properties: AutomationPropertiesRead }
  | {
      readonly kind: 'customProperties';
      readonly properties: readonly AutomationCustomPropertyRead[];
    }
  | {
      readonly kind: 'customXmlParts';
      readonly parts: readonly AutomationCustomXmlPartRead[];
    }
  | { readonly kind: 'table'; readonly table: AutomationTableRead }
  | { readonly kind: 'tableCell'; readonly cell: AutomationTableCellRead }
  | { readonly kind: 'tableCellFormat'; readonly format: AutomationTableCellFormatRead }
  | { readonly kind: 'adjustments'; readonly adjustments: AutomationAdjustmentsRead }
  | { readonly kind: 'hyperlinks'; readonly hyperlinks: readonly AutomationHyperlinkRead[] }
  | { readonly kind: 'bindings'; readonly bindings: readonly AutomationBindingRead[] }
  | { readonly kind: 'selection'; readonly selection: AutomationSelectionRead }
  /** A command that committed. The effect is the response's revision. */
  | { readonly kind: 'applied' };

/**
 * One operation's outcome.
 *
 * `skipped` is what makes an atomic batch honest. When a batch fails, every operation other
 * than the failing one reports `skipped`, including reads and the operations before it:
 * reporting those as `ok` would describe a state that was never published.
 */
export type AutomationOperationResult =
  | { readonly status: 'ok'; readonly value: AutomationValue }
  | { readonly status: 'error'; readonly error: AutomationError }
  | { readonly status: 'skipped' };

/** One ordered batch, and the revision it was planned against. */
export interface AutomationBatchRequest {
  readonly operations: readonly AutomationOperation[];
  /** Refuse the whole batch unless the host is at this revision. */
  readonly expectedRevision?: number;
  /**
   * Batches that carry the same group, one after the other, are one undo step in a host that
   * keeps an undo history. A client sends one group for the batches of one sync that it splits.
   * A host without a history ignores it.
   */
  readonly historyGroup?: string;
}

/** A verdict, one result per operation in request order, and the revision afterwards. */
export interface AutomationBatchResponse {
  /** True only when every operation succeeded and the commands committed. */
  readonly ok: boolean;
  readonly results: readonly AutomationOperationResult[];
  /** The host's revision after the batch. Unchanged when nothing committed. */
  readonly revision: number;
  /** Whether the batch moved the presentation. */
  readonly changed: boolean;
}

/** `.pptx` bytes, or why they could not be produced. */
export type AutomationSaveResult =
  | { readonly ok: true; readonly bytes: Uint8Array }
  | { readonly ok: false; readonly error: AutomationError };

/** What {@link AutomationHost.renderImage} draws. */
export interface AutomationImageRequest {
  /** A slide, or a shape on a slide, as a `prepareImage` operation answered it. */
  readonly target: AutomationHandle;
  /** The width of the box the picture fits into, in pixels, keeping its aspect ratio. */
  readonly width?: number;
  /** The height of that box. Absent sizes: the picture has its own size. */
  readonly height?: number;
}

/** A PNG picture as base64, or why it could not be drawn. */
export type AutomationImageResult =
  | { readonly ok: true; readonly base64: string }
  | { readonly ok: false; readonly error: AutomationError };

/** The presentation moved. Coarse on purpose: a consumer reads again what it needs. */
export interface AutomationChangeEvent {
  readonly revision: number;
}

/** What a subscription returns. Calling it twice is safe. */
export type AutomationUnsubscribe = () => void;

/**
 * The protocol both hosts implement.
 *
 * A batch is one revision, one undo step, and one transaction. Operations run in order, so a
 * read sees the writes before it in the same batch. If any operation is refused, nothing is
 * written.
 */
export interface AutomationHost {
  readonly capabilities: AutomationCapabilities;
  /** Monotonic revision of the presentation this host acts on. */
  revision(): number;
  /** Run one ordered batch. */
  execute(request: AutomationBatchRequest): AutomationBatchResponse;
  /** The current presentation as `.pptx` bytes, through the normalizing writer. */
  save(): AutomationSaveResult;
  /**
   * A PNG picture of a slide or a shape, drawn from the current presentation. A host without
   * the `images` capability answers `unsupported-capability`.
   */
  renderImage?(request: AutomationImageRequest): Promise<AutomationImageResult>;
  /** Change notification. Returns an unsubscribe that is safe to call more than once. */
  subscribe(listener: (event: AutomationChangeEvent) => void): AutomationUnsubscribe;
  /** Release everything this host holds. Idempotent; every later call fails `disposed`. */
  dispose(): void;
}
