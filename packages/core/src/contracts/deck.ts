// Deck-level summaries of the editor contract: custom shows, sections, speaker notes, and
// slide layouts.
// Declaration-only, DOM-free, re-exported by `neutral.ts`.

/** A custom show of the deck (`p:custShow`, ECMA-376 Part 1, 19.2.1.7), for links. @public */
export interface CustomShowSummary {
  /** The show's `id` in the file. */
  readonly id: number;
  /** The show's name, on one line. */
  readonly name: string;
}

/** One section of the slide list (`p14:section`). @public */
export interface SlideSection {
  /** The section's identity: its GUID in the file. */
  readonly id: string;
  readonly name: string;
  /** The index of its first slide; for an empty section, the index its slides would take. */
  readonly firstSlide: number;
  readonly slideCount: number;
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
