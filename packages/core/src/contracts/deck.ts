// Deck-level summaries of the editor contract: custom shows, sections, and speaker notes.
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
