// The find contract: how find compares text, and what a match reports.

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
