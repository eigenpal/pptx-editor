// The slide-number, footer, and date contract: the slides those commands act on, and the
// date that the snapshot reports.

/**
 * The slides that `setSlideNumbers`, `setFooter`, and `setDateTime` act on. @public
 */
export interface HeaderFooterTargets {
  /** The slide indexes to change. Default: all slides. */
  readonly indices?: readonly number[];
  /**
   * With `show: true`, slides whose layout is a title slide layout (`p:sldLayout type="title"`)
   * hide the placeholder instead. Default false.
   */
  readonly hideOnTitleSlides?: boolean;
}

/** The date that the slides show (`EditorSnapshot.dateTime`). @public */
export interface DateTimeState {
  /** The text that the date shows. */
  readonly text: string;
  /** True for fixed text; false for a date field. */
  readonly fixed: boolean;
}
