/**
 * Keeps the default editor's floating corner parts above the speaker notes card.
 *
 * The status area (the view switch and the zoom) and the font notice float at the bottom corners
 * of the workspace. In the Normal view the notes card fills the bottom of the workspace, so the
 * parts would cover the last line of the notes. The tracker writes the height that the card takes
 * from the bottom of the workspace into `--pptx-notes-inset`, and the stylesheet adds it to the
 * `bottom` of the parts. The card opens, closes, and resizes, so the tracker follows its size.
 */

/** The custom property that holds the inset, in pixels. @internal */
export const NOTES_INSET_PROPERTY = '--pptx-notes-inset';

/** A running notes inset tracker. @internal */
export interface NotesInsetTracker {
  /** Find the notes card again and measure, for example after a view change. */
  refresh(): void;
  /** Stop observing and remove the property. */
  dispose(): void;
}

/**
 * Track the notes card in `box`, the default editor's views box. The inset is the distance from
 * the bottom of `box` to the top of the card. It is 0 when `box` holds no card (the Slide Sorter
 * view, or no notes pane) or when the card is not rendered. @internal
 */
export function trackNotesInset(box: HTMLElement): NotesInsetTracker {
  const Observer = box.ownerDocument.defaultView?.ResizeObserver;
  let notes: HTMLElement | null = null;
  let last: string | null = null;
  const measure = (): void => {
    let inset = 0;
    if (notes?.isConnected) {
      const card = notes.getBoundingClientRect();
      if (card.width > 0 || card.height > 0) {
        const frame = box.getBoundingClientRect();
        // The rectangles are in viewport pixels. In a scaled container (`transform: scale`),
        // `bottom` counts in the box's own pixels, so divide by the scale.
        const scale =
          frame.height > 0 && box.offsetHeight > 0 ? frame.height / box.offsetHeight : 1;
        inset = Math.max(0, Math.round((frame.bottom - card.top) / scale));
      }
    }
    const value = `${inset}px`;
    if (value === last) return;
    last = value;
    box.style.setProperty(NOTES_INSET_PROPERTY, value);
  };
  const observer = Observer ? new Observer(measure) : null;
  observer?.observe(box);
  const refresh = (): void => {
    const next = box.querySelector<HTMLElement>('.pptx-notes');
    if (next !== notes) {
      if (notes) observer?.unobserve(notes);
      notes = next;
      if (notes) observer?.observe(notes);
    }
    measure();
  };
  refresh();
  return {
    refresh,
    dispose() {
      observer?.disconnect();
      notes = null;
      box.style.removeProperty(NOTES_INSET_PROPERTY);
    },
  };
}
