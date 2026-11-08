// Which color each review author draws in. Adapted from DOCX Editor's
// packages/core/src/output/revision-presentation.ts (`REVIEW_AUTHOR_SLOTS`, `ReviewAuthorInfo`,
// `reviewAuthorSlotColor`, `reviewAuthorsOf`, `reviewAuthorSlotsOf`, `StableReviewAuthorSlots`,
// `createStableReviewAuthorSlots`) and packages/core/src/editor/surface-review-authors.ts at
// f189d6503.
//
// A comment carries an author name and no color. The color is presentation policy: each author
// takes a slot of the `--doc-review-author-N` token ramp by order of first appearance in the
// deck (slide order, each thread before its replies), and keeps it for the editor's session. A
// hash of the name would be stable too, but two names collide in an eight-slot hash often
// enough to draw two reviewers in one color, which tells the reader the wrong thing. The canvas
// markers, the text bands, the shape rings, and the Pro comments pane all read this one roster,
// so a card and the content it annotates cannot disagree about who draws in what.

/** How many author slots the `--doc-review-author-N` token ramp defines. @public */
export const REVIEW_AUTHOR_SLOTS = 8;

/**
 * One review author, resolved: who, which ramp slot, and what they draw in. @public
 */
export interface ReviewAuthorInfo {
  /** The author's display name, exactly as the file gives it. */
  readonly author: string;
  /**
   * Stable session rank, by order of first appearance: an unbounded index. The ramp defines
   * {@link REVIEW_AUTHOR_SLOTS} colors and every DOM hook wraps into it, so the ninth author
   * ranks 8 here and carries `data-review-author-slot="0"`. A selector built from this value
   * needs the same wrap: `slot % 8`.
   */
  readonly slot: number;
  /** The color this author draws in: `var(--doc-review-author-<slot % 8>)`. */
  readonly color: string;
}

/** The slot's color, as the token reference that every surface resolves through CSS. @public */
export function reviewAuthorSlotColor(slot: number): string {
  return `var(--doc-review-author-${slot % REVIEW_AUTHOR_SLOTS})`;
}

/** The roster: every author of `slots`, in slot order, with the color they draw in. */
export function reviewAuthorsOf(slots: ReadonlyMap<string, number>): readonly ReviewAuthorInfo[] {
  return [...slots].map(([author, slot]) => ({
    author,
    slot,
    color: reviewAuthorSlotColor(slot),
  }));
}

/**
 * `documentSlots` and then every other author of `reviewAuthors`, appended in order. Returns
 * `documentSlots` itself when there is nothing to add, so a caller keying a cache on the map's
 * identity sees no change.
 */
export function reviewAuthorSlotsOf(
  documentSlots: ReadonlyMap<string, number>,
  reviewAuthors: Iterable<string>
): ReadonlyMap<string, number> {
  let combined: Map<string, number> | null = null;
  for (const author of reviewAuthors) {
    if (author === '' || documentSlots.has(author)) continue;
    if (combined === null) combined = new Map(documentSlots);
    if (!combined.has(author)) combined.set(author, combined.size);
  }
  return combined ?? documentSlots;
}

/**
 * One open deck's author slot allocator. The first resolve seeds the slots in deck order. Later
 * resolves keep every slot already issued, also for an author who is gone for now, and new
 * authors take the slots after them. So deleting a comment and undoing the delete cannot
 * recolor another reviewer.
 */
export interface StableReviewAuthorSlots {
  resolve(
    documentSlots: ReadonlyMap<string, number>,
    reviewAuthors: Iterable<string>
  ): ReadonlyMap<string, number>;
}

export function createStableReviewAuthorSlots(): StableReviewAuthorSlots {
  const reserved = new Map<string, number>();
  let nextSlot = 0;
  return {
    resolve(documentSlots, reviewAuthors) {
      const current = reviewAuthorSlotsOf(documentSlots, reviewAuthors);
      for (const author of current.keys()) {
        if (reserved.has(author)) continue;
        reserved.set(author, nextSlot);
        nextSlot += 1;
      }
      return new Map(
        [...current.keys()]
          .map((author) => [author, reserved.get(author)!] as const)
          .sort((left, right) => left[1] - right[1])
      );
    },
  };
}

/**
 * The author hooks of an element that annotates `author`'s comment: `data-review-author`,
 * `data-review-author-slot`, and the `--doc-review-author-current` property. Nothing for an
 * anonymous comment, so it never matches a `[data-review-author]` selector.
 */
export function reviewAuthorHooks(
  author: string,
  info: ReviewAuthorInfo | undefined
): { readonly author: string; readonly slot: number; readonly color: string } | null {
  if (!author) return null;
  const slot = (info?.slot ?? 0) % REVIEW_AUTHOR_SLOTS;
  return { author, slot, color: info?.color ?? reviewAuthorSlotColor(slot) };
}
