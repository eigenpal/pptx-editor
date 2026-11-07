// How the comment chrome draws one author: the color slot and the avatar text. The canvas
// markers (`comment-markers.ts`) and the Pro comments pane read the same answers, so a marker and
// its card cannot disagree about who draws in what color.

/** How many author slots the `--doc-review-author-N` token ramp defines. @public */
export const COMMENT_AUTHOR_SLOTS = 8;

/**
 * The color slot of a comment author, 0 to 7: the token `--doc-review-author-<slot>` colors the
 * author's avatar and the leading edge of the author's cards. The slot comes from a hash of the
 * name, so an author keeps one color on every slide, in every session, and for every
 * collaborator. An empty name takes slot 0. @public
 */
export function commentAuthorSlot(author: string): number {
  // FNV-1a over the UTF-16 code units: small, stable, and spread well over short names.
  let hash = 0x811c9dc5;
  for (let i = 0; i < author.length; i += 1) {
    hash ^= author.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return author ? hash % COMMENT_AUTHOR_SLOTS : 0;
}

/**
 * The avatar text of a comment: its own initials (up to three characters), else the first letters
 * of the first two words of the author's name, else "?". @public
 */
export function commentInitials(comment: {
  readonly initials: string;
  readonly author: string;
}): string {
  const own = comment.initials.trim();
  if (own) return [...own].slice(0, 3).join('');
  const words = comment.author.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return '?';
  return words
    .slice(0, 2)
    .map((word) => String.fromCodePoint(word.codePointAt(0)!).toUpperCase())
    .join('');
}
