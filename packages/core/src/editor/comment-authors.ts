// The avatar text of a comment author. The canvas markers (`comment-markers.ts`) and the Pro
// comments pane read the same answer. The author's color comes from the review roster
// (`review-authors.ts`).

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
