// The comments contract: the comment threads of a slide as the snapshot reports them, and the
// commands that add, reply to, edit, delete, resolve, and reopen them. Two file formats hold
// comments: the comment lists of ECMA-376 Part 1, 19.4 (`p:cmLst`, `legacy`), and the threaded
// comments of the `p188` extension (`modern`).

/** The file format of a comment thread. @public */
export type CommentFormat = 'modern' | 'legacy';

/**
 * Where a thread points on its slide. `shape`: an element of the slide, by its node id. `point`:
 * a place on the slide, in EMU. `slide`: the slide as a whole, with no place. @public
 */
export type CommentAnchor =
  | { readonly kind: 'slide' }
  | { readonly kind: 'point'; readonly x: number; readonly y: number }
  | { readonly kind: 'shape'; readonly nodeId: string };

/** One comment or reply. Its text and author come from the file: show them as text. @public */
export interface SlideComment {
  /**
   * The id that the comment commands take. When the slide repeats the comment's GUID, or a legacy
   * comment's author and index, a delete of another comment can change it, and a twin (see
   * `identity`) can take the id of a deleted one. Send `expect` with a command, so the editor
   * refuses it when the comment at the id is not the one you showed.
   */
  readonly id: string;
  /**
   * A token that names the comment in this editor's snapshots. Key UI state, such as an open
   * field, by it; pass `id` to the commands. Once shown, an identity never names another
   * comment, also after the slide changes and comes back, or the editor opens another deck.
   *
   * - A comment that never had a twin keeps its identity until it is deleted. A twin shares the
   *   GUID, author, and creation time of another comment of the slide.
   * - Twins keep their identities when one change explains their group: an edit or a resolve
   *   that changes at most one text, or one twin more or less at one place. Otherwise every
   *   twin of the group takes a new identity, so state keyed by the old ones ends.
   * - A comment that a slide shows after its first read takes a new identity.
   *
   * `readCommentThreads` reads the identities from the file alone, without this carry.
   */
  readonly identity: string;
  /** The author's display name. Empty when the file names no author for the comment. */
  readonly author: string;
  /** The author's initials. Empty when the file has none. */
  readonly initials: string;
  /** The date as the file stores it (ISO 8601), or absent. */
  readonly date?: string;
  /** The text, with one line feed between paragraphs. */
  readonly text: string;
  /**
   * Whether `editComment` can replace the text of this comment in a deck open for editing.
   * False for a body that holds more than plain text, such as a mention.
   */
  readonly textEditable: boolean;
}

/** A comment and its replies (`EditorSnapshot.comments`). @public */
export interface CommentThread extends SlideComment {
  readonly format: CommentFormat;
  /** Whether the thread is resolved. Always false for the `legacy` format, which has no state. */
  readonly resolved: boolean;
  readonly anchor: CommentAnchor;
  /** The replies, oldest first. */
  readonly replies: readonly SlideComment[];
}

/**
 * Who writes a comment or a reply. Absent, the editor's `author` option (`setAuthor`) applies,
 * then the collaboration identity. Without a name, the command is refused
 * (`disabled.noCommentAuthor`). @public
 */
export interface CommentAuthorInput {
  /** The display name, 1 to 255 characters. */
  readonly author?: string;
  /** Initials, up to 8 characters. Default: the first letters of the first two words. */
  readonly initials?: string;
  /** ISO 8601 date and time. Default: now, in UTC. */
  readonly date?: string;
}

/**
 * What a host showed of a comment, for a command that acts on it: pass the comment's `author`,
 * `date`, and `text` as the snapshot gave them. The editor refuses the command
 * (`errors.commentChanged`) when the comment at `commentId` has another author or date.
 *
 * A twin shares its GUID, author, and creation time with another comment of the slide, or, in
 * the legacy format, its author and index. A twin's id holds its place among its twins
 * (`~<n>`, or `-<n>` for a legacy comment). For an id with a place, the editor also compares
 * the text, and it refuses the command when the comment has no twin left: the id may have named
 * the twin that is gone. So a collaborator's edit of a comment without a twin never refuses
 * yours. @public
 */
export interface CommentExpect {
  /** The author's display name, as `SlideComment.author` gave it. */
  readonly author: string;
  /** The date, as `SlideComment.date` gave it; absent when the comment has none. */
  readonly date?: string;
  /** The text, as `SlideComment.text` gave it. */
  readonly text: string;
}

/**
 * The comment commands. Text is 1 to 10,000 characters; a line feed starts a new paragraph.
 * The commands on an existing comment take `expect`, what the host showed of that comment.
 * A new comment uses the threaded format, unless the deck has only legacy comments. @public
 */
export type CommentCommand =
  /**
   * Add a thread on the current slide. With `nodeId`, it points at that element; without it, at
   * the selected element when one element is selected, else at the slide.
   */
  | ({
      readonly type: 'addComment';
      readonly text: string;
      readonly nodeId?: string;
    } & CommentAuthorInput)
  /** Reply to the thread of `commentId` (a thread or one of its replies). */
  | ({
      readonly type: 'replyToComment';
      readonly commentId: string;
      readonly text: string;
      readonly expect?: CommentExpect;
    } & CommentAuthorInput)
  /** Replace the text of a comment or a reply. */
  | {
      readonly type: 'editComment';
      readonly commentId: string;
      readonly text: string;
      readonly expect?: CommentExpect;
    }
  /** Delete a reply, or a whole thread when `commentId` is a thread. */
  | { readonly type: 'deleteComment'; readonly commentId: string; readonly expect?: CommentExpect }
  /** Mark a thread resolved. Refused for the legacy format (`disabled.commentNoStatus`). */
  | { readonly type: 'resolveComment'; readonly commentId: string; readonly expect?: CommentExpect }
  /** Mark a resolved thread active again. */
  | {
      readonly type: 'reopenComment';
      readonly commentId: string;
      readonly expect?: CommentExpect;
    };

/**
 * How a comment marker on the slide looks. `icon`: the comment icon. `initials`: a badge with
 * the thread author's initials in the author's color, the reply count, and a check mark when the
 * thread is resolved. @public
 */
export type CommentMarkerStyle = 'initials' | 'icon';

/**
 * What a press on a comment marker does besides picking the thread. `auto`: it also asks for the
 * **Comments** tab of the side panel (`requestPane`). `manual`: it only picks the thread, and the
 * host opens its own comments view. @public
 */
export type ReviewPaneOpening = 'auto' | 'manual';

/** The comment view settings of a review module. Every field is optional. @public */
export interface ReviewPaneOptions {
  /** Default `'auto'`. */
  readonly opening?: ReviewPaneOpening;
  /** Default `'icon'`. */
  readonly commentMarkers?: CommentMarkerStyle;
}

/** The comment fields of the editor snapshot (`EditorSnapshot`). @public */
export interface CommentSnapshot {
  /**
   * The current slide's comment threads, in file order. Null without a review module
   * (`@pptx-editor.dev/pro`); the comment commands still work. Reference-stable while unchanged.
   */
  readonly comments?: readonly CommentThread[] | null;
  /** The thread a comment marker or `selectComment` picked, or null. */
  readonly activeComment?: string | null;
  /**
   * The name that signs a new comment or reply: the editor's `author` option (`setAuthor`), else
   * the collaboration identity's name, else null. With null, the editor refuses a new comment
   * that names no author (`disabled.noCommentAuthor`).
   */
  readonly commentAuthor?: string | null;
}
