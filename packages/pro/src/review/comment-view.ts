/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comment review view, framework-free: the items of the current slide's threads, the actions
// on them, the date text, and the keys of a comment field. The React and Vue `useReview` and the
// `PptxEditorReview` parts of both entries read these, so the two adapters behave the same. The
// edits are engine commands (`addComment`, `replyToComment`, `editComment`, `deleteComment`,
// `resolveComment`, `reopenComment`).

import {
  commentAuthorSlot,
  commentInitials,
  requestPane,
  type CommentAnchor,
  type CommentFormat,
  type CommentThread,
  type EditorSelection,
  type EditorSnapshot,
  type PptxEditorInstance,
  type ShapeSummary,
  type SlideComment,
} from '@pptx-editor.dev/core/editor';

/** The longest comment text the commands take. */
export const COMMENT_TEXT_MAX = 10_000;

/**
 * One comment or reply of the current slide, as the review hooks and parts read it. A thread's
 * replies follow it in `UseReviewReturn.items`, and the thread names them in `replyIds`. Text and
 * names come from the file: render them as text, never as markup.
 *
 * @public
 */
export interface ReviewItemView {
  /** The key the actions take. Stable while the comment exists. */
  readonly key: string;
  /** The comment id, as the comment commands take it. Equal to `key`. */
  readonly id: string;
  /** Always `comment`: PresentationML has no tracked-change markup. */
  readonly kind: 'comment';
  /** The author's display name, or an empty string when the file names none. */
  readonly author: string;
  /** The avatar text: the file's initials, else the first letters of the author's name. */
  readonly initials: string;
  /** The date as the file stores it (ISO 8601), or absent. */
  readonly date?: string;
  /** The text, with one line feed between paragraphs. */
  readonly text: string;
  /** The file format of the thread. */
  readonly format: CommentFormat;
  /** Whether the thread is resolved. A reply reports its thread's state. */
  readonly resolved: boolean;
  /** Where the thread points on the slide. A reply reports its thread's anchor. */
  readonly anchor: CommentAnchor;
  /** The name of the element the thread points at, or absent for the slide or a place. */
  readonly anchorName?: string;
  /** For a reply, the id of its thread; absent on a thread. */
  readonly parentId?: string;
  /** For a thread, the ids of its replies, oldest first; empty on a reply. */
  readonly replyIds: readonly string[];
  /** Whether the thread is the picked one (`activeKey`). A reply reports its thread's state. */
  readonly isActive: boolean;
  /** Whether `edit` can replace the text. False for content that the editor cannot write back. */
  readonly editable: boolean;
  /** Whether `resolve` and `reopen` apply. False for the legacy format, which has no state. */
  readonly resolvable: boolean;
}

/**
 * The element that a new comment points at: the one selected element of the slide's top level.
 * Null when the comment points at the slide. @public
 */
export interface ReviewCommentTarget {
  readonly nodeId: string;
  readonly name: string;
}

/** Options of `UseReviewReturn.setActive`. @public */
export interface ReviewActivationOptions {
  /**
   * Also ask for the **Comments** tab of the side panel (`requestPane`), as a press on the
   * thread's marker does. Default `false`.
   */
  readonly announce?: boolean;
}

/** The resolved look of one comment author. @public */
export interface ReviewAuthorInfo {
  readonly author: string;
  /** The color slot, 0 to 7: the token `--doc-review-author-<slot>`. */
  readonly slot: number;
  /** The CSS color: `var(--doc-review-author-<slot>)`. */
  readonly color: string;
}

/** The look of `author`, from the shared slot rule of the canvas markers. */
export function reviewAuthorInfo(author: string): ReviewAuthorInfo {
  const slot = commentAuthorSlot(author);
  return { author, slot, color: `var(--doc-review-author-${slot})` };
}

/** The author hooks of an element: `data-review-author` and `data-review-author-slot`. */
export function authorHooks(author: string): Record<string, string | number> {
  if (!author) return {};
  return { 'data-review-author': author, 'data-review-author-slot': commentAuthorSlot(author) };
}

/** The accent variable of `author`, for an element's `style`. */
export function authorStyle(author: string): Record<string, string> {
  return author ? { '--doc-review-author-current': reviewAuthorInfo(author).color } : {};
}

/** The name of a slide element by node id, at any depth of the group tree. */
function shapeName(shapes: readonly ShapeSummary[] | undefined, id: string): string | undefined {
  for (const s of shapes ?? []) {
    if (s.nodeId === id) return s.name;
    const hit = shapeName(s.children, id);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

/** The element a new comment points at, as `addComment` picks it, or null for the slide. */
export function commentTarget(
  selection: EditorSelection,
  shapes: readonly ShapeSummary[] | undefined
): ReviewCommentTarget | null {
  const id =
    selection.kind === 'shapes' && selection.nodeIds.length === 1
      ? selection.nodeIds[0]
      : selection.kind === 'text' || selection.kind === 'cells'
        ? selection.nodeId
        : undefined;
  if (id === undefined) return null;
  const name = shapes?.find((s) => s.nodeId === id)?.name;
  return name === undefined ? null : { nodeId: id, name };
}

/**
 * The element of a draft from a menu while the slide still holds it at its top level, else null:
 * the draft then points at the selected element, or at the slide.
 */
export function liveDraftTarget(
  target: ReviewCommentTarget | null,
  shapes: readonly ShapeSummary[] | undefined
): ReviewCommentTarget | null {
  return target && shapes?.some((s) => s.nodeId === target.nodeId) ? target : null;
}

/**
 * What a field that edits or answers `item` remembers of it when it opens: the author, the date,
 * and the text. When the comment under the field's key no longer matches (a peer deleted it, and
 * a repeat of its id took the key), the field closes and saves nothing.
 */
export function itemIdentity(item: ReviewItemView): string {
  return [item.author, item.date ?? '', item.text].join('\u0000');
}

/** The items of the snapshot's threads: each thread, then its replies. */
export function reviewItems(
  editor: PptxEditorInstance | null,
  snapshot: EditorSnapshot | null
): readonly ReviewItemView[] {
  const threads: readonly CommentThread[] = snapshot?.comments ?? [];
  if (!editor || !snapshot || threads.length === 0) return [];
  // From the thread data, not one `editor.can` per item: each `can` reads the whole list, so a
  // pane of many threads would take quadratic time. `editComment` and the status commands are
  // refused only in a read-only deck, for a comment the editor cannot write back, and for the
  // legacy format's status; the review tests hold the two in agreement.
  const writable = !reviewReadOnly(snapshot);
  const out: ReviewItemView[] = [];
  for (const thread of threads) {
    const isActive = thread.id === snapshot.activeComment;
    const anchorName =
      thread.anchor.kind === 'shape' ? shapeName(snapshot.shapes, thread.anchor.nodeId) : undefined;
    const resolvable = writable && thread.format === 'modern';
    const item = (c: SlideComment, parentId?: string): ReviewItemView => ({
      key: c.id,
      id: c.id,
      kind: 'comment',
      author: c.author,
      initials: commentInitials(c),
      ...(c.date !== undefined ? { date: c.date } : {}),
      text: c.text,
      format: thread.format,
      resolved: thread.resolved,
      anchor: thread.anchor,
      ...(anchorName !== undefined ? { anchorName } : {}),
      ...(parentId !== undefined ? { parentId } : {}),
      replyIds: parentId === undefined ? thread.replies.map((r) => r.id) : [],
      isActive,
      editable: writable && c.textEditable,
      resolvable,
    });
    out.push(item(thread));
    for (const reply of thread.replies) out.push(item(reply, thread.id));
  }
  return out;
}

/** Whether the snapshot is a ready deck open for editing. */
export function reviewReadOnly(snapshot: EditorSnapshot | null): boolean {
  return !snapshot || snapshot.status !== 'ready' || snapshot.mode !== 'edit';
}

/** Whether the snapshot has an open deck. */
export function reviewReady(snapshot: EditorSnapshot | null): boolean {
  return !!snapshot && snapshot.status === 'ready';
}

/** The actions of `UseReviewReturn`, over one editor. Each reports whether the editor applied it. */
export interface ReviewActions {
  setActive(key: string | null, options?: ReviewActivationOptions): boolean;
  comment(text: string, author?: string, target?: { readonly nodeId: string }): boolean;
  reply(item: ReviewItemView, text: string, author?: string): boolean;
  edit(item: ReviewItemView, text: string): boolean;
  remove(item: ReviewItemView): boolean;
  resolve(item: ReviewItemView): boolean;
  reopen(item: ReviewItemView): boolean;
  setAuthor(author: string | undefined): void;
}

/** The actions over `editor`, or actions that do nothing and report false without one. */
export function reviewActions(editor: PptxEditorInstance | null): ReviewActions {
  const run = (command: Parameters<PptxEditorInstance['exec']>[0]): boolean =>
    editor ? editor.exec(command).ok : false;
  const author = (name: string | undefined) => (name ? { author: name } : {});
  return {
    setActive: (key, options) => {
      if (!editor) return false;
      if (key !== null && !(editor.snapshot().comments ?? []).some((t) => t.id === key))
        return false;
      editor.selectComment(key);
      if (key !== null && options?.announce) requestPane(editor, 'comments');
      return true;
    },
    comment: (text, name, target) =>
      text.trim() !== '' &&
      run({
        type: 'addComment',
        text,
        ...author(name),
        ...(target ? { nodeId: target.nodeId } : {}),
      }),
    reply: (item, text, name) =>
      text.trim() !== '' &&
      run({ type: 'replyToComment', commentId: item.parentId ?? item.id, text, ...author(name) }),
    edit: (item, text) =>
      text.trim() !== '' && run({ type: 'editComment', commentId: item.id, text }),
    remove: (item) => run({ type: 'deleteComment', commentId: item.id }),
    resolve: (item) =>
      item.resolved || run({ type: 'resolveComment', commentId: item.parentId ?? item.id }),
    reopen: (item) =>
      !item.resolved || run({ type: 'reopenComment', commentId: item.parentId ?? item.id }),
    setAuthor: (name) => editor?.setAuthor(name),
  };
}

/** `2026-11-07T10:00:00.000`: an ISO date and time with no zone, which the file means as UTC. */
const NO_ZONE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?$/;
let dateFormat: Intl.DateTimeFormat | null = null;

/**
 * The date of a comment as a card shows it, after DOCX Editor's review rail
 * (docx-editor/packages/pro/src/review/review-date-format.ts at adf0e178c): month, day, and time
 * in the reader's locale and time zone. A date the file stores with no zone is UTC. An unreadable
 * date or none shows nothing.
 */
export function commentDateText(date: string | undefined): string {
  if (!date) return '';
  const ms = Date.parse(NO_ZONE.test(date) ? `${date}Z` : date);
  if (Number.isNaN(ms)) return '';
  dateFormat ??= new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  });
  return dateFormat.format(ms);
}

/**
 * What a key in a comment field does, after DOCX Editor's compose box: Enter posts, Shift+Enter
 * starts a new line, and Escape cancels. A key that composes text (an input method) does nothing.
 */
export function commentFieldKey(e: {
  readonly key: string;
  readonly shiftKey: boolean;
  readonly isComposing?: boolean;
}): 'post' | 'cancel' | null {
  if (e.isComposing) return null;
  if (e.key === 'Enter' && !e.shiftKey) return 'post';
  if (e.key === 'Escape') return 'cancel';
  return null;
}

/** Whether a press on a card landed on one of its controls, which the card does not answer. */
export function isCardControl(target: EventTarget | null): boolean {
  const el = target as { closest?: (s: string) => unknown } | null;
  return !!el?.closest?.('button, textarea, input, a, [role="button"]');
}
