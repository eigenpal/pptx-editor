// The comment commands: add, reply, edit, delete, resolve, and reopen. The editor's `compile`
// (`editor/object-commands.ts`) and the automation host (`plan-comments.ts`) both send them here;
// each compiles into the caller's batch, so one transaction and one undo step. A new thread uses
// the threaded format, unless the deck has legacy comments and no threaded ones; a reply, an
// edit, and a delete keep the format of their thread (`comment-modern.ts`, `comment-legacy.ts`).
// The legacy format is ECMA-376 Part 1, 19.4: `p:cm` (`CT_Comment`, 19.4.1) in a `p:cmLst`
// (`CT_CommentList`, 19.4.4), by an author of the `p:cmAuthorLst` (19.4.3).

import { presentationInfo, xmlPartOf, child, elements } from '../../store/index.ts';
import type {
  CommentCommand,
  EditorCommand,
  EditorSelection,
  TextPoint,
} from '../../contracts/neutral.ts';
import { rangeChars, shapeParagraphLengths } from './comment-text-range.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { deckUses, legacyFacts, modernFacts } from './comment-read.ts';
import { commentAuthor, commentDate, commentText, splitCommentId } from './comment-write.ts';
import {
  addModernComment,
  deleteModern,
  editModern,
  findModern,
  replyModern,
  setModernStatus,
  type TextChars,
} from './comment-modern.ts';
import {
  addLegacyComment,
  deleteLegacy,
  editLegacy,
  findLegacy,
  replyLegacy,
} from './comment-legacy.ts';

/**
 * What a comment command compiles against: the editor's `CommandContext`, or the automation
 * host's view of one slide (no selection, no snapshot identities).
 */
export interface CommentContext {
  /** The zero-based index of the slide a new thread goes on. */
  readonly slide: number;
  /** The selection a new thread without `nodeId` points at. */
  readonly selection: EditorSelection;
  /** The author of new comments and replies when the command names none. */
  readonly author?: string;
  /** The `identity` at a comment id, and whether it named another one before. @internal */
  readonly commentAt?: (id: string) => { readonly identity?: string; readonly moved: boolean };
}

/** What a comment command did: its history label, and the id of the comment it wrote. */
export interface CommentEffect {
  readonly label: string;
  /** The comment id (`<slide part>#<key>`) of the new thread or reply. */
  readonly commentId?: string;
}

const TYPES: ReadonlySet<string> = new Set([
  'addComment',
  'replyToComment',
  'editComment',
  'deleteComment',
  'resolveComment',
  'reopenComment',
]);

export function isCommentCommand(command: EditorCommand): command is CommentCommand {
  return TYPES.has(command.type);
}

/** Whether `v` is a `TextPoint` the command can read. */
const isPoint = (v: unknown): v is TextPoint => {
  const p = v as { paragraph?: unknown; offset?: unknown } | null;
  return (
    typeof p === 'object' &&
    p !== null &&
    Number.isSafeInteger(p.paragraph) &&
    Number.isSafeInteger(p.offset) &&
    (p.paragraph as number) >= 0 &&
    (p.offset as number) >= 0
  );
};

/**
 * The top-level element a new thread points at, and the range of its text: `nodeId` and
 * `range`, else the selected text of a shape, else the one selected element.
 */
function targetOf(
  ctx: CommentContext,
  batch: OpBatch,
  slidePart: string,
  nodeId: unknown,
  range: unknown
): { readonly id: string; readonly chars?: TextChars } | undefined {
  const sel = ctx.selection;
  if (range !== undefined && nodeId === undefined)
    throw new CommandRefusal('disabled.noCommentTarget');
  const id =
    nodeId !== undefined
      ? nodeId
      : sel.kind === 'shapes' && sel.nodeIds.length === 1
        ? sel.nodeIds[0]
        : sel.kind === 'text' || sel.kind === 'cells'
          ? sel.nodeId
          : undefined;
  if (id === undefined) return undefined;
  const tree = child(child(xmlPartOf(batch.pkg, slidePart)?.root, 'p:cSld'), 'p:spTree');
  const el = typeof id === 'string' ? elements(tree).find((e) => e.id === id) : undefined;
  if (!el) {
    // An element inside a group, or one that is gone: an explicit id is refused, a selection
    // points the thread at the slide.
    if (nodeId !== undefined) throw new CommandRefusal('disabled.noCommentTarget');
    return undefined;
  }
  // The selected text of a shape, not of a table cell or a diagram node: a text moniker counts
  // the characters of a shape's own text body.
  const picked =
    range !== undefined
      ? range
      : nodeId === undefined && sel.kind === 'text' && !sel.cell && !sel.diagramNode
        ? { start: sel.anchor, end: sel.focus }
        : undefined;
  if (picked === undefined) return { id: el.id };
  const r = picked as { start?: unknown; end?: unknown } | null;
  const lengths = shapeParagraphLengths(el, ctx.slide + 1);
  const chars =
    typeof r === 'object' && r !== null && isPoint(r.start) && isPoint(r.end) && lengths
      ? rangeChars(lengths, { start: r.start, end: r.end })
      : null;
  if (chars) return { id: el.id, chars };
  // A range the host named must hold text; a collapsed caret points at the element.
  if (range !== undefined) throw new CommandRefusal('disabled.noCommentTarget');
  return { id: el.id };
}

/**
 * Refuse a command whose `expect` does not match the comment it reached (`errors.commentChanged`):
 * another author or date, or, for an id that holds a place among twins, other text. An id with
 * a place that reaches a comment without a twin is refused too: its twin is gone, and the id
 * may have named that twin. Before any write, so nothing changes.
 */
function checkExpect(
  ctx: CommentContext,
  commentId: string,
  expect: unknown,
  facts: () => { readonly author: string; readonly date?: string; readonly text: string },
  found: { readonly placed: boolean; readonly group: number; readonly twinTexts: () => string[] }
): void {
  if (expect === undefined) return;
  const e = expect as {
    author?: unknown;
    date?: unknown;
    text?: unknown;
    identity?: unknown;
  } | null;
  if (
    typeof e !== 'object' ||
    e === null ||
    typeof e.author !== 'string' ||
    typeof e.text !== 'string' ||
    (e.date !== undefined && typeof e.date !== 'string') ||
    (e.identity !== undefined && typeof e.identity !== 'string')
  )
    throw new CommandRefusal('errors.commentChanged');
  // A place among twins that no twin shares any more: the id was a twin's, maybe a gone one.
  if (found.placed && found.group < 2) throw new CommandRefusal('errors.noComment');
  const now = facts();
  if (now.author !== e.author || now.date !== e.date || (found.placed && now.text !== e.text))
    throw new CommandRefusal('errors.commentChanged');
  const at = ctx.commentAt?.(commentId);
  // The identity names one list node, so it tells apart twins that text cannot.
  if (e.identity !== undefined) {
    if (at?.identity !== e.identity) throw new CommandRefusal('errors.commentChanged');
    return;
  }
  if (!found.placed) return;
  // Without the identity: a twin with the same text, which nothing in the file tells apart.
  if (found.twinTexts().includes(now.text)) throw new CommandRefusal('errors.commentAmbiguous');
  // An id that a twin took from another since the last snapshot may be the host's old id for
  // that one. The next snapshot shows it at its new identity, and ends the refusal.
  if (at?.moved) throw new CommandRefusal('errors.commentChanged');
}

/** The date of a new reply: the command's, else now. */
const replyDate = (command: { readonly date?: string }): string => commentDate(command.date);

/** A thread id key: `cm-…` is a legacy comment, a GUID a threaded one. */
const isLegacyKey = (key: string): boolean => key.startsWith('cm-');

export function compileCommentCommand(
  ctx: CommentContext,
  batch: OpBatch,
  command: CommentCommand
): CommentEffect {
  const label = command.type;
  if (command.type === 'addComment') {
    const slidePart = presentationInfo(batch.pkg).slides[ctx.slide]?.partName;
    if (!slidePart) throw new CommandRefusal('errors.slideNotFound');
    const text = commentText(command.text);
    const author = commentAuthor(command, ctx.author);
    const date = commentDate(command.date);
    const target = targetOf(ctx, batch, slidePart, command.nodeId, command.range);
    const legacy = deckUses(batch.pkg, 'legacy') && !deckUses(batch.pkg, 'modern');
    // The legacy format has no text anchors: its thread points at the element.
    const key = legacy
      ? addLegacyComment(batch, slidePart, text, author, date, target?.id)
      : addModernComment(batch, slidePart, text, author, date, target?.id, target?.chars);
    return { label, commentId: `${slidePart}#${key}` };
  }
  const { slidePart, key } = splitCommentId(batch, command.commentId);
  if (isLegacyKey(key)) {
    if (command.type === 'resolveComment' || command.type === 'reopenComment')
      throw new CommandRefusal('disabled.commentNoStatus');
    const found = findLegacy(batch, slidePart, key);
    checkExpect(
      ctx,
      command.commentId,
      command.expect,
      () => legacyFacts(batch.pkg, found.cm),
      found
    );
    switch (command.type) {
      case 'replyToComment': {
        const text = commentText(command.text);
        const author = commentAuthor(command, ctx.author);
        const reply = replyLegacy(
          batch,
          found.part,
          found.thread,
          text,
          author,
          replyDate(command)
        );
        return { label, commentId: `${slidePart}#${reply}` };
      }
      case 'editComment':
        editLegacy(batch, found.part, found.cm, commentText(command.text));
        break;
      case 'deleteComment':
        deleteLegacy(batch, found.part, found.cm);
        break;
    }
    return { label };
  }
  const found = findModern(batch, slidePart, key);
  const el = found.reply ?? found.thread;
  checkExpect(ctx, command.commentId, command.expect, () => modernFacts(batch.pkg, el), found);
  switch (command.type) {
    case 'replyToComment': {
      const text = commentText(command.text);
      const author = commentAuthor(command, ctx.author);
      const reply = replyModern(batch, found.part, found.thread, text, author, replyDate(command));
      return { label, commentId: `${slidePart}#${reply}` };
    }
    case 'editComment':
      editModern(batch, found.part, found.reply ?? found.thread, commentText(command.text));
      break;
    case 'deleteComment':
      deleteModern(batch, found.part, found.thread, found.reply);
      break;
    case 'resolveComment':
    case 'reopenComment':
      setModernStatus(batch, found.part, found.thread, command.type === 'resolveComment');
      break;
  }
  return { label };
}
