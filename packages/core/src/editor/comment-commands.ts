// The comment commands: add, reply, edit, delete, resolve, and reopen. `commands.ts` sends them
// here; each compiles to one transaction, so one undo step. A new thread uses the threaded
// format, unless the deck has legacy comments and no threaded ones; a reply, an edit, and a
// delete keep the format of their thread (`comment-modern.ts`, `comment-legacy.ts`).

import { presentationInfo, xmlPartOf, child, elements } from '../store/index.ts';
import type { CommentCommand, EditorCommand } from '../contracts/index.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { deckUses } from '../automation/commands/comment-read.ts';
import {
  commentAuthor,
  commentDate,
  commentText,
  splitCommentId,
} from '../automation/commands/comment-write.ts';
import {
  addModernComment,
  deleteModern,
  editModern,
  findModern,
  replyModern,
  setModernStatus,
} from '../automation/commands/comment-modern.ts';
import {
  addLegacyComment,
  deleteLegacy,
  editLegacy,
  findLegacy,
  replyLegacy,
} from '../automation/commands/comment-legacy.ts';

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

/** The top-level element a new thread points at: `nodeId`, or the one selected element. */
function targetOf(ctx: CommandContext, batch: OpBatch, slidePart: string, nodeId: unknown) {
  const sel = ctx.selection;
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
  if (typeof id === 'string' && elements(tree).some((e) => e.id === id)) return id;
  // An element inside a group, or one that is gone: an explicit id is refused, a selection
  // points the thread at the slide.
  if (nodeId !== undefined) throw new CommandRefusal('disabled.noCommentTarget');
  return undefined;
}

/** A thread id key: `cm-…` is a legacy comment, a GUID a threaded one. */
const isLegacyKey = (key: string): boolean => key.startsWith('cm-');

export function compileCommentCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: CommentCommand
): CommandEffect {
  const label = command.type;
  if (command.type === 'addComment') {
    const slidePart = presentationInfo(batch.pkg).slides[ctx.slide]?.partName;
    if (!slidePart) throw new CommandRefusal('errors.slideNotFound');
    const text = commentText(command.text);
    const author = commentAuthor(command, ctx.author);
    const date = commentDate(command.date);
    const target = targetOf(ctx, batch, slidePart, command.nodeId);
    const legacy = deckUses(batch.pkg, 'legacy') && !deckUses(batch.pkg, 'modern');
    if (legacy) addLegacyComment(batch, slidePart, text, author, date, target);
    else addModernComment(batch, slidePart, text, author, date, target);
    return { label };
  }
  const { slidePart, key } = splitCommentId(batch, command.commentId);
  if (isLegacyKey(key)) {
    if (command.type === 'resolveComment' || command.type === 'reopenComment')
      throw new CommandRefusal('disabled.commentNoStatus');
    const found = findLegacy(batch, slidePart, key);
    switch (command.type) {
      case 'replyToComment': {
        const text = commentText(command.text);
        const author = commentAuthor(command, ctx.author);
        replyLegacy(batch, found.part, found.thread, text, author, commentDate(command.date));
        break;
      }
      case 'editComment':
        editLegacy(batch, found.part, found.cm, commentText(command.text));
        break;
      case 'deleteComment':
        deleteLegacy(batch, slidePart, found.part, found.cm);
        break;
    }
    return { label };
  }
  const found = findModern(batch, slidePart, key);
  switch (command.type) {
    case 'replyToComment': {
      const text = commentText(command.text);
      const author = commentAuthor(command, ctx.author);
      replyModern(batch, found.part, found.thread, text, author, commentDate(command.date));
      break;
    }
    case 'editComment':
      editModern(batch, found.part, found.reply ?? found.thread, commentText(command.text));
      break;
    case 'deleteComment':
      deleteModern(batch, slidePart, found.part, found.thread, found.reply);
      break;
    case 'resolveComment':
    case 'reopenComment':
      setModernStatus(batch, found.part, found.thread, command.type === 'resolveComment');
      break;
  }
  return { label };
}
