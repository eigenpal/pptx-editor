// The comments of a slide: read both formats, and add, reply to, edit, resolve, reopen, and
// delete through the comment commands the editor runs (`commands/comment-commands.ts`), so a
// write here and the same edit in the editor produce the same package.
//
// The legacy format is ECMA-376 Part 1, 19.4: `p:cm` (`CT_Comment`, 19.4.1) in the slide's
// comments part (13.3.2), by a `p:cmAuthor` (`CT_CommentAuthor`, 19.4.2) of the comment authors
// part (13.3.1), at `p:pos` (19.4.5). It has no resolved state and no text anchor. The threaded
// format is a part that the slide names in a `p:ext` of its `p:extLst` (19.3.1.20); it is not
// part of ECMA-376.
//
// Comment text and author names come from the file, or from the caller: both are plain text. The
// writers check them (`comment-write.ts`): 1 to 10,000 characters of text, an author name of 1 to
// 255 characters, no control characters, and XML-safe on save.

import { attr, child, nodeById, xmlPartOf, type Element } from '../store/index.ts';
import type {
  CommentAnchor,
  CommentCommand,
  CommentExpect,
  SlideComment,
  TextPoint,
} from '../contracts/neutral.ts';
import {
  compileCommentCommand,
  type CommentContext,
  type CommentEffect,
} from './commands/comment-commands.ts';
import { deckUses, readSlideComments } from './commands/comment-read.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import type { AutomationCommentExpect, AutomationTextSpan } from './operations.ts';
import type {
  AutomationCommentAnchor,
  AutomationCommentRead,
  AutomationRef,
  AutomationValue,
} from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  cNvPrOf,
  edit,
  shapeAt,
  slideOf,
  slidePartOf,
  targetOf,
  toPoints,
  type PlanContext,
} from './targets.ts';
import { textOffsetOf, textPointOf } from './text-offsets.ts';

/** What a comment operation plans against. */
export interface CommentPlanContext extends PlanContext {
  /**
   * The author of a new comment or reply that names none: the host's `author` option, or the
   * editor's author in a browser host. Without one, such a write refuses.
   */
  readonly commentAuthor?: string;
}

/** The longest comment id key a write takes. A key from a file is far shorter. */
const MAX_KEY = 1024;

/** The slide a ref names: its part (refused when it did not open) and its index. */
function slideAt(ctx: PlanContext, ref: AutomationRef): { part: string; index: number } {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const part = slidePartOf(ctx.batch.pkg, sldId);
  return { part, index: slideOf(ctx.batch.pkg, sldId).index };
}

/** The element a comment anchor names, or undefined when it is gone. */
function elementOf(ctx: PlanContext, part: string, nodeId: string): Element | undefined {
  const xml = xmlPartOf(ctx.batch.pkg, part);
  const node = xml ? nodeById(xml, nodeId) : undefined;
  return node?.type === 'element' ? node : undefined;
}

/** A thread anchor as the protocol shows it: shape ids and plain text offsets, points. */
function anchorOf(
  ctx: PlanContext,
  part: string,
  slideNumber: number,
  anchor: CommentAnchor
): AutomationCommentAnchor {
  switch (anchor.kind) {
    case 'slide':
      return { kind: 'slide' };
    case 'point':
      return { kind: 'point', left: toPoints(anchor.x), top: toPoints(anchor.y) };
    case 'shape':
    case 'text': {
      // An element the anchor no longer finds leaves the thread on the slide.
      const el = elementOf(ctx, part, anchor.nodeId);
      const shapeId = el ? attr(cNvPrOf(el), 'id') : undefined;
      if (!el || shapeId === undefined) return { kind: 'slide' };
      const body = anchor.kind === 'text' ? child(el, 'p:txBody') : undefined;
      if (anchor.kind === 'shape' || !body) return { kind: 'shape', shapeId };
      const start = textOffsetOf(body, anchor.start, slideNumber);
      const end = textOffsetOf(body, anchor.end, slideNumber);
      return { kind: 'text', shapeId, start, length: Math.max(0, end - start) };
    }
  }
}

/** One comment as the protocol shows it, with its id key in place of the full id. */
function commentOf(part: string, c: SlideComment): AutomationCommentRead {
  return {
    id: c.id.slice(part.length + 1),
    author: c.author,
    initials: c.initials,
    ...(c.date === undefined ? {} : { date: c.date }),
    text: c.text,
  };
}

/** Every comment thread of a slide, in file order: the legacy list first, then the threaded. */
export function readComments(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  const comments = readSlideComments(ctx.batch.pkg, part).map((thread) => ({
    ...commentOf(part, thread),
    format: thread.format,
    resolved: thread.resolved,
    anchor: anchorOf(ctx, part, index + 1, thread.anchor),
    replies: thread.replies.map((reply) => commentOf(part, reply)),
  }));
  return { kind: 'comments', comments };
}

/** A comment id key, or an `invalid-argument` failure. */
function keyOf(value: unknown): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > MAX_KEY)
    badArgument('comment');
  return value;
}

/** An author name, initials, or text the caller passed: a string, or absent. */
function optionalText(value: unknown, name: string): string | undefined {
  if (value !== undefined && typeof value !== 'string') badArgument(name);
  return value as string | undefined;
}

/** The engine's check of what the caller read, from the protocol's. */
function expectOf(value: AutomationCommentExpect | undefined): CommentExpect | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'object' || value === null) badArgument('comment-expect');
  return value;
}

/**
 * Run one comment command on slide `index`. A refusal of the caller's input is an argument
 * failure; an id with no comment is `item-not-found`; any other refusal is the edit's.
 */
function run(ctx: CommentPlanContext, index: number, command: CommentCommand): CommentEffect {
  const context: CommentContext = {
    slide: index,
    selection: { kind: 'none' },
    ...(ctx.commentAuthor === undefined ? {} : { author: ctx.commentAuthor }),
  };
  try {
    return compileCommentCommand(context, ctx.batch, command);
  } catch (e) {
    if (e instanceof CommandRefusal) {
      if (e.reason === 'errors.commentText' || e.reason === 'errors.commentAuthor')
        throw new AutomationFailure('invalid-argument', 'invalid comment input', e.reason);
      if (e.reason === 'errors.noComment')
        throw new AutomationFailure('item-not-found', 'no comment with that id', e.reason);
    }
    return edit(() => {
      throw e;
    });
  }
}

/** The id key of a comment the command wrote, as `readComments` shows it. */
function keyAnswer(part: string, effect: CommentEffect): AutomationValue {
  const id = effect.commentId ?? '';
  return { kind: 'text', text: id.slice(part.length + 1) };
}

/**
 * The text range of a new thread: two points in the text of `shape`, which must be a shape
 * (`p:sp`) with a text body that holds the whole span, and at least one character of it.
 */
function rangeOf(
  shape: ReturnType<typeof shapeAt>,
  slideNumber: number,
  span: AutomationTextSpan
): { start: TextPoint; end: TextPoint } {
  const { start, length } = (span ?? {}) as Partial<AutomationTextSpan>;
  if (
    typeof start !== 'number' ||
    typeof length !== 'number' ||
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(length) ||
    start < 0 ||
    length < 1
  )
    badArgument('comment-span');
  const body = shape.el.local === 'sp' ? child(shape.el, 'p:txBody') : undefined;
  const from = body ? textPointOf(body, start, slideNumber) : undefined;
  const to = body ? textPointOf(body, start + length, slideNumber) : undefined;
  if (!from || !to) badArgument('comment-span');
  return { start: from, end: to };
}

export function addComment(
  ctx: CommentPlanContext,
  ref: AutomationRef,
  text: string,
  shapeRef: AutomationRef | undefined,
  span: AutomationTextSpan | undefined,
  author: string | undefined,
  initials: string | undefined
): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  if (typeof text !== 'string') badArgument('comment-text');
  const name = optionalText(author, 'comment-author');
  const short = optionalText(initials, 'comment-initials');
  if (span !== undefined && shapeRef === undefined) badArgument('comment-span');
  let nodeId: string | undefined;
  let range: { start: TextPoint; end: TextPoint } | undefined;
  if (shapeRef !== undefined) {
    const shape = shapeAt(ctx, shapeRef);
    // A thread points at an element on the top level of its own slide.
    if (shape.part !== part || shape.level !== 0) badArgument('comment-shape');
    nodeId = shape.el.id;
    if (span !== undefined) range = rangeOf(shape, index + 1, span);
  }
  // A new thread of a deck with only legacy comments is legacy, and the legacy format has no
  // text anchor (ECMA-376 Part 1, 19.4.1 holds only `p:pos`): refuse, not drop, the range.
  if (range && deckUses(ctx.batch.pkg, 'legacy') && !deckUses(ctx.batch.pkg, 'modern'))
    throw new AutomationFailure(
      'unsupported-content',
      'the legacy comment format has no text anchor',
      'disabled.commentNoTextRange'
    );
  const effect = run(ctx, index, {
    type: 'addComment',
    text,
    ...(nodeId === undefined ? {} : { nodeId }),
    ...(range === undefined ? {} : { range }),
    ...(name === undefined ? {} : { author: name }),
    ...(short === undefined ? {} : { initials: short }),
  });
  return keyAnswer(part, effect);
}

export function replyToComment(
  ctx: CommentPlanContext,
  ref: AutomationRef,
  comment: string,
  text: string,
  author: string | undefined,
  initials: string | undefined,
  expect: AutomationCommentExpect | undefined
): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  const key = keyOf(comment);
  if (typeof text !== 'string') badArgument('comment-text');
  const name = optionalText(author, 'comment-author');
  const short = optionalText(initials, 'comment-initials');
  const checked = expectOf(expect);
  const effect = run(ctx, index, {
    type: 'replyToComment',
    commentId: `${part}#${key}`,
    text,
    ...(name === undefined ? {} : { author: name }),
    ...(short === undefined ? {} : { initials: short }),
    ...(checked === undefined ? {} : { expect: checked }),
  });
  return keyAnswer(part, effect);
}

export function setCommentText(
  ctx: CommentPlanContext,
  ref: AutomationRef,
  comment: string,
  text: string,
  expect: AutomationCommentExpect | undefined
): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  const key = keyOf(comment);
  if (typeof text !== 'string') badArgument('comment-text');
  const checked = expectOf(expect);
  run(ctx, index, {
    type: 'editComment',
    commentId: `${part}#${key}`,
    text,
    ...(checked === undefined ? {} : { expect: checked }),
  });
  return { kind: 'applied' };
}

export function setCommentResolved(
  ctx: CommentPlanContext,
  ref: AutomationRef,
  comment: string,
  resolved: boolean,
  expect: AutomationCommentExpect | undefined
): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  const key = keyOf(comment);
  if (typeof resolved !== 'boolean') badArgument('resolved');
  const checked = expectOf(expect);
  run(ctx, index, {
    type: resolved ? 'resolveComment' : 'reopenComment',
    commentId: `${part}#${key}`,
    ...(checked === undefined ? {} : { expect: checked }),
  });
  return { kind: 'applied' };
}

export function deleteComment(
  ctx: CommentPlanContext,
  ref: AutomationRef,
  comment: string,
  expect: AutomationCommentExpect | undefined
): AutomationValue {
  const { part, index } = slideAt(ctx, ref);
  const key = keyOf(comment);
  const checked = expectOf(expect);
  run(ctx, index, {
    type: 'deleteComment',
    commentId: `${part}#${key}`,
    ...(checked === undefined ? {} : { expect: checked }),
  });
  return { kind: 'applied' };
}
