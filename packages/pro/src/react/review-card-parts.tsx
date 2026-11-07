/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The card and its parts, after DOCX Editor's review card
// (docx-editor/packages/pro/src/react/review-card-parts.tsx at adf0e178c): the author avatar in
// the author's color, the name and the date, the actions in the head, the text, the replies, and
// the reply line of the picked thread. A part passed as a child of a card replaces the packaged
// copy of it. Comment text and names come from the file and render as text only. The Vue twin is
// `../vue/review-card-parts.tsx`.
/** @jsxImportSource react */

import {
  Children,
  Fragment,
  cloneElement,
  isValidElement,
  useContext,
  useMemo,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from 'react';
import { Icon } from '@pptx-editor.dev/react';
import {
  authorHooks,
  authorStyle,
  commentDateText,
  isCardControl,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { ReviewItemContext, usePane, useReviewLabel } from './review-context.ts';
import type { ReviewActionProps, ReviewPartProps } from './review-types.ts';
import { ReviewActionSlot } from './review-action-slot.tsx';
import { ReviewEditBox, ReviewReply } from './review-compose.tsx';

/** A part's marker, which the card and the pane read to replace a packaged part in place. */
export type PartMarked<T> = T & { readonly pptxReviewPart: string };

export function markPart<T extends object>(component: T, part: string): PartMarked<T> {
  return Object.assign(component, { pptxReviewPart: part });
}

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

/** `asChild`: the part's props on its single child element. */
function slotted(children: ReactNode, shared: Record<string, unknown>): ReactNode {
  const child = Children.only(children);
  if (!isValidElement<Record<string, unknown>>(child)) return null;
  const own = child.props;
  return cloneElement(child, {
    ...shared,
    ...own,
    className: cx(shared.className as string, own.className as string),
  });
}

/** Map each child part to its marker, so the card replaces the packaged copy in place. */
export function partOverrides(children: ReactNode): Record<string, ReactNode> {
  const found: Record<string, ReactNode> = {};
  const extra: ReactNode[] = [];
  const visit = (node: ReactNode): void => {
    if (Array.isArray(node)) return void node.forEach(visit);
    if (!isValidElement(node)) {
      if (node !== null && node !== undefined && node !== false) extra.push(node);
      return;
    }
    if (node.type === Fragment) return visit((node.props as { children?: ReactNode }).children);
    const marker = (node.type as { pptxReviewPart?: string }).pptxReviewPart;
    if (marker) found[marker] = node;
    else extra.push(node);
  };
  visit(children);
  if (extra.length > 0) found.__extra = extra;
  return found;
}

/** Shown when the slide has no thread. @public */
export const ReviewEmpty = markPart(function ReviewEmpty({
  className,
  hidden,
  children,
}: ReviewPartProps) {
  const t = useReviewLabel();
  const { review } = usePane();
  if (hidden) return null;
  return (
    <div className={cx('pptx-review__empty', className)}>
      {children ?? (
        <>
          <p className="pptx-review__empty-title">{t('comments.emptyTitle')}</p>
          {review.readOnly ? null : (
            <p className="pptx-review__empty-body">{t('comments.emptyBody')}</p>
          )}
        </>
      )}
    </div>
  );
}, 'Empty');

/**
 * One thread. A press on it, or Enter or Space while it has the focus, picks the thread: its
 * marker on the slide shows the pick, and the card opens its reply line. @public
 */
export const ReviewCard = markPart(function ReviewCard({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const { review, cardClassName } = usePane();
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  if (hidden || !entry) return null;
  const pick = (): void => {
    if (!entry.isActive) review.setActive(entry.key);
  };
  const shared = {
    className: cx('pptx-review__card', cardClassName, className),
    'data-comment-id': entry.id,
    ...(entry.isActive ? { 'data-active': '', 'aria-current': 'true' as const } : {}),
    ...(entry.resolved ? { 'data-resolved': '' } : {}),
    ...authorHooks(entry.author),
    style: authorStyle(entry.author),
    role: 'group' as const,
    'aria-label': t('comments.thread', { author: entry.author || t('comments.unknownAuthor') }),
    tabIndex: 0,
    onClick: (e: MouseEvent) => {
      if (!isCardControl(e.target)) pick();
    },
    onKeyDown: (e: KeyboardEvent) => {
      if (e.target !== e.currentTarget || (e.key !== 'Enter' && e.key !== ' ')) return;
      e.preventDefault();
      pick();
    },
  };
  if (asChild) return slotted(children, shared);
  return (
    <div {...shared}>
      <ReviewCardPreset>{children}</ReviewCardPreset>
    </div>
  );
}, 'Card');

/** The packaged card. A part passed as a child replaces the packaged copy of it. */
function ReviewCardPreset({ children }: { children?: ReactNode }) {
  const overrides = useMemo(() => partOverrides(children), [children]);
  const take = (key: string, fallback: ReactNode): ReactNode =>
    key in overrides ? overrides[key] : fallback;
  return (
    <>
      <div className="pptx-review__head">
        {take('Avatar', <ReviewAvatar />)}
        <div className="pptx-review__meta">
          {take('Author', <ReviewAuthor />)}
          {take('Time', <ReviewTime />)}
        </div>
        <div className="pptx-review__actions">
          {take('Resolve', <ReviewResolve />)}
          {take('Reopen', <ReviewReopen />)}
          {take('Edit', <ReviewEdit />)}
          {take('Delete', <ReviewDelete />)}
        </div>
      </div>
      {take('Anchor', <ReviewAnchor />)}
      {take('Summary', <ReviewSummary />)}
      {take('Replies', <ReviewReplies />)}
      {take('Reply', <ReviewReply />)}
      {overrides.__extra}
    </>
  );
}

/** The author's initials, in the author's color. @public */
export const ReviewAvatar = markPart(function ReviewAvatar({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const entry = useContext(ReviewItemContext);
  if (hidden || !entry) return null;
  const shared = { className: cx('pptx-review__avatar', className), 'aria-hidden': true };
  if (asChild) return slotted(children, shared);
  return <span {...shared}>{children ?? entry.initials}</span>;
}, 'Avatar');

/** The author's name, or "Unknown author". @public */
export const ReviewAuthor = markPart(function ReviewAuthor({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  if (hidden || !entry) return null;
  const shared = { className: cx('pptx-review__author', className) };
  if (asChild) return slotted(children, shared);
  return <span {...shared}>{children ?? (entry.author || t('comments.unknownAuthor'))}</span>;
}, 'Author');

/** When the comment was written: month, day, and time. Nothing for a comment with no date. @public */
export const ReviewTime = markPart(function ReviewTime({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  if (hidden || !entry) return null;
  const text = commentDateText(entry.date);
  const status = entry.resolved && entry.parentId === undefined ? t('comments.resolved') : '';
  if (!text && !status) return null;
  const shared = { className: cx('pptx-review__time', className) };
  if (asChild) return slotted(children, shared);
  return (
    <span {...shared}>
      {children ?? (
        <>
          {text ? <time dateTime={entry.date}>{text}</time> : null}
          {status ? <span className="pptx-review__status">{status}</span> : null}
        </>
      )}
    </span>
  );
}, 'Time');

/** The element that the thread points at, as "On Title 1". Nothing for the slide or a place. @public */
export const ReviewAnchor = markPart(function ReviewAnchor({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  if (hidden || !entry || entry.parentId !== undefined || entry.anchorName === undefined)
    return null;
  const shared = { className: cx('pptx-review__chip', className) };
  if (asChild) return slotted(children, shared);
  return <span {...shared}>{children ?? t('comments.on', { name: entry.anchorName })}</span>;
}, 'Anchor');

/** The comment's text, or its edit box while the comment is open for an edit. @public */
export const ReviewSummary = markPart(function ReviewSummary({
  className,
  asChild,
  hidden,
  children,
}: ReviewPartProps) {
  const entry = useContext(ReviewItemContext);
  const { editing } = usePane();
  if (hidden || !entry) return null;
  if (editing === entry.identity) return <ReviewEditBox item={entry} />;
  const shared = { className: cx('pptx-review__summary', className) };
  if (asChild) return slotted(children, shared);
  return <div {...shared}>{children ?? <p className="pptx-review__text">{entry.text}</p>}</div>;
}, 'Summary');

/** The replies of a thread, oldest first, each with its own author color and actions. @public */
export const ReviewReplies = markPart(function ReviewReplies({
  className,
  hidden,
}: ReviewPartProps) {
  const { byId } = usePane();
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  if (hidden || !entry || entry.replyIds.length === 0) return null;
  const replies = entry.replyIds
    .map((id) => byId.get(id))
    .filter((reply): reply is ReviewItemView => reply !== undefined);
  if (replies.length === 0) return null;
  return (
    <ol
      className={cx('pptx-review__replies', className)}
      aria-label={t('comments.replies', { author: entry.author || t('comments.unknownAuthor') })}
    >
      {replies.map((reply) => (
        <ReviewItemContext.Provider key={reply.identity} value={reply}>
          <li
            className="pptx-review__reply"
            data-reply-id={reply.id}
            {...authorHooks(reply.author)}
            style={authorStyle(reply.author)}
          >
            <div className="pptx-review__head">
              <ReviewAvatar />
              <div className="pptx-review__meta">
                <ReviewAuthor />
                <ReviewTime />
              </div>
              <div className="pptx-review__actions">
                <ReviewEdit />
                <ReviewDelete />
              </div>
            </div>
            <ReviewSummary />
          </li>
        </ReviewItemContext.Provider>
      ))}
    </ol>
  );
}, 'Replies');

interface ActionSpec {
  readonly part: string;
  readonly icon: string;
  /** The visible tooltip and the accessible name, with the author. */
  readonly label: (
    item: ReviewItemView
  ) => 'comments.resolveOf' | 'comments.reopenOf' | 'comments.editOf' | 'comments.deleteOf';
  /** Whether the card shows the action for this item. */
  readonly shows: (item: ReviewItemView) => boolean;
  readonly run: (pane: ReturnType<typeof usePane>, item: ReviewItemView) => void;
}

/** One action part: a 26 px icon button in the card head, as DOCX Editor's. */
function actionPart(spec: ActionSpec) {
  function ReviewAction({ className, asChild, hidden, children, icon }: ReviewActionProps) {
    const pane = usePane();
    const entry = useContext(ReviewItemContext);
    const t = useReviewLabel();
    if (hidden || !entry || pane.review.readOnly || !spec.shows(entry)) return null;
    const label = t(spec.label(entry), { author: entry.author || t('comments.unknownAuthor') });
    const shared = {
      type: 'button' as const,
      className: cx('pptx-review__action', className),
      'data-action': spec.part.toLowerCase(),
      'aria-label': label,
      'data-pptx-tooltip': label,
      onClick: (e: MouseEvent) => {
        e.stopPropagation();
        spec.run(pane, entry);
      },
    };
    if (asChild)
      return (
        <ReviewActionSlot engineDisabled={false} disabledReason={null} slotProps={shared}>
          {children}
        </ReviewActionSlot>
      );
    return <button {...shared}>{icon ?? children ?? <Icon name={spec.icon} />}</button>;
  }
  return markPart(ReviewAction, spec.part);
}

/** Resolve the thread. Shown on an open thread of the threaded format. @public */
export const ReviewResolve = actionPart({
  part: 'Resolve',
  icon: 'check',
  label: () => 'comments.resolveOf',
  shows: (item) => item.parentId === undefined && !item.resolved && item.resolvable,
  run: ({ review }, item) => review.resolve(item),
});

/** Reopen the resolved thread. @public */
export const ReviewReopen = actionPart({
  part: 'Reopen',
  icon: 'refresh',
  label: () => 'comments.reopenOf',
  shows: (item) => item.parentId === undefined && item.resolved && item.resolvable,
  run: ({ review }, item) => review.reopen(item),
});

/** Open the comment's text for an edit. Shown on hover and focus, where the editor can write it. @public */
export const ReviewEdit = actionPart({
  part: 'Edit',
  icon: 'edit',
  label: () => 'comments.editOf',
  shows: (item) => item.editable,
  run: ({ setEditing }, item) => setEditing(item.identity),
});

/** Delete the reply, or the thread with its replies. Shown on hover and focus. @public */
export const ReviewDelete = actionPart({
  part: 'Delete',
  icon: 'delete',
  label: () => 'comments.deleteOf',
  shows: () => true,
  run: ({ review, focusCard }, item) => {
    // The focus goes to the thread of a deleted reply, or past a deleted thread.
    if (review.remove(item)) focusCard(item.parentId ?? item.key);
  },
});
