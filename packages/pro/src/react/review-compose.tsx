/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comment fields, after DOCX Editor's compose boxes
// (docx-editor/packages/pro/src/react/review-compose-boxes.tsx at adf0e178c): the compose card
// of a new comment, the reply line of the picked thread, and the edit box of a comment. Enter
// posts, Shift+Enter starts a new line, and Escape cancels. A refused write keeps the text. The
// Vue twin is `../vue/review-compose.tsx`.
/** @jsxImportSource react */

import { useContext, useEffect, useId, useRef, useState } from 'react';
import { Icon } from '@pptx-editor.dev/react';
import {
  COMMENT_TEXT_MAX,
  authorHooks,
  authorStyle,
  commentFieldKey,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { commentInitials } from '@pptx-editor.dev/core/editor';
import { ReviewItemContext, usePane, useReviewLabel } from './review-context.ts';
import type { ReviewDraftProps, ReviewPartProps } from './review-types.ts';
import { ReviewActionSlot } from './review-action-slot.tsx';

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

/** A comment field: a text area that grows with its text, with the key rules above. */
function CommentField(props: {
  readonly id: string;
  readonly value: string;
  readonly label: string;
  readonly placeholder?: string;
  readonly describedBy: string;
  readonly autoFocus?: boolean;
  readonly onChange: (value: string) => void;
  readonly onPost: () => void;
  readonly onCancel: () => void;
}) {
  const field = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    const el = field.current;
    if (!props.autoFocus || !el) return;
    el.focus({ preventScroll: true });
    el.setSelectionRange(el.value.length, el.value.length);
    // Mount only: the field takes the focus when it opens.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return (
    <textarea
      ref={field}
      id={props.id}
      className="pptx-review__input"
      data-comment-field=""
      rows={1}
      value={props.value}
      maxLength={COMMENT_TEXT_MAX}
      aria-label={props.label}
      aria-describedby={props.describedBy}
      placeholder={props.placeholder ?? props.label}
      onChange={(e) => props.onChange(e.target.value)}
      onClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        const key = commentFieldKey({ ...e, isComposing: e.nativeEvent.isComposing });
        if (!key) return;
        e.preventDefault();
        e.stopPropagation();
        if (key === 'post') props.onPost();
        else props.onCancel();
      }}
    />
  );
}

/** The key hint of a field, for screen readers. */
function KeyHint({ id }: { readonly id: string }) {
  const t = useReviewLabel();
  return (
    <span id={id} className="pptx-sr-only">
      {t('comments.keyHint')}
    </span>
  );
}

/**
 * The **New comment** button. A root part: it opens the compose box (`Draft`). Off in view mode
 * and while the compose box is open. @public
 */
export function ReviewAddComment({ className, asChild, hidden, children }: ReviewPartProps) {
  const { beginDraft, drafting, review } = usePane();
  const t = useReviewLabel();
  if (hidden) return null;
  const label = t('comments.new');
  const disabled = drafting || review.readOnly;
  if (asChild) {
    return (
      <div className="pptx-review__bar">
        <ReviewActionSlot
          engineDisabled={disabled}
          disabledReason={null}
          slotProps={{
            className: cx('pptx-review__add', className),
            'data-review-add': '',
            onClick: beginDraft,
          }}
        >
          {children}
        </ReviewActionSlot>
      </div>
    );
  }
  return (
    <div className="pptx-review__bar">
      <button
        type="button"
        className={cx('pptx-review__add', className)}
        data-review-add=""
        disabled={disabled}
        onClick={beginDraft}
      >
        <Icon name="add_comment" />
        <span>{children ?? label}</span>
      </button>
    </div>
  );
}
ReviewAddComment.pptxReviewPart = 'AddComment' as const;

/**
 * The compose box of a new comment, in the signing author's color. It renders only while a draft
 * is open (`AddComment`). The comment points at `nodeId`, else at the selected element, else at
 * the slide; the box names the element. Without an author, it says so and Comment is off. @public
 */
export function ReviewDraft({ className, hidden, nodeId }: ReviewDraftProps) {
  const { drafting, draftTarget, endDraft, review } = usePane();
  const t = useReviewLabel();
  const [text, setText] = useState('');
  const id = useId();
  if (hidden || !drafting) return null;
  const author = review.author ?? '';
  // `nodeId`, else the element of a draft from a menu, else the selected element.
  const pinned = nodeId ?? draftTarget?.nodeId;
  const target = nodeId ? null : (draftTarget ?? review.commentTarget);
  const blocked = review.author === null;
  const empty = text.trim() === '';
  const post = (): void => {
    if (empty || blocked || review.readOnly) return;
    if (review.comment(text, undefined, pinned ? { nodeId: pinned } : undefined)) {
      setText('');
      endDraft();
    }
  };
  const cancel = (): void => {
    setText('');
    endDraft();
  };
  return (
    <div
      className={cx('pptx-review__card', className)}
      data-draft=""
      {...authorHooks(author)}
      style={authorStyle(author)}
    >
      <form
        className="pptx-review__compose"
        onSubmit={(e) => {
          e.preventDefault();
          post();
        }}
      >
        {target ? (
          <span className="pptx-review__chip">{t('comments.pointsAt', { name: target.name })}</span>
        ) : null}
        <CommentField
          id={id}
          value={text}
          label={t('comments.addComment')}
          describedBy={blocked ? `${id}-blocked` : `${id}-hint`}
          autoFocus
          onChange={setText}
          onPost={post}
          onCancel={cancel}
        />
        <KeyHint id={`${id}-hint`} />
        {blocked ? (
          <span id={`${id}-blocked`} className="pptx-review__refused">
            {t('comments.noAuthor')}
          </span>
        ) : null}
        <div className="pptx-review__compose-actions">
          <button type="button" className="pptx-review__text-button" onClick={cancel}>
            {t('comments.cancel')}
          </button>
          <button type="submit" className="pptx-review__submit" disabled={empty || blocked}>
            {t('comments.text')}
          </button>
        </div>
      </form>
    </div>
  );
}
ReviewDraft.pptxReviewPart = 'Draft' as const;

/**
 * The reply line of the picked thread: the signing author's avatar, a field, Cancel while it holds
 * text, and a round send button. Not on a resolved thread or in view mode. @public
 */
export function ReviewReply({ className, hidden, children }: ReviewPartProps) {
  const { review } = usePane();
  const entry = useContext(ReviewItemContext);
  const t = useReviewLabel();
  // Held here, so the draft stays while the reply line shows.
  const [text, setText] = useState('');
  const id = useId();
  // Another comment in this card (the one answered is gone) drops the reply meant for it. An edit
  // or a resolve of the thread keeps the reply.
  const who = entry?.identity ?? '';
  const [openedFor, setOpenedFor] = useState(who);
  if (who !== openedFor) {
    setOpenedFor(who);
    setText('');
  }
  if (hidden || !entry || !entry.isActive || entry.resolved || review.readOnly) return null;
  if (children) return <>{children}</>;
  const author = review.author ?? '';
  const name = entry.author || t('comments.unknownAuthor');
  const empty = text.trim() === '';
  const post = (): void => {
    if (!empty && review.reply(entry, text)) setText('');
  };
  return (
    <form
      className={cx('pptx-review__reply-line', className)}
      onSubmit={(e) => {
        e.preventDefault();
        post();
      }}
    >
      {author ? (
        <span
          className="pptx-review__avatar pptx-review__reply-avatar"
          aria-hidden="true"
          {...authorHooks(author)}
          style={authorStyle(author)}
        >
          {commentInitials({ initials: '', author })}
        </span>
      ) : null}
      <CommentField
        id={id}
        value={text}
        label={t('comments.replyTo', { author: name })}
        placeholder={t('comments.replyText')}
        describedBy={review.author === null ? `${id}-blocked` : `${id}-hint`}
        onChange={setText}
        onPost={post}
        onCancel={() => {
          if (text) setText('');
          else review.setActive(null);
        }}
      />
      <KeyHint id={`${id}-hint`} />
      {text ? (
        <button type="button" className="pptx-review__text-button" onClick={() => setText('')}>
          {t('comments.cancel')}
        </button>
      ) : null}
      <button
        type="submit"
        className="pptx-review__send"
        aria-label={t('comments.replyTo', { author: name })}
        data-pptx-tooltip={t('comments.reply')}
        disabled={empty || review.author === null}
      >
        <Icon name="send" />
      </button>
      {review.author === null ? (
        <span id={`${id}-blocked`} className="pptx-review__refused">
          {t('comments.noAuthor')}
        </span>
      ) : null}
    </form>
  );
}
ReviewReply.pptxReviewPart = 'Reply' as const;

/** The edit box of a comment: its text in a field, with Cancel and Save. */
export function ReviewEditBox({ item }: { readonly item: ReviewItemView }) {
  const { review, setEditing, focusCard } = usePane();
  const t = useReviewLabel();
  const [text, setText] = useState(item.text);
  const id = useId();
  // Another comment in this card (the one opened is gone): close and save nothing. An edit of the
  // same comment by a peer keeps the box and its text.
  const opened = useRef(item.identity);
  const moved = item.identity !== opened.current;
  useEffect(() => {
    if (moved) setEditing(null);
  }, [moved, setEditing]);
  const close = (): void => {
    setEditing(null);
    focusCard(item.parentId ?? item.key);
  };
  const save = (): void => {
    if (!moved && text.trim() !== '' && review.edit(item, text)) close();
  };
  return (
    <form
      className="pptx-review__compose"
      onSubmit={(e) => {
        e.preventDefault();
        save();
      }}
    >
      <CommentField
        id={id}
        value={text}
        label={t('comments.editText')}
        describedBy={`${id}-hint`}
        autoFocus
        onChange={setText}
        onPost={save}
        onCancel={close}
      />
      <KeyHint id={`${id}-hint`} />
      <div className="pptx-review__compose-actions">
        <button type="button" className="pptx-review__text-button" onClick={close}>
          {t('comments.cancel')}
        </button>
        <button type="submit" className="pptx-review__submit" disabled={text.trim() === ''}>
          {t('comments.save')}
        </button>
      </div>
    </form>
  );
}
