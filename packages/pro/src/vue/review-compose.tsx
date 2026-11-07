/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comment fields: the twin of `../react/review-compose.tsx`, after DOCX Editor's Vue compose
// boxes (docx-editor/packages/pro/src/vue/review-compose-boxes.tsx at adf0e178c). Enter posts,
// Shift+Enter starts a new line, and Escape cancels. A refused write keeps the text.
/** @jsxImportSource vue */

import { defineComponent, onMounted, ref, useId, watch, type PropType } from 'vue';
import { Icon } from '@pptx-editor.dev/vue';
import { commentInitials } from '@pptx-editor.dev/core/editor';
import {
  COMMENT_TEXT_MAX,
  authorHooks,
  authorStyle,
  commentFieldKey,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { usePane, useReviewItem, useReviewLabel } from './review-context.ts';
import { ReviewActionSlot } from './review-action-slot.tsx';

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

/** A comment field: a text area that grows with its text, with the key rules above. */
const CommentField = defineComponent({
  name: 'PptxReviewCommentField',
  props: {
    id: { type: String, required: true },
    value: { type: String, required: true },
    label: { type: String, required: true },
    placeholder: { type: String, default: undefined },
    describedBy: { type: String, required: true },
    autoFocus: Boolean,
    onChange: { type: Function as PropType<(value: string) => void>, required: true },
    onPost: { type: Function as PropType<() => void>, required: true },
    onCancel: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const field = ref<HTMLTextAreaElement | null>(null);
    onMounted(() => {
      const el = field.value;
      if (!props.autoFocus || !el) return;
      el.focus({ preventScroll: true });
      el.setSelectionRange(el.value.length, el.value.length);
    });
    return () => (
      <textarea
        ref={field}
        id={props.id}
        class="pptx-review__input"
        data-comment-field=""
        rows={1}
        value={props.value}
        maxlength={COMMENT_TEXT_MAX}
        aria-label={props.label}
        aria-describedby={props.describedBy}
        placeholder={props.placeholder ?? props.label}
        onInput={(e: Event) => props.onChange((e.target as HTMLTextAreaElement).value)}
        onClick={(e: MouseEvent) => e.stopPropagation()}
        onKeydown={(e: KeyboardEvent) => {
          const key = commentFieldKey(e);
          if (!key) return;
          e.preventDefault();
          e.stopPropagation();
          if (key === 'post') props.onPost();
          else props.onCancel();
        }}
      />
    );
  },
});

function keyHint(id: string, t: ReturnType<typeof useReviewLabel>) {
  return (
    <span id={id} class="pptx-sr-only">
      {t('comments.keyHint')}
    </span>
  );
}

/** The **New comment** button. A root part: it opens the compose box. @public */
export const ReviewAddComment = Object.assign(
  defineComponent({
    name: 'ReviewAddComment',
    props: { className: { type: String, default: undefined }, asChild: Boolean, hidden: Boolean },
    setup(props, { slots }) {
      const { beginDraft, drafting, review } = usePane();
      const t = useReviewLabel();
      return () => {
        if (props.hidden) return null;
        const disabled = drafting.value || review.readOnly.value;
        if (props.asChild)
          return (
            <div class="pptx-review__bar">
              <ReviewActionSlot
                engineDisabled={disabled}
                disabledReason={null}
                slotProps={{
                  class: cx('pptx-review__add', props.className),
                  'data-review-add': '',
                  onClick: beginDraft,
                }}
              >
                {{ default: () => slots.default?.() }}
              </ReviewActionSlot>
            </div>
          );
        return (
          <div class="pptx-review__bar">
            <button
              type="button"
              class={cx('pptx-review__add', props.className)}
              data-review-add=""
              disabled={disabled}
              onClick={beginDraft}
            >
              <Icon name="add_comment" />
              <span>{slots.default?.() ?? t('comments.new')}</span>
            </button>
          </div>
        );
      };
    },
  }),
  { pptxReviewPart: 'AddComment' }
);

/** The compose box of a new comment, while a draft is open. @public */
export const ReviewDraft = Object.assign(
  defineComponent({
    name: 'ReviewDraft',
    props: {
      className: { type: String, default: undefined },
      asChild: Boolean,
      hidden: Boolean,
      nodeId: { type: String, default: undefined },
    },
    setup(props) {
      const { drafting, draftTarget, endDraft, review } = usePane();
      const t = useReviewLabel();
      const text = ref('');
      const id = useId()!;
      return () => {
        if (props.hidden || !drafting.value) return null;
        const author = review.author.value ?? '';
        // `nodeId`, else the element of a draft from a menu, else the selected element.
        const pinned = props.nodeId ?? draftTarget.value?.nodeId;
        const target = props.nodeId ? null : (draftTarget.value ?? review.commentTarget.value);
        const blocked = review.author.value === null;
        const empty = text.value.trim() === '';
        const post = (): void => {
          if (empty || blocked || review.readOnly.value) return;
          const at = pinned ? { nodeId: pinned } : undefined;
          if (review.comment(text.value, undefined, at)) {
            text.value = '';
            endDraft();
          }
        };
        const cancel = (): void => {
          text.value = '';
          endDraft();
        };
        return (
          <div
            class={cx('pptx-review__card', props.className)}
            data-draft=""
            {...authorHooks(author)}
            style={authorStyle(author)}
          >
            <form
              class="pptx-review__compose"
              onSubmit={(e: Event) => {
                e.preventDefault();
                post();
              }}
            >
              {target ? (
                <span class="pptx-review__chip">
                  {t('comments.pointsAt', { name: target.name })}
                </span>
              ) : null}
              <CommentField
                id={id}
                value={text.value}
                label={t('comments.addComment')}
                describedBy={blocked ? `${id}-blocked` : `${id}-hint`}
                autoFocus
                onChange={(v: string) => (text.value = v)}
                onPost={post}
                onCancel={cancel}
              />
              {keyHint(`${id}-hint`, t)}
              {blocked ? (
                <span id={`${id}-blocked`} class="pptx-review__refused">
                  {t('comments.noAuthor')}
                </span>
              ) : null}
              <div class="pptx-review__compose-actions">
                <button type="button" class="pptx-review__text-button" onClick={cancel}>
                  {t('comments.cancel')}
                </button>
                <button type="submit" class="pptx-review__submit" disabled={empty || blocked}>
                  {t('comments.text')}
                </button>
              </div>
            </form>
          </div>
        );
      };
    },
  }),
  { pptxReviewPart: 'Draft' }
);

/** The reply line of the picked thread. Not on a resolved thread or in view mode. @public */
export const ReviewReply = Object.assign(
  defineComponent({
    name: 'ReviewReply',
    props: { className: { type: String, default: undefined }, asChild: Boolean, hidden: Boolean },
    setup(props, { slots }) {
      const { review } = usePane();
      const entry = useReviewItem();
      const t = useReviewLabel();
      // Held here, so the draft stays while the reply line shows.
      const text = ref('');
      const id = useId()!;
      // Another comment in this card (the one answered is gone) drops the reply meant for it. An
      // edit or a resolve of the thread keeps the reply.
      watch(
        () => entry.value?.identity ?? '',
        () => (text.value = '')
      );
      return () => {
        const item = entry.value;
        if (props.hidden || !item || !item.isActive || item.resolved || review.readOnly.value)
          return null;
        const custom = slots.default?.();
        if (custom) return custom;
        const author = review.author.value ?? '';
        const name = item.author || t('comments.unknownAuthor');
        const empty = text.value.trim() === '';
        const blocked = review.author.value === null;
        const post = (): void => {
          if (!empty && review.reply(item, text.value)) text.value = '';
        };
        return (
          <form
            class={cx('pptx-review__reply-line', props.className)}
            onSubmit={(e: Event) => {
              e.preventDefault();
              post();
            }}
          >
            {author ? (
              <span
                class="pptx-review__avatar pptx-review__reply-avatar"
                aria-hidden="true"
                {...authorHooks(author)}
                style={authorStyle(author)}
              >
                {commentInitials({ initials: '', author })}
              </span>
            ) : null}
            <CommentField
              id={id}
              value={text.value}
              label={t('comments.replyTo', { author: name })}
              placeholder={t('comments.replyText')}
              describedBy={blocked ? `${id}-blocked` : `${id}-hint`}
              onChange={(v: string) => (text.value = v)}
              onPost={post}
              onCancel={() => {
                if (text.value) text.value = '';
                else review.setActive(null);
              }}
            />
            {keyHint(`${id}-hint`, t)}
            {text.value ? (
              <button
                type="button"
                class="pptx-review__text-button"
                onClick={() => (text.value = '')}
              >
                {t('comments.cancel')}
              </button>
            ) : null}
            <button
              type="submit"
              class="pptx-review__send"
              aria-label={t('comments.replyTo', { author: name })}
              data-pptx-tooltip={t('comments.reply')}
              disabled={empty || blocked}
            >
              <Icon name="send" />
            </button>
            {blocked ? (
              <span id={`${id}-blocked`} class="pptx-review__refused">
                {t('comments.noAuthor')}
              </span>
            ) : null}
          </form>
        );
      };
    },
  }),
  { pptxReviewPart: 'Reply' }
);

/** The edit box of a comment: its text in a field, with Cancel and Save. */
export const ReviewEditBox = defineComponent({
  name: 'PptxReviewEditBox',
  props: { item: { type: Object as PropType<ReviewItemView>, required: true } },
  setup(props) {
    const { review, setEditing, focusCard } = usePane();
    const t = useReviewLabel();
    const text = ref(props.item.text);
    const id = useId()!;
    // Another comment in this card (the one opened is gone): close and save nothing. An edit of
    // the same comment by a peer keeps the box and its text.
    const opened = props.item.identity;
    const moved = (): boolean => props.item.identity !== opened;
    watch(moved, (gone) => {
      if (gone) setEditing(null);
    });
    const close = (): void => {
      setEditing(null);
      focusCard(props.item.parentId ?? props.item.key);
    };
    const save = (): void => {
      if (!moved() && text.value.trim() !== '' && review.edit(props.item, text.value)) close();
    };
    return () => (
      <form
        class="pptx-review__compose"
        onSubmit={(e: Event) => {
          e.preventDefault();
          save();
        }}
      >
        <CommentField
          id={id}
          value={text.value}
          label={t('comments.editText')}
          describedBy={`${id}-hint`}
          autoFocus
          onChange={(v: string) => (text.value = v)}
          onPost={save}
          onCancel={close}
        />
        {keyHint(`${id}-hint`, t)}
        <div class="pptx-review__compose-actions">
          <button type="button" class="pptx-review__text-button" onClick={close}>
            {t('comments.cancel')}
          </button>
          <button type="submit" class="pptx-review__submit" disabled={text.value.trim() === ''}>
            {t('comments.save')}
          </button>
        </div>
      </form>
    );
  },
});
