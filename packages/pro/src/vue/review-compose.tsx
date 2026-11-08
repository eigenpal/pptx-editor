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
import { Icon, usePptxEditor } from '@pptx-editor.dev/vue';
import { commentInitials } from '@pptx-editor.dev/core/editor';
import {
  COMMENT_TEXT_MAX,
  authorHooks,
  authorStyle,
  commentFieldKey,
  fieldNotice,
  fieldNoticeText,
  refusalOf,
  seenOf,
  shownAs,
  type CommentSeen,
  type FieldNotice,
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

/**
 * The notice of a field after a refusal. The live region stays mounted, so each refusal reads,
 * also the same one again.
 */
function noticeOf(
  id: string,
  notice: FieldNotice | null,
  t: ReturnType<typeof useReviewLabel>,
  current?: string
) {
  const changed = notice?.rejection.reason === 'errors.commentChanged';
  return (
    <span id={id} class="pptx-review__refused pptx-review__notice" role="status">
      {notice ? <span key={notice.serial}>{fieldNoticeText(t, notice)}</span> : null}
      {notice && changed && current !== undefined ? (
        <span class="pptx-review__current">{t('comments.nowReads', { text: current })}</span>
      ) : null}
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
      const { drafting, draftTarget, endDraft, review, authors } = usePane();
      const editor = usePptxEditor();
      const t = useReviewLabel();
      const text = ref('');
      const notice = ref<FieldNotice | null>(null);
      const id = useId()!;
      watch(drafting, (open) => {
        if (!open) notice.value = null;
      });
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
          // A draft from a menu keeps the text that was selected when it opened.
          const range = !props.nodeId ? draftTarget.value?.range : undefined;
          const at = pinned ? { nodeId: pinned, ...(range ? { range } : {}) } : undefined;
          const refused = refusalOf(editor.value, () => review.comment(text.value, undefined, at));
          if (refused === null) {
            text.value = '';
            notice.value = null;
            endDraft();
          } else notice.value = fieldNotice(refused, 'draft', notice.value);
        };
        const cancel = (): void => {
          text.value = '';
          notice.value = null;
          endDraft();
        };
        return (
          <div
            class={cx('pptx-review__card', props.className)}
            data-draft=""
            {...authorHooks(author, authors.value)}
            style={authorStyle(author, authors.value)}
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
                describedBy={`${blocked ? `${id}-blocked` : `${id}-hint`} ${id}-notice`}
                autoFocus
                onChange={(v: string) => {
                  text.value = v;
                  notice.value = null;
                }}
                onPost={post}
                onCancel={cancel}
              />
              {keyHint(`${id}-hint`, t)}
              {noticeOf(`${id}-notice`, notice.value, t)}
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
      const { review, authors } = usePane();
      const editor = usePptxEditor();
      const entry = useReviewItem();
      const t = useReviewLabel();
      // Held here, so the draft stays while the reply line shows.
      const text = ref('');
      const id = useId()!;
      // The thread as it showed when the draft began: the editor refuses the reply when the
      // comment under its id is no longer that one, and the draft stays with a notice.
      let seen: CommentSeen | null = null;
      const notice = ref<FieldNotice | null>(null);
      // Another comment in this card (the one answered is gone) drops the reply meant for it. An
      // edit or a resolve of the thread keeps the reply.
      watch(
        () => entry.value?.identity ?? '',
        () => {
          text.value = '';
          seen = null;
          notice.value = null;
        }
      );
      // A notice belongs to one visit of the reply line: it goes when the thread is deselected.
      watch(
        () => entry.value?.isActive ?? false,
        (active) => {
          if (!active) notice.value = null;
        }
      );
      const change = (value: string): void => {
        if (!value) seen = null;
        else if (!seen && entry.value) seen = seenOf(entry.value);
        text.value = value;
        notice.value = null;
      };
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
          if (empty || blocked) return;
          const refused = refusalOf(editor.value, () =>
            review.reply(seen ? shownAs(item, seen) : item, text.value)
          );
          if (refused === null) {
            text.value = '';
            seen = null;
            notice.value = null;
            return;
          }
          // The thread changed: the next post answers it as it shows now, by the user's choice.
          if (refused?.reason === 'errors.commentChanged') seen = seenOf(item);
          notice.value = fieldNotice(refused, 'reply', notice.value);
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
                {...authorHooks(author, authors.value)}
                style={authorStyle(author, authors.value)}
              >
                {commentInitials({ initials: '', author })}
              </span>
            ) : null}
            <CommentField
              id={id}
              value={text.value}
              label={t('comments.replyTo', { author: name })}
              placeholder={t('comments.replyText')}
              describedBy={`${blocked ? `${id}-blocked` : `${id}-hint`} ${id}-notice`}
              onChange={change}
              onPost={post}
              onCancel={() => {
                if (text.value) change('');
                else review.setActive(null);
              }}
            />
            {keyHint(`${id}-hint`, t)}
            {noticeOf(`${id}-notice`, notice.value, t)}
            {text.value ? (
              <button type="button" class="pptx-review__text-button" onClick={() => change('')}>
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
    const editor = usePptxEditor();
    const t = useReviewLabel();
    const text = ref(props.item.text);
    const id = useId()!;
    // Another comment in this card (the one opened is gone): close and save nothing. An edit of
    // the same comment by a peer keeps the box and its text.
    const opened = props.item.identity;
    const moved = (): boolean => props.item.identity !== opened;
    // The comment as it showed when the box opened: the editor refuses the save when the comment
    // under its id is no longer that one, and the box keeps the text with a notice.
    let seen = seenOf(props.item);
    const notice = ref<FieldNotice | null>(null);
    watch(moved, (gone) => {
      if (gone) setEditing(null);
    });
    const close = (): void => {
      setEditing(null);
      focusCard(props.item.parentId ?? props.item.key);
    };
    const save = (): void => {
      if (moved() || text.value.trim() === '') return;
      const refused = refusalOf(editor.value, () =>
        review.edit(shownAs(props.item, seen), text.value)
      );
      if (refused === null) return close();
      // The comment changed: the next save replaces it as it shows now, by the user's choice. A
      // twin that text cannot tell apart stays refused.
      if (refused?.reason === 'errors.commentChanged') seen = seenOf(props.item);
      notice.value = fieldNotice(refused, 'edit', notice.value);
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
          describedBy={`${id}-hint ${id}-notice`}
          autoFocus
          onChange={(v: string) => {
            text.value = v;
            notice.value = null;
          }}
          onPost={save}
          onCancel={close}
        />
        {keyHint(`${id}-hint`, t)}
        {noticeOf(`${id}-notice`, notice.value, t, props.item.text)}
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
