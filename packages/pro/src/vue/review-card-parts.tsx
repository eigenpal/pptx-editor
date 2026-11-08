/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The card and its parts: the twin of `../react/review-card-parts.tsx`, with the same markup,
// classes, and ARIA, after DOCX Editor's Vue review card
// (docx-editor/packages/pro/src/vue/review-card-parts.tsx at adf0e178c). Comment text and names
// come from the file and render as text only.
/** @jsxImportSource vue */

import {
  Fragment,
  cloneVNode,
  defineComponent,
  isVNode,
  type PropType,
  type VNode,
  type VNodeChild,
} from 'vue';
import { Icon } from '@pptx-editor.dev/vue';
import {
  authorHooks,
  authorStyle,
  commentDateText,
  isCardControl,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { ReviewItemScope, usePane, useReviewItem, useReviewLabel } from './review-context.ts';
import { ReviewActionSlot } from './review-action-slot.tsx';
import { ReviewEditBox, ReviewReply } from './review-compose.tsx';

/** Mark a part, so the card and the pane replace a packaged part in place. */
export function markPart<T extends object>(component: T, part: string): T {
  return Object.assign(component, { pptxReviewPart: part });
}

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

const partProps = {
  className: { type: String, default: undefined },
  asChild: Boolean,
  hidden: Boolean,
} as const;

/** `asChild`: the part's props on its single slot element. */
function slotted(children: VNode[] | undefined, shared: Record<string, unknown>): VNodeChild {
  const child = children?.length === 1 ? children[0] : undefined;
  return child ? cloneVNode(child, shared) : null;
}

/** Map each slot part to its marker, so the card replaces the packaged copy in place. */
export function partOverrides(children: VNode[] | undefined): Record<string, VNode | VNode[]> {
  const found: Record<string, VNode | VNode[]> = {};
  const extra: VNode[] = [];
  const visit = (nodes: VNode[] | undefined): void => {
    for (const node of nodes ?? []) {
      if (!isVNode(node)) continue;
      if (node.type === Fragment) {
        visit(node.children as VNode[]);
        continue;
      }
      const marker = (node.type as { pptxReviewPart?: string }).pptxReviewPart;
      if (marker) found[marker] = node;
      else extra.push(node);
    }
  };
  visit(children);
  if (extra.length > 0) found.__extra = extra;
  return found;
}

/** Shown when the slide has no thread. @public */
export const ReviewEmpty = markPart(
  defineComponent({
    name: 'ReviewEmpty',
    props: partProps,
    setup(props, { slots }) {
      const t = useReviewLabel();
      const { review } = usePane();
      return () =>
        props.hidden ? null : (
          <div class={cx('pptx-review__empty', props.className)}>
            {slots.default?.() ?? [
              <p class="pptx-review__empty-title">{t('comments.emptyTitle')}</p>,
              review.readOnly.value ? null : (
                <p class="pptx-review__empty-body">{t('comments.emptyBody')}</p>
              ),
            ]}
          </div>
        );
    },
  }),
  'Empty'
);

/** One thread; a press, or Enter or Space while it has the focus, picks it. @public */
export const ReviewCard = markPart(
  defineComponent({
    name: 'ReviewCard',
    props: partProps,
    setup(props, { slots }) {
      const pane = usePane();
      const entry = useReviewItem();
      const t = useReviewLabel();
      return () => {
        const item = entry.value;
        if (props.hidden || !item) return null;
        const pick = (): void => {
          if (!item.isActive) pane.review.setActive(item.key, { identity: item.identity });
        };
        const shared = {
          class: cx('pptx-review__card', pane.cardClassName(), props.className),
          'data-comment-id': item.id,
          ...(item.isActive ? { 'data-active': '', 'aria-current': 'true' as const } : {}),
          ...(item.resolved ? { 'data-resolved': '' } : {}),
          ...authorHooks(item.author, pane.authors.value),
          style: authorStyle(item.author, pane.authors.value),
          role: 'group' as const,
          'aria-label': t('comments.thread', {
            author: item.author || t('comments.unknownAuthor'),
          }),
          tabindex: 0,
          onClick: (e: MouseEvent) => {
            if (!isCardControl(e.target)) pick();
          },
          onKeydown: (e: KeyboardEvent) => {
            if (e.target !== e.currentTarget || (e.key !== 'Enter' && e.key !== ' ')) return;
            e.preventDefault();
            pick();
          },
        };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return <div {...shared}>{cardPreset(slots.default?.())}</div>;
      };
    },
  }),
  'Card'
);

/** The packaged card. A part passed in the slot replaces the packaged copy of it. */
function cardPreset(children: VNode[] | undefined): VNodeChild {
  const overrides = partOverrides(children);
  const take = (key: string, fallback: VNodeChild): VNodeChild =>
    key in overrides ? overrides[key] : fallback;
  return [
    <div class="pptx-review__head">
      {take('Avatar', <ReviewAvatar />)}
      <div class="pptx-review__meta">
        {take('Author', <ReviewAuthor />)}
        {take('Time', <ReviewTime />)}
      </div>
      <div class="pptx-review__actions">
        {take('Resolve', <ReviewResolve />)}
        {take('Reopen', <ReviewReopen />)}
        {take('Edit', <ReviewEdit />)}
        {take('Delete', <ReviewDelete />)}
      </div>
    </div>,
    take('Anchor', <ReviewAnchor />),
    take('Summary', <ReviewSummary />),
    take('Replies', <ReviewReplies />),
    take('Reply', <ReviewReply />),
    overrides.__extra,
  ];
}

/** The author's initials, in the author's color. @public */
export const ReviewAvatar = markPart(
  defineComponent({
    name: 'ReviewAvatar',
    props: partProps,
    setup(props, { slots }) {
      const entry = useReviewItem();
      return () => {
        const item = entry.value;
        if (props.hidden || !item) return null;
        const shared = {
          class: cx('pptx-review__avatar', props.className),
          'aria-hidden': true as const,
        };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return <span {...shared}>{slots.default?.() ?? item.initials}</span>;
      };
    },
  }),
  'Avatar'
);

/** The author's name, or "Unknown author". @public */
export const ReviewAuthor = markPart(
  defineComponent({
    name: 'ReviewAuthor',
    props: partProps,
    setup(props, { slots }) {
      const entry = useReviewItem();
      const t = useReviewLabel();
      return () => {
        const item = entry.value;
        if (props.hidden || !item) return null;
        const shared = { class: cx('pptx-review__author', props.className) };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return (
          <span {...shared}>
            {slots.default?.() ?? (item.author || t('comments.unknownAuthor'))}
          </span>
        );
      };
    },
  }),
  'Author'
);

/** When the comment was written: month, day, and time. @public */
export const ReviewTime = markPart(
  defineComponent({
    name: 'ReviewTime',
    props: partProps,
    setup(props, { slots }) {
      const entry = useReviewItem();
      const t = useReviewLabel();
      return () => {
        const item = entry.value;
        if (props.hidden || !item) return null;
        const text = commentDateText(item.date);
        const status = item.resolved && item.parentId === undefined ? t('comments.resolved') : '';
        if (!text && !status) return null;
        const shared = { class: cx('pptx-review__time', props.className) };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return (
          <span {...shared}>
            {slots.default?.() ?? [
              text ? <time datetime={item.date}>{text}</time> : null,
              status ? <span class="pptx-review__status">{status}</span> : null,
            ]}
          </span>
        );
      };
    },
  }),
  'Time'
);

/** The element that the thread points at, as "On Title 1". @public */
export const ReviewAnchor = markPart(
  defineComponent({
    name: 'ReviewAnchor',
    props: partProps,
    setup(props, { slots }) {
      const entry = useReviewItem();
      const t = useReviewLabel();
      return () => {
        const item = entry.value;
        if (props.hidden || !item || item.parentId !== undefined || item.anchorName === undefined)
          return null;
        const shared = { class: cx('pptx-review__chip', props.className) };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return (
          <span {...shared}>
            {slots.default?.() ?? t('comments.on', { name: item.anchorName })}
          </span>
        );
      };
    },
  }),
  'Anchor'
);

/** The comment's text, or its edit box while the comment is open for an edit. @public */
export const ReviewSummary = markPart(
  defineComponent({
    name: 'ReviewSummary',
    props: partProps,
    setup(props, { slots }) {
      const entry = useReviewItem();
      const { editing } = usePane();
      return () => {
        const item = entry.value;
        if (props.hidden || !item) return null;
        if (editing.value === item.identity)
          return <ReviewEditBox item={item} key={item.identity} />;
        const shared = { class: cx('pptx-review__summary', props.className) };
        if (props.asChild) return slotted(slots.default?.(), shared);
        return (
          <div {...shared}>{slots.default?.() ?? <p class="pptx-review__text">{item.text}</p>}</div>
        );
      };
    },
  }),
  'Summary'
);

/** The replies of a thread, oldest first, each with its own author color and actions. @public */
export const ReviewReplies = markPart(
  defineComponent({
    name: 'ReviewReplies',
    props: partProps,
    setup(props) {
      const { byId, authors } = usePane();
      const entry = useReviewItem();
      const t = useReviewLabel();
      return () => {
        const item = entry.value;
        if (props.hidden || !item || item.replyIds.length === 0) return null;
        const replies = item.replyIds
          .map((id) => byId.value.get(id))
          .filter((reply): reply is ReviewItemView => reply !== undefined);
        if (replies.length === 0) return null;
        return (
          <ol
            class={cx('pptx-review__replies', props.className)}
            aria-label={t('comments.replies', {
              author: item.author || t('comments.unknownAuthor'),
            })}
          >
            {replies.map((reply) => (
              <ReviewItemScope key={reply.identity} item={reply}>
                {{
                  default: () => (
                    <li
                      class="pptx-review__reply"
                      data-reply-id={reply.id}
                      {...authorHooks(reply.author, authors.value)}
                      style={authorStyle(reply.author, authors.value)}
                    >
                      <div class="pptx-review__head">
                        <ReviewAvatar />
                        <div class="pptx-review__meta">
                          <ReviewAuthor />
                          <ReviewTime />
                        </div>
                        <div class="pptx-review__actions">
                          <ReviewEdit />
                          <ReviewDelete />
                        </div>
                      </div>
                      <ReviewSummary />
                    </li>
                  ),
                }}
              </ReviewItemScope>
            ))}
          </ol>
        );
      };
    },
  }),
  'Replies'
);

interface ActionSpec {
  readonly part: string;
  readonly icon: string;
  readonly label:
    'comments.resolveOf' | 'comments.reopenOf' | 'comments.editOf' | 'comments.deleteOf';
  readonly shows: (item: ReviewItemView) => boolean;
  readonly run: (pane: ReturnType<typeof usePane>, item: ReviewItemView) => void;
}

/** One action part: a 26 px icon button in the card head, as DOCX Editor's. */
function actionPart(spec: ActionSpec) {
  return markPart(
    defineComponent({
      name: `Review${spec.part}`,
      props: {
        ...partProps,
        icon: { type: [Object, String] as PropType<VNode | string>, default: undefined },
      },
      setup(props, { slots }) {
        const pane = usePane();
        const entry = useReviewItem();
        const t = useReviewLabel();
        return () => {
          const item = entry.value;
          if (props.hidden || !item || pane.review.readOnly.value || !spec.shows(item)) return null;
          const label = t(spec.label, { author: item.author || t('comments.unknownAuthor') });
          const shared = {
            type: 'button' as const,
            class: cx('pptx-review__action', props.className),
            'data-action': spec.part.toLowerCase(),
            'aria-label': label,
            'data-pptx-tooltip': label,
            onClick: (e: MouseEvent) => {
              e.stopPropagation();
              spec.run(pane, item);
            },
          };
          if (props.asChild)
            return (
              <ReviewActionSlot engineDisabled={false} disabledReason={null} slotProps={shared}>
                {{ default: () => slots.default?.() }}
              </ReviewActionSlot>
            );
          return (
            <button {...shared}>
              {props.icon ?? slots.default?.() ?? <Icon name={spec.icon} />}
            </button>
          );
        };
      },
    }),
    spec.part
  );
}

/** Resolve the thread. Shown on an open thread of the threaded format. @public */
export const ReviewResolve = actionPart({
  part: 'Resolve',
  icon: 'check',
  label: 'comments.resolveOf',
  shows: (item) => item.parentId === undefined && !item.resolved && item.resolvable,
  run: ({ review }, item) => review.resolve(item),
});

/** Reopen the resolved thread. @public */
export const ReviewReopen = actionPart({
  part: 'Reopen',
  icon: 'refresh',
  label: 'comments.reopenOf',
  shows: (item) => item.parentId === undefined && item.resolved && item.resolvable,
  run: ({ review }, item) => review.reopen(item),
});

/** Open the comment's text for an edit, where the editor can write it. @public */
export const ReviewEdit = actionPart({
  part: 'Edit',
  icon: 'edit',
  label: 'comments.editOf',
  shows: (item) => item.editable,
  run: ({ setEditing }, item) => setEditing(item.identity),
});

/** Delete the reply, or the thread with its replies. @public */
export const ReviewDelete = actionPart({
  part: 'Delete',
  icon: 'delete',
  label: 'comments.deleteOf',
  shows: () => true,
  run: ({ review, focusCard }, item) => {
    // The focus goes to the thread of a deleted reply, or past a deleted thread.
    if (review.remove(item)) focusCard(item.parentId ?? item.key);
  },
});
