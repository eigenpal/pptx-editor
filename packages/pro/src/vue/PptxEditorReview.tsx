/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `PptxEditorReview`: the comments pane as a compound component, the twin of
// `../react/PptxEditorReview.tsx`, after DOCX Editor's Vue `DocxEditorReview`
// (docx-editor/packages/pro/src/vue/DocxEditorReview.tsx at adf0e178c). React's render callback
// child is the `#card` scoped slot (`{ item }`), on the root and on `List`.
/** @jsxImportSource vue */

import {
  Fragment,
  computed,
  defineComponent,
  isVNode,
  nextTick,
  onBeforeUnmount,
  provide,
  ref,
  shallowRef,
  watch,
  type DefineComponent,
  type PropType,
  type SlotsType,
  type VNode,
  type VNodeChild,
} from 'vue';
import { registerCommentDraft } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from '@pptx-editor.dev/vue';
import {
  commentTarget,
  liveDraftTarget,
  reviewReadOnly,
  type ReviewCommentTarget,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { useReview, type UseReviewReturn } from './useReview.ts';
import {
  ReviewItemScope,
  reviewPaneKey,
  usePane,
  useReviewLabel,
  type ReviewPaneValue,
} from './review-context.ts';
import type {
  ReviewActionProps,
  ReviewDraftProps,
  ReviewPartProps,
  ReviewProps,
  ReviewTranslate,
} from './review-types.ts';
import {
  ReviewAnchor,
  ReviewAuthor,
  ReviewAvatar,
  ReviewCard,
  ReviewDelete,
  ReviewEdit,
  ReviewEmpty,
  ReviewReopen,
  ReviewReplies,
  ReviewResolve,
  ReviewSummary,
  ReviewTime,
  markPart,
  partOverrides,
} from './review-card-parts.tsx';
import { ReviewAddComment, ReviewDraft, ReviewReply } from './review-compose.tsx';

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

/** The slots of `PptxEditorReview` and `PptxEditorReview.List`. */
type ReviewSlots = SlotsType<{
  default?: () => VNode[];
  card?: (scope: { item: ReviewItemView }) => VNode[];
}>;

/** The list's card for each thread: the `#card` slot, a `Card` part, or the packaged card. */
function listBody(
  roots: readonly ReviewItemView[],
  template: VNode[] | undefined,
  card: ((scope: { item: ReviewItemView }) => VNode[]) | undefined
): VNodeChild {
  const parts = partOverrides(template);
  return roots.map((item) => (
    <ReviewItemScope key={item.key} item={item}>
      {{
        default: () => (
          <li class="pptx-review__slot">
            {card ? (
              card({ item })
            ) : 'Card' in parts ? (
              parts.Card
            ) : (
              <ReviewCard>
                {{
                  default: () =>
                    Object.entries(parts)
                      .filter(([key]) => key !== 'Empty')
                      .flatMap(([, node]) => node),
                }}
              </ReviewCard>
            )}
          </li>
        ),
      }}
    </ReviewItemScope>
  ));
}

/** The threads of the slide, one card each, or `Empty` when there is none. @public */
export const ReviewList = markPart(
  defineComponent({
    name: 'ReviewList',
    props: { className: { type: String, default: undefined }, hidden: Boolean },
    slots: Object as ReviewSlots,
    setup(props, { slots }) {
      const { review } = usePane();
      const t = useReviewLabel();
      return () => {
        if (props.hidden) return null;
        const roots = review.items.value.filter((item) => item.parentId === undefined);
        const template = slots.default?.();
        if (roots.length === 0) {
          if (slots.card) return null;
          const parts = partOverrides(template);
          return 'Empty' in parts ? parts.Empty : <ReviewEmpty />;
        }
        return (
          <ul class={cx('pptx-review__list', props.className)} aria-label={t('comments.label')}>
            {listBody(roots, template, slots.card)}
          </ul>
        );
      };
    },
  }),
  'List'
);

const ROOT_PARTS = new Set(['List', 'AddComment', 'Draft']);

/** Split the root's slot: root parts by name, and the card template for the list. */
function rootChildren(children: VNode[] | undefined): {
  readonly parts: Record<string, VNode>;
  readonly template: VNode[];
} {
  const parts: Record<string, VNode> = {};
  const template: VNode[] = [];
  const visit = (nodes: VNode[] | undefined): void => {
    for (const node of nodes ?? []) {
      if (!isVNode(node)) continue;
      if (node.type === Fragment) {
        visit(node.children as VNode[]);
        continue;
      }
      const marker = (node.type as { pptxReviewPart?: string }).pptxReviewPart;
      if (marker && ROOT_PARTS.has(marker)) parts[marker] = node;
      else template.push(node);
    }
  };
  visit(children);
  return { parts, template };
}

const reviewProps = {
  /** A class name added to the pane. */
  className: { type: String, default: undefined },
  /** Render nothing. */
  hidden: Boolean,
  /** Label resolver. A key it does not resolve falls back to the bundled catalog. */
  t: { type: Function as PropType<ReviewTranslate>, default: undefined },
  /** A class name for each card. */
  card: { type: Object as PropType<{ className?: string }>, default: undefined },
  /** Render the packaged arrangement. `false` renders only the root parts you pass. */
  preset: { type: Boolean, default: true },
  /** Show only some threads. Replies follow their thread. */
  filter: { type: Function as PropType<(item: ReviewItemView) => boolean>, default: undefined },
} as const;

/**
 * The comments pane: the twin of React's `PptxEditorReview`. Render it in `PptxEditorRoot`, for
 * example in the `#comments` slot of `PptxEditorSidePanel`. It needs the review module
 * (`reviewModule()`); without one, it says that comments are kept when you save. @public
 */
const ReviewRoot = defineComponent({
  name: 'PptxEditorReview',
  props: reviewProps,
  slots: Object as ReviewSlots,
  setup(props, { slots }) {
    const base = useReview();
    const t = useReviewLabel();
    const label = (key: Parameters<typeof t>[0]): string => props.t?.(key) ?? t(key);
    const noModule = useEditorState((s) => s.status === 'ready' && s.comments == null);
    const slide = useEditorState((s) => s.slides[s.currentSlide]?.id ?? '');
    const drafting = ref(false);
    // The element a draft from a menu points at: the one the right-click selected. A deleted
    // element no longer counts: the draft then points at the selection or the slide.
    const pinnedTarget = shallowRef<ReviewCommentTarget | null>(null);
    const shapes = useEditorState((s) => s.shapes);
    const draftTarget = computed(() => liveDraftTarget(pinnedTarget.value, shapes.value));
    watch(draftTarget, (live) => {
      if (live === null) pinnedTarget.value = null;
    });
    const editing = shallowRef<string | null>(null);
    const root = ref<HTMLElement | null>(null);
    const editor = usePptxEditor();

    // `filter` picks threads; a reply follows its thread.
    const items = computed(() => {
      const all = base.items.value;
      const filter = props.filter;
      if (!filter) return all;
      const kept = new Set(
        all.filter((i) => i.parentId === undefined && filter(i)).map((i) => i.id)
      );
      return all.filter((i) => kept.has(i.parentId ?? i.id));
    });
    const review: UseReviewReturn = { ...base, items };
    // A new slide, or view mode, closes the open fields.
    watch([slide, base.readOnly], () => {
      drafting.value = false;
      editing.value = null;
    });
    // The canvas menus' New comment row opens a draft here, on the selected element.
    let unregister: (() => void) | null = null;
    watch(
      [editor, noModule, () => props.hidden],
      ([e, off, hidden]) => {
        unregister?.();
        unregister = null;
        // A hidden pane opens no draft: the row then asks a pane that shows.
        if (!e || off || hidden) return;
        unregister = registerCommentDraft(e, () => {
          const snap = e.snapshot();
          if (reviewReadOnly(snap)) return;
          pinnedTarget.value = commentTarget(snap.selection, snap.shapes);
          drafting.value = true;
          // A draft that was already open takes the focus again.
          void nextTick(() =>
            requestAnimationFrame(() =>
              root.value?.querySelector<HTMLElement>('[data-draft] [data-comment-field]')?.focus()
            )
          );
        });
      },
      { immediate: true }
    );
    onBeforeUnmount(() => unregister?.());
    const focusCard = (key: string): void =>
      void nextTick(() =>
        requestAnimationFrame(() => {
          const el = root.value;
          const card = el?.querySelector<HTMLElement>(`[data-comment-id="${CSS.escape(key)}"]`);
          (card ?? el?.querySelector<HTMLElement>('[data-review-add]'))?.focus();
        })
      );
    // A thread picked elsewhere (its marker on the slide) scrolls into view.
    watch(base.activeKey, (key) => {
      if (!key) return;
      void nextTick(() =>
        root.value
          ?.querySelector<HTMLElement>(`[data-comment-id="${CSS.escape(key)}"]`)
          ?.scrollIntoView?.({ block: 'nearest' })
      );
    });
    const pane: ReviewPaneValue = {
      t: () => props.t,
      cardClassName: () => props.card?.className,
      review,
      byId: computed(() => new Map(items.value.map((i) => [i.id, i] as const))),
      drafting,
      draftTarget,
      beginDraft: () => {
        pinnedTarget.value = null;
        drafting.value = true;
      },
      endDraft: () => {
        drafting.value = false;
        pinnedTarget.value = null;
        void nextTick(() =>
          requestAnimationFrame(() =>
            root.value?.querySelector<HTMLElement>('[data-review-add]')?.focus()
          )
        );
      },
      editing,
      setEditing: (key) => (editing.value = key),
      focusCard,
    };
    provide(reviewPaneKey, pane);
    return () => {
      if (props.hidden) return null;
      if (noModule.value)
        return (
          <div class={cx('pptx-empty', props.className)}>
            <p class="pptx-empty__title">{label('panel.comments.emptyTitle')}</p>
            <p class="pptx-empty__body">{label('panel.comments.emptyBody')}</p>
          </div>
        );
      const { parts, template } = rootChildren(slots.default?.());
      const take = (key: string, fallback: () => VNodeChild): VNodeChild =>
        key in parts ? parts[key] : props.preset ? fallback() : null;
      const list = parts.List;
      return (
        <section
          ref={root}
          class={cx('pptx-review', props.className)}
          aria-label={label('panel.tabs.comments')}
          data-count={items.value.filter((i) => i.parentId === undefined).length}
        >
          {take('AddComment', () => (
            <ReviewAddComment />
          ))}
          {take('Draft', () => (
            <ReviewDraft />
          ))}
          {list ? (
            list
          ) : props.preset ? (
            <ReviewList>
              {{
                default: () => template,
                ...(slots.card ? { card: slots.card } : {}),
              }}
            </ReviewList>
          ) : (
            template
          )}
        </section>
      );
    };
  },
});

/**
 * The type of the comments pane compound. It is written out, not inferred, so the declaration
 * files and the API report stay the same on every build, as `FontFamilyNamespace` in the Vue
 * adapter. @public
 */
export type PptxEditorReviewNamespace = DefineComponent<ReviewProps> & {
  /** The thread list. The `#card` slot renders each thread. */
  readonly List: DefineComponent<Omit<ReviewPartProps, 'asChild'>>;
  /** The button that opens the compose box of a new comment. */
  readonly AddComment: DefineComponent<ReviewPartProps>;
  /** The compose box of a new comment. */
  readonly Draft: DefineComponent<ReviewDraftProps>;
  /** The text shown when the slide has no threads. */
  readonly Empty: DefineComponent<ReviewPartProps>;
  /** One thread card. */
  readonly Card: DefineComponent<ReviewPartProps>;
  /** The author's initials in the author's color. */
  readonly Avatar: DefineComponent<ReviewPartProps>;
  /** The author's name. */
  readonly Author: DefineComponent<ReviewPartProps>;
  /** The date of the comment. */
  readonly Time: DefineComponent<ReviewPartProps>;
  /** The element that the thread points at. */
  readonly Anchor: DefineComponent<ReviewPartProps>;
  /** The comment text. */
  readonly Summary: DefineComponent<ReviewPartProps>;
  /** Resolve the thread. */
  readonly Resolve: DefineComponent<ReviewActionProps>;
  /** Reopen a resolved thread. */
  readonly Reopen: DefineComponent<ReviewActionProps>;
  /** Edit the comment text. */
  readonly Edit: DefineComponent<ReviewActionProps>;
  /** Delete the comment or the thread. */
  readonly Delete: DefineComponent<ReviewActionProps>;
  /** The replies of the thread. */
  readonly Replies: DefineComponent<ReviewPartProps>;
  /** The reply box of the thread. */
  readonly Reply: DefineComponent<ReviewPartProps>;
};

/**
 * The comments pane compound: `PptxEditorReview` is the root, and every part hangs off it, as
 * DOCX Editor's `DocxEditorReview`. @public
 */
export const PptxEditorReview: PptxEditorReviewNamespace = Object.assign(ReviewRoot, {
  List: ReviewList,
  AddComment: ReviewAddComment,
  Draft: ReviewDraft,
  Empty: ReviewEmpty,
  Card: ReviewCard,
  Avatar: ReviewAvatar,
  Author: ReviewAuthor,
  Time: ReviewTime,
  Anchor: ReviewAnchor,
  Summary: ReviewSummary,
  Resolve: ReviewResolve,
  Reopen: ReviewReopen,
  Edit: ReviewEdit,
  Delete: ReviewDelete,
  Replies: ReviewReplies,
  Reply: ReviewReply,
}) as unknown as PptxEditorReviewNamespace;
