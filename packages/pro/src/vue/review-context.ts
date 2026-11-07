/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The pane's private injections and the two public composables that read them: the twin of
// `../react/review-context.ts`, after DOCX Editor's Vue review context
// (docx-editor/packages/pro/src/vue/review-context.ts at adf0e178c).

import {
  computed,
  defineComponent,
  inject,
  provide,
  shallowRef,
  toValue,
  type ComputedRef,
  type InjectionKey,
  type MaybeRefOrGetter,
  type PropType,
  type Ref,
} from 'vue';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { useTranslation } from '@pptx-editor.dev/vue';
import {
  reviewAuthorInfo,
  reviewActions,
  type ReviewAuthorInfo,
  type ReviewCommentTarget,
  type ReviewItemView,
} from '../review/comment-view.ts';
import type { UseReviewReturn } from './useReview.ts';
import type { ReviewTranslate } from './review-types.ts';

/** What the pane gives its parts. */
export interface ReviewPaneValue {
  readonly t: () => ReviewTranslate | undefined;
  readonly cardClassName: () => string | undefined;
  /** The review, with `items` narrowed by the pane's `filter`. */
  readonly review: UseReviewReturn;
  readonly byId: ComputedRef<ReadonlyMap<string, ReviewItemView>>;
  readonly drafting: Ref<boolean>;
  /**
   * The element a draft from the canvas menus' New comment row points at, or null: the draft then
   * points at the selected element.
   */
  readonly draftTarget: Readonly<Ref<ReviewCommentTarget | null>>;
  readonly beginDraft: () => void;
  readonly endDraft: () => void;
  readonly editing: Ref<string | null>;
  readonly setEditing: (key: string | null) => void;
  readonly focusCard: (key: string) => void;
}

export const reviewPaneKey: InjectionKey<ReviewPaneValue> = Symbol('pptxReviewPane');
const reviewItemKey: InjectionKey<ComputedRef<ReviewItemView | null>> = Symbol('pptxReviewItem');

const none = computed(() => null);

/** Provide `item` to the parts below, as the React `ReviewItemContext.Provider` does. */
export const ReviewItemScope = defineComponent({
  name: 'PptxReviewItemScope',
  props: { item: { type: Object as PropType<ReviewItemView>, required: true } },
  setup(props, { slots }) {
    provide(
      reviewItemKey,
      computed(() => props.item)
    );
    return () => slots.default?.();
  },
});

/** What `useReviewItem` returns. @public */
export type UseReviewItemReturn = ComputedRef<ReviewItemView | null>;

/**
 * The comment that the surrounding card (or reply) renders, or null outside one: the twin of
 * React's `useReviewItem`. @public
 */
export function useReviewItem(): UseReviewItemReturn {
  return inject(reviewItemKey, none);
}

/** What `useReviewAuthor` returns. @public */
export type UseReviewAuthorReturn = ComputedRef<ReviewAuthorInfo | undefined>;

/**
 * The look of one comment author: the color slot and the CSS color of the canvas marker and the
 * packaged cards, or `undefined` without an author. Takes a ref or a getter. @public
 */
export function useReviewAuthor(
  author: MaybeRefOrGetter<string | undefined>
): UseReviewAuthorReturn {
  return computed(() => {
    const name = toValue(author);
    return name ? reviewAuthorInfo(name) : undefined;
  });
}

const off = computed(() => true);
const INERT_PANE: ReviewPaneValue = {
  t: () => undefined,
  cardClassName: () => undefined,
  review: {
    items: computed(() => []),
    activeKey: computed(() => null),
    ...reviewActions(null),
    author: computed(() => null),
    commentTarget: computed(() => null),
    readOnly: off,
    ready: computed(() => false),
  },
  byId: computed(() => new Map()),
  drafting: shallowRef(false),
  draftTarget: shallowRef(null),
  beginDraft: () => {},
  endDraft: () => {},
  editing: shallowRef(null),
  setEditing: () => {},
  focusCard: () => {},
};

/** The pane's data. A part outside the pane renders nothing, not an error. */
export function usePane(): ReviewPaneValue {
  return inject(reviewPaneKey, INERT_PANE);
}

/** Placeholder values for a label. */
export type ReviewLabelParams = Readonly<Record<string, string | number>>;

/** A part's strings: the host's `t`, else the bundled catalog. */
export function useReviewLabel(): (key: TranslationKey, params?: ReviewLabelParams) => string {
  const pane = inject(reviewPaneKey, null);
  const { t } = useTranslation();
  return (key, params) => {
    const own = pane?.t()?.(key);
    if (own === undefined) return t(key, params);
    return params
      ? own.replace(/\{(\w+)\}/g, (whole, name: string) =>
          Object.hasOwn(params, name) ? String(params[name]) : whole
        )
      : own;
  };
}
