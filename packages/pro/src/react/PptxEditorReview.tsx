/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `PptxEditorReview`: the comments pane of the current slide, as a compound component, after
// DOCX Editor's `DocxEditorReview` (docx-editor/packages/pro/src/react/DocxEditorReview.tsx at
// adf0e178c). PresentationML has no tracked-change markup and a slide has no text margin, so the
// pane is a list in the side panel, not a rail beside the page. The customization ladder is
// DOCX Editor's:
//
//   1. `className`, `card.className`, and `data-*`   restyle the packaged parts with CSS
//   2. `icon`                                          change one action's icon
//   3. `asChild`                                       put a part's wiring on your own element
//   4. a part as a child                              replaces the packaged part; `hidden`
//                                                      removes it; `preset={false}` drops them
//   5. `useReview()`                                   the data and the actions, for your own UI
//
// The Vue twin is `../vue/PptxEditorReview.tsx`.
/** @jsxImportSource react */

import {
  Fragment,
  cloneElement,
  isValidElement,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import { registerCommentDraft } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from '@pptx-editor.dev/react';
import {
  commentTarget,
  liveDraftTarget,
  reviewReadOnly,
  type ReviewCommentTarget,
  type ReviewItemView,
} from '../review/comment-view.ts';
import { useReview } from './useReview.ts';
import {
  ReviewContext,
  ReviewItemContext,
  usePane,
  useReviewLabel,
  type ReviewPaneValue,
} from './review-context.ts';
import type { ReviewPartProps, ReviewProps } from './review-types.ts';
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

/** Props of `PptxEditorReview.List`. @public */
export interface ReviewListProps extends Omit<ReviewPartProps, 'children' | 'asChild'> {
  /** The card template (a `Card`, or parts that replace the packaged ones), or a render callback. */
  children?: ReactNode | ((item: ReviewItemView) => ReactNode);
}

const cx = (...names: (string | false | undefined)[]): string => names.filter(Boolean).join(' ');

/** The threads of the slide, one card each, or `Empty` when there is none. @public */
export const ReviewList = markPart(function ReviewList({
  className,
  hidden,
  children,
}: ReviewListProps) {
  const { review } = usePane();
  const t = useReviewLabel();
  const parts = useMemo(
    () => (typeof children === 'function' ? null : partOverrides(children)),
    [children]
  );
  if (hidden) return null;
  const roots = review.items.filter((item) => item.parentId === undefined);
  if (roots.length === 0) {
    if (typeof children === 'function') return null;
    return <>{parts && 'Empty' in parts ? parts.Empty : <ReviewEmpty />}</>;
  }
  const card = (): ReactNode => {
    if (parts && isValidElement(parts.Card)) return parts.Card;
    const rest = parts ? Object.entries(parts).filter(([key]) => key !== 'Empty') : [];
    return (
      <ReviewCard>
        {rest.map(([key, node]) => (
          <Fragment key={key}>{node}</Fragment>
        ))}
      </ReviewCard>
    );
  };
  return (
    <ul className={cx('pptx-review__list', className)} aria-label={t('comments.label')}>
      {roots.map((item) => (
        <ReviewItemContext.Provider key={item.identity} value={item}>
          <li className="pptx-review__slot">
            {typeof children === 'function' ? children(item) : card()}
          </li>
        </ReviewItemContext.Provider>
      ))}
    </ul>
  );
}, 'List');

const ROOT_PARTS = new Set(['List', 'AddComment', 'Draft']);

/** Split the root's children: root parts by name, and the card template for the list. */
function rootChildren(children: ReviewProps['children']): {
  readonly parts: Record<string, ReactNode>;
  readonly template: ReviewListProps['children'];
} {
  if (typeof children === 'function') return { parts: {}, template: children };
  const parts: Record<string, ReactNode> = {};
  const rest: ReactNode[] = [];
  const visit = (node: ReactNode): void => {
    if (Array.isArray(node)) return void node.forEach(visit);
    if (isValidElement(node) && node.type === Fragment)
      return visit((node.props as { children?: ReactNode }).children);
    const marker = isValidElement(node)
      ? (node.type as { pptxReviewPart?: string }).pptxReviewPart
      : undefined;
    if (marker && ROOT_PARTS.has(marker)) parts[marker] = node;
    else if (node !== null && node !== undefined && node !== false) rest.push(node);
  };
  visit(children);
  return { parts, template: rest.length > 0 ? rest : undefined };
}

/**
 * The comments pane. Render it in the editor root, for example as the **Comments** tab of the
 * side panel (`<PptxEditor.SidePanel panes={{ comments: <PptxEditorReview /> }} />`). It needs
 * the review module (`reviewModule()`); without one, it says that comments are kept when you
 * save. @public
 */
function ReviewRoot({ className, hidden, t, card, children, preset = true, filter }: ReviewProps) {
  const review = useReview();
  const editor = usePptxEditor();
  const label = useReviewLabelFrom(t);
  const noModule = useEditorState((s) => s.status === 'ready' && s.comments == null);
  const [drafting, setDrafting] = useState(false);
  // The element a draft from a menu points at: the one the right-click selected. A deleted
  // element no longer counts: the draft then points at the selection or the slide.
  const [pinnedTarget, setDraftTarget] = useState<ReviewCommentTarget | null>(null);
  const shapes = useEditorState((s) => s.shapes);
  const draftTarget = liveDraftTarget(pinnedTarget, shapes);
  const targetGone = pinnedTarget !== null && draftTarget === null;
  useEffect(() => {
    if (targetGone) setDraftTarget(null);
  }, [targetGone]);
  const [editing, setEditing] = useState<string | null>(null);
  const root = useRef<HTMLElement>(null);

  // `filter` picks threads; a reply follows its thread.
  const items = useMemo(() => {
    if (!filter) return review.items;
    const kept = new Set(
      review.items.filter((i) => i.parentId === undefined && filter(i)).map((i) => i.id)
    );
    return review.items.filter((i) => kept.has(i.parentId ?? i.id));
  }, [review.items, filter]);
  const byId = useMemo(() => new Map(items.map((i) => [i.id, i] as const)), [items]);
  // A new slide, or view mode, closes the open fields.
  const slide = useEditorState((s) => s.slides[s.currentSlide]?.id ?? '');
  useEffect(() => {
    setDrafting(false);
    setEditing(null);
  }, [slide, review.readOnly]);
  // The canvas menus' New comment row opens a draft here, on the selected element.
  // A hidden pane opens no draft: the row then asks a pane that shows.
  useEffect(() => {
    if (!editor || noModule || hidden) return undefined;
    return registerCommentDraft(editor, () => {
      const snap = editor.snapshot();
      if (reviewReadOnly(snap)) return;
      setDraftTarget(commentTarget(snap.selection, snap.shapes));
      setDrafting(true);
      // A draft that was already open takes the focus again.
      requestAnimationFrame(() =>
        root.current?.querySelector<HTMLElement>('[data-draft] [data-comment-field]')?.focus()
      );
    });
  }, [editor, noModule, hidden]);

  const focusCard = useCallback((key: string) => {
    requestAnimationFrame(() => {
      const el = root.current;
      const card = el?.querySelector<HTMLElement>(`[data-comment-id="${CSS.escape(key)}"]`);
      (card ?? el?.querySelector<HTMLElement>('[data-review-add]'))?.focus();
    });
  }, []);
  const beginDraft = useCallback(() => {
    setDraftTarget(null);
    setDrafting(true);
  }, []);
  const endDraft = useCallback(() => {
    setDrafting(false);
    setDraftTarget(null);
    requestAnimationFrame(() =>
      root.current?.querySelector<HTMLElement>('[data-review-add]')?.focus()
    );
  }, []);

  // A thread picked elsewhere (its marker on the slide) scrolls into view.
  useEffect(() => {
    if (!review.activeKey) return;
    root.current
      ?.querySelector<HTMLElement>(`[data-comment-id="${CSS.escape(review.activeKey)}"]`)
      ?.scrollIntoView?.({ block: 'nearest' });
  }, [review.activeKey]);

  const value = useMemo<ReviewPaneValue>(
    () => ({
      t,
      cardClassName: card?.className,
      review: items === review.items ? review : { ...review, items },
      byId,
      drafting,
      draftTarget,
      beginDraft,
      endDraft,
      editing,
      setEditing,
      focusCard,
    }),
    [
      t,
      card?.className,
      review,
      items,
      byId,
      drafting,
      draftTarget,
      beginDraft,
      endDraft,
      editing,
      focusCard,
    ]
  );

  if (hidden) return null;
  if (noModule)
    return (
      <div className={cx('pptx-empty', className)}>
        <p className="pptx-empty__title">{label('panel.comments.emptyTitle')}</p>
        <p className="pptx-empty__body">{label('panel.comments.emptyBody')}</p>
      </div>
    );
  const { parts, template } = rootChildren(children);
  const take = (key: string, fallback: ReactElement): ReactNode =>
    key in parts ? parts[key] : preset ? fallback : null;
  const list = parts.List;
  return (
    <ReviewContext.Provider value={value}>
      <section
        ref={root}
        className={cx('pptx-review', className)}
        aria-label={label('panel.tabs.comments')}
        data-count={items.filter((i) => i.parentId === undefined).length}
      >
        {take('AddComment', <ReviewAddComment />)}
        {take('Draft', <ReviewDraft />)}
        {list !== undefined ? (
          isValidElement<ReviewListProps>(list) && list.props.children === undefined && template ? (
            cloneElement(list, { children: template })
          ) : (
            list
          )
        ) : preset ? (
          <ReviewList>{template}</ReviewList>
        ) : (
          (template as ReactNode)
        )}
      </section>
    </ReviewContext.Provider>
  );
}

/** The root's own label resolver, before its context exists. */
function useReviewLabelFrom(host: ReviewProps['t']) {
  const fallback = useReviewLabel();
  return (key: Parameters<typeof fallback>[0]): string => host?.(key) ?? fallback(key);
}

/** The comments pane compound. @public */
export interface PptxEditorReviewNamespace {
  (props: ReviewProps): ReturnType<typeof ReviewRoot>;
  /** The list of the slide's threads. A root part. */
  readonly List: typeof ReviewList;
  /** The **New comment** button. A root part. */
  readonly AddComment: typeof ReviewAddComment;
  /** The compose box of a new comment, while a draft is open. A root part. */
  readonly Draft: typeof ReviewDraft;
  readonly Empty: typeof ReviewEmpty;
  readonly Card: typeof ReviewCard;
  readonly Avatar: typeof ReviewAvatar;
  readonly Author: typeof ReviewAuthor;
  readonly Time: typeof ReviewTime;
  /** The element that the thread points at. */
  readonly Anchor: typeof ReviewAnchor;
  readonly Summary: typeof ReviewSummary;
  readonly Resolve: typeof ReviewResolve;
  readonly Reopen: typeof ReviewReopen;
  readonly Edit: typeof ReviewEdit;
  readonly Delete: typeof ReviewDelete;
  readonly Replies: typeof ReviewReplies;
  /** The reply line of the picked thread. */
  readonly Reply: typeof ReviewReply;
}

/**
 * The comments pane as a compound component: `PptxEditorReview` is the root, and every part hangs
 * off it, as DOCX Editor's `DocxEditorReview`. Root parts (`List`, `AddComment`, `Draft`) go as
 * direct children of the root. Card parts go in `List`, or as other direct children, which
 * become the card template.
 *
 * @example
 * ```tsx
 * <PptxEditorReview>
 *   <PptxEditorReview.List>
 *     <PptxEditorReview.Card>
 *       <PptxEditorReview.Delete hidden />
 *     </PptxEditorReview.Card>
 *   </PptxEditorReview.List>
 * </PptxEditorReview>
 * ```
 *
 * @public
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
});
