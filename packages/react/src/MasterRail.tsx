/** @jsxImportSource react */
// The slide master view's chrome: the rail of masters and layouts, and the bar over the canvas
// that names the template being edited and closes the view. `SlideRail` shows `MasterRail` in
// that view, so a host that composes its own workspace gets it too. The Vue twin is
// `packages/vue/src/MasterRail.tsx`.

import { useEffect, useRef, type KeyboardEvent } from 'react';
import {
  currentSlideSummary,
  focusSurfaceOf,
  keepCurrentSlideInView,
  masterBarHint,
  masterBarTitle,
  masterRailLabel,
  resolveRailOrientation,
  masterRailUsage,
  masterStepTarget,
  type SlideSummary,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { paneStyle } from './pane-state.ts';
import { Thumbnail, useFocusFollowsCurrent, type SlideRailProps } from './panes.tsx';

/**
 * The rail of the slide master view: each master, then its layouts, with their names and how
 * many slides use them. A click or the arrow keys show a template on the canvas; Enter moves to
 * the canvas. @internal
 */
export function MasterRail(props: SlideRailProps) {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const templates = useEditorState((s) => s.slides);
  const current = useEditorState((s) => s.currentSlide);
  const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
  const list = useRef<HTMLOListElement>(null);
  const trackFocus = useFocusFollowsCurrent(list, '.pptx-rail__item', current, templates.length);
  useEffect(() => keepCurrentSlideInView(list.current, '.pptx-rail__item'), [current]);
  const orientation = resolveRailOrientation(props.orientation, 'PptxEditor.SlideRail');
  const onKeyDown = (e: KeyboardEvent<HTMLOListElement>): void => {
    if (!editor) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      focusSurfaceOf(editor);
      return;
    }
    const next = masterStepTarget(e.key, templates, current);
    if (next === null) return;
    e.preventDefault();
    editor.goToSlide(next);
  };
  const className = ['pptx-rail', 'pptx-rail--master', props.className].filter(Boolean).join(' ');
  // A host that closes the rail (`open={false}`) keeps it closed in this view too.
  if (props.open === false) return null;
  return (
    <nav
      className={className}
      aria-label={t('masterView.title')}
      data-open="true"
      data-orientation={orientation}
      style={paneStyle(props.paneWidth, props.style)}
    >
      <ol
        ref={list}
        className="pptx-rail__list pptx-master-rail__list"
        role="listbox"
        aria-label={t('masterView.railLabel')}
        aria-orientation={orientation}
        onKeyDown={onKeyDown}
        onFocus={trackFocus}
      >
        {templates.map((s) => (
          <MasterRailItem
            key={s.id}
            template={s}
            ratio={ratio}
            current={s.index === current}
            onPick={() => editor?.goToSlide(s.index)}
          />
        ))}
      </ol>
    </nav>
  );
}

function MasterRailItem({
  template: s,
  ratio,
  current,
  onPick,
}: {
  template: SlideSummary;
  ratio: number;
  current: boolean;
  onPick: () => void;
}) {
  const { t } = useTranslation();
  const name = masterRailLabel(t, s);
  const usage = masterRailUsage(t, s);
  return (
    <li
      className="pptx-rail__item pptx-master-rail__item"
      role="option"
      aria-selected={current}
      aria-current={current}
      data-template={s.template}
      data-index={s.index}
      data-unused={s.usedBy ? undefined : ''}
      tabIndex={current ? 0 : -1}
      onClick={onPick}
    >
      <Thumbnail slide={s} ratio={ratio} />
      <span className="pptx-master-rail__name">{name}</span>
      <span className="pptx-visually-hidden">, </span>
      <span className="pptx-master-rail__usage">{usage}</span>
    </li>
  );
}

/**
 * The bar over the canvas in the slide master view: what is being edited, what the edit
 * changes, and **Close master**, which returns to the slide the view opened from.
 * `PptxEditor.Content` renders it, so every layout has it; it renders nothing in the other
 * views. @internal
 */
export function MasterViewBar() {
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const on = useEditorState((s) => s.presentationView === 'slideMaster');
  const current = useEditorState((s) => (on ? currentSlideSummary(s) : undefined));
  if (!on || !current) return null;
  const close = (): void => {
    if (!editor) return;
    editor.exec({ type: 'setPresentationView', view: 'normal' });
    // The bar leaves with focus in it: focus goes to the slide, not to the page.
    focusSurfaceOf(editor);
  };
  return (
    <div className="pptx-master-bar" role="region" aria-label={masterBarTitle(t, current)}>
      <Icon name="dashboard" className="pptx-icon pptx-master-bar__icon" />
      <div className="pptx-master-bar__text">
        <span className="pptx-master-bar__title" title={masterBarTitle(t, current)}>
          {masterBarTitle(t, current)}
        </span>
        <span className="pptx-master-bar__hint">{masterBarHint(t, current)}</span>
      </div>
      <button
        type="button"
        className="pptx-master-bar__close"
        onMouseDown={keepFocus}
        onClick={close}
      >
        {t('masterView.close')}
      </button>
    </div>
  );
}
