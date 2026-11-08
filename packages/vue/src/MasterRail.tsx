/** @jsxImportSource vue */
// The slide master view's chrome: the rail of masters and layouts, and the bar over the canvas
// that names the template being edited and closes the view. `PptxEditorSlideRail` shows
// `MasterRail` in that view, so a host that composes its own workspace gets it too. The React
// twin is `packages/react/src/MasterRail.tsx`.

import { defineComponent, ref, watch, type PropType, type StyleValue } from 'vue';
import {
  currentSlideSummary,
  focusSurfaceOf,
  keepCurrentSlideInView,
  masterBarHint,
  masterBarTitle,
  masterRailLabel,
  resolveRailOrientation,
  type SlideRailOrientation,
  masterRailUsage,
  masterStepTarget,
  type SlideSummary,
} from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { Thumbnail, paneStyle, useFocusFollowsCurrent } from './panes.tsx';

const sameSlides = (a: readonly SlideSummary[], b: readonly SlideSummary[]): boolean =>
  a.length === b.length && a.every((s, i) => s === b[i]);

/**
 * The rail of the slide master view: each master, then its layouts, with their names and how
 * many slides use them. A click or the arrow keys show a template on the canvas; Enter moves to
 * the canvas. @internal
 */
export const MasterRail = defineComponent({
  name: 'PptxMasterRail',
  props: {
    className: { type: String, default: undefined },
    paneWidth: { type: Number, default: undefined },
    orientation: { type: String as PropType<SlideRailOrientation>, default: undefined },
    railStyle: { type: [Object, String, Array] as PropType<StyleValue>, default: undefined },
  },
  setup(props) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const templates = useEditorState((s) => s.slides, sameSlides);
    const current = useEditorState((s) => s.currentSlide);
    const ratio = useEditorState((s) => s.slideHeight / s.slideWidth);
    const list = ref<HTMLOListElement | null>(null);
    const trackFocus = useFocusFollowsCurrent(list, '.pptx-rail__item', () => [
      current.value,
      templates.value.length,
    ]);
    watch(
      [list, current],
      ([el], _old, onCleanup) => onCleanup(keepCurrentSlideInView(el, '.pptx-rail__item')),
      { flush: 'post' }
    );
    const onKeydown = (e: KeyboardEvent): void => {
      const ed = editor.value;
      if (!ed) return;
      if (e.key === 'Enter') {
        e.preventDefault();
        focusSurfaceOf(ed);
        return;
      }
      const next = masterStepTarget(e.key, templates.value, current.value);
      if (next === null) return;
      e.preventDefault();
      ed.goToSlide(next);
    };
    return () => {
      const orientation = resolveRailOrientation(props.orientation, 'PptxEditorSlideRail');
      return (
        <nav
          class={['pptx-rail', 'pptx-rail--master', props.className]}
          aria-label={t('masterView.title')}
          data-open="true"
          data-orientation={orientation}
          style={paneStyle(props.paneWidth, props.railStyle)}
        >
          <ol
            ref={list}
            class="pptx-rail__list pptx-master-rail__list"
            role="listbox"
            aria-label={t('masterView.railLabel')}
            aria-orientation={orientation}
            onKeydown={onKeydown}
            onFocusin={trackFocus}
          >
            {templates.value.map((s) => {
              const on = s.index === current.value;
              const name = masterRailLabel(t, s);
              const usage = masterRailUsage(t, s);
              return (
                <li
                  key={s.id}
                  class="pptx-rail__item pptx-master-rail__item"
                  role="option"
                  aria-selected={on}
                  aria-current={on}
                  data-template={s.template}
                  data-index={s.index}
                  data-unused={s.usedBy ? undefined : ''}
                  tabindex={on ? 0 : -1}
                  onClick={() => editor.value?.goToSlide(s.index)}
                >
                  <Thumbnail slide={s} ratio={ratio.value} />
                  <span class="pptx-master-rail__name">{name}</span>
                  <span class="pptx-visually-hidden">, </span>
                  <span class="pptx-master-rail__usage">{usage}</span>
                </li>
              );
            })}
          </ol>
        </nav>
      );
    };
  },
});

/**
 * The bar over the canvas in the slide master view: what is being edited, what the edit
 * changes, and **Close master**, which returns to the slide the view opened from.
 * `PptxEditorContent` renders it, so every layout has it; it renders nothing in the other
 * views. @internal
 */
export const MasterViewBar = defineComponent({
  name: 'PptxMasterViewBar',
  setup() {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const current = useEditorState((s) =>
      s.presentationView === 'slideMaster' ? currentSlideSummary(s) : undefined
    );
    return () => {
      const s = current.value;
      if (!s) return null;
      const close = (): void => {
        const ed = editor.value;
        if (!ed) return;
        ed.exec({ type: 'setPresentationView', view: 'normal' });
        // The bar leaves with focus in it: focus goes to the slide, not to the page.
        focusSurfaceOf(ed);
      };
      return (
        <div class="pptx-master-bar" role="region" aria-label={masterBarTitle(t, s)}>
          <Icon name="dashboard" class="pptx-icon pptx-master-bar__icon" />
          <div class="pptx-master-bar__text">
            <span class="pptx-master-bar__title" title={masterBarTitle(t, s)}>
              {masterBarTitle(t, s)}
            </span>
            <span class="pptx-master-bar__hint">{masterBarHint(t, s)}</span>
          </div>
          <button
            type="button"
            class="pptx-master-bar__close"
            onMousedown={keepFocus}
            onClick={close}
          >
            {t('masterView.close')}
          </button>
        </div>
      );
    };
  },
});
