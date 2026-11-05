/** @jsxImportSource vue */
// The slide show: the Vue twin of the React adapter's `Slideshow.tsx`. Full screen (the
// Fullscreen API, or a window-covering overlay when the browser refuses), one slide painted by
// the engine's `paintSlide`, the usual show keys, and an auto-hiding control bar. Navigation
// state is the engine's `SlideshowController`; full screen, the start keys, the show keys, the
// screen size, and links are core's (`slideshow-screen.ts`), so the two adapters cannot drift.

import {
  computed,
  defineComponent,
  inject,
  onBeforeUnmount,
  onMounted,
  onUnmounted,
  provide,
  ref,
  shallowRef,
  watch,
  watchPostEffect,
  type InjectionKey,
  type PropType,
  type Ref,
  type ShallowRef,
} from 'vue';
import {
  SLIDESHOW_IDLE_MS,
  enterSlideshowScreen,
  fitSlideToScreen,
  listenSlideshowKeys,
  listenSlideshowStartKeys,
  observeSlideshowScreen,
  openSlideshowLink,
  paintSlideshowSlide,
  slideshowLinkUnder,
  startSlideshowController,
  translateKeyOrText,
  type PptxEditorInstance,
  type SlideshowController,
  type SlideshowScreenSize,
  type SlideshowState,
} from '@pptx-editor.dev/core/editor';
import { Icon } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';

/**
 * The slide show of the editor in `PptxEditorRoot`, the twin of React's `UseSlideshowReturn`. The
 * running show is a ref, so a template reads it reactively. @public
 */
export interface UseSlideshowReturn {
  /** The running show, or null. `PptxEditorSlideshow` paints it. */
  readonly controller: Readonly<ShallowRef<SlideshowController | null>>;
  /**
   * Start the show from the first shown slide or from the current slide. Does nothing without
   * an open document with slides. Viewing mode can present.
   */
  start(from: 'start' | 'current'): void;
  /** Close the show. */
  stop(): void;
}

/** The slide show published by `PptxEditorRoot`, the twin of React's `SlideshowContext`. @public */
export const pptxSlideshowKey: InjectionKey<UseSlideshowReturn> = Symbol('pptxSlideshow');

const noShow: UseSlideshowReturn = {
  controller: shallowRef(null),
  start: () => undefined,
  stop: () => undefined,
};

/**
 * Provide the slide show of one root and return it. The root passes its editor: an `inject` in
 * the providing component would read the parent's value. @internal
 */
export function provideSlideshow(
  editorRef: Readonly<Ref<PptxEditorInstance | null>>
): UseSlideshowReturn {
  const controller = shallowRef<SlideshowController | null>(null);
  const handle: UseSlideshowReturn = {
    controller,
    start(from) {
      const next = startSlideshowController(editorRef.value, from);
      if (next) controller.value = next;
    },
    stop: () => {
      controller.value = null;
    },
  };
  provide(pptxSlideshowKey, handle);
  return handle;
}

/** The slide show of the editor in `PptxEditorRoot`, as React's `useSlideshow`. @public */
export function useSlideshow(): UseSlideshowReturn {
  return inject(pptxSlideshowKey, noShow);
}

/**
 * One key per show and editor, so a new show, or a new editor under a running show, mounts a
 * new screen (full screen, focus, keys, and the deck subscription).
 */
const showKeys = new WeakMap<object, number>();
let nextShowKey = 0;
function keyOf(value: object): number {
  let key = showKeys.get(value);
  if (key === undefined) showKeys.set(value, (key = ++nextShowKey));
  return key;
}
function showKey(controller: SlideshowController, editor: PptxEditorInstance): string {
  return `${keyOf(controller)}:${keyOf(editor)}`;
}

/** A CSS length as React writes a number: unitless zero, pixels otherwise. */
const px = (n: number): string => (n === 0 ? '0' : `${n}px`);

const SlideshowScreen = defineComponent({
  name: 'PptxSlideshowScreen',
  props: {
    editor: { type: Object as PropType<PptxEditorInstance>, required: true as const },
    controller: { type: Object as PropType<SlideshowController>, required: true as const },
    onClose: { type: Function as PropType<() => void>, required: true as const },
  },
  setup(props) {
    const { t } = useTranslation();
    const { editor, controller } = props;
    const root = ref<HTMLDivElement | null>(null);
    const slide = ref<HTMLDivElement | null>(null);
    const state = shallowRef<SlideshowState>(controller.state);
    const offState = controller.subscribe(() => {
      state.value = controller.state;
    });
    const width = useEditorState((s) => s.slideWidth);
    const height = useEditorState((s) => s.slideHeight);
    const slides = useEditorState((s) => s.slides);
    const renderKey = computed(() => slides.value[state.value.index]?.key ?? '');
    const size = shallowRef<SlideshowScreenSize>({ w: 0, h: 0, dpr: 1 });
    const bar = ref(false);
    const link = ref(false);
    let idle: ReturnType<typeof setTimeout> | null = null;
    let overBar = false;

    // The deck can change under a running show (fonts load, a collaborator edits).
    const offDeck = editor.subscribe(() => controller.update(editor.snapshot().slides));

    // The show ended: the editor returns to the slide shown last.
    watch(
      () => [state.value.done, state.value.index] as const,
      ([done, index]) => {
        if (!done) return;
        editor.goToSlide(index);
        props.onClose();
      },
      { immediate: true }
    );

    // Full screen, the show keys, and the screen size (core's `slideshow-screen.ts`). They
    // leave before the show leaves the DOM, so focus returns in the same update and a start key
    // pressed right after the show ends is not swallowed.
    let stops: (() => void)[] = [];
    onMounted(() => {
      const el = root.value;
      if (!el) return;
      stops = [
        enterSlideshowScreen(el),
        listenSlideshowKeys(controller, () => root.value),
        observeSlideshowScreen(el, (next) => {
          size.value = next;
        }),
      ];
    });
    onBeforeUnmount(() => {
      for (const stop of stops) stop();
      stops = [];
    });
    onUnmounted(() => {
      offState();
      offDeck();
      if (idle) clearTimeout(idle);
    });

    // Paint the slide on show with the editor's own painter.
    watchPostEffect(() => {
      const el = slide.value;
      const { index, screen } = state.value;
      void renderKey.value;
      const label = (kind: string): string =>
        t('canvas.placeholderObject', { kind: translateKeyOrText(t, `canvas.kinds.${kind}`) });
      // Read the translator now, so a locale change repaints the frame labels.
      void t('slideshow.ariaLabel');
      const session = editor.session;
      if (!el) return;
      if (!session || screen === 'end') {
        el.replaceChildren();
        return;
      }
      const svg = paintSlideshowSlide(el.ownerDocument, session, index, { frameLabel: label });
      el.replaceChildren(...(svg ? [svg] : []));
    });

    const linkAt = (e: MouseEvent): string | null =>
      slideshowLinkUnder(
        editor.session,
        state.value,
        slide.value,
        width.value,
        height.value,
        e.clientX,
        e.clientY
      );

    const onMove = (e: MouseEvent): void => {
      bar.value = true;
      link.value = !!linkAt(e);
      if (idle) clearTimeout(idle);
      idle = setTimeout(() => {
        if (!overBar) bar.value = false;
      }, SLIDESHOW_IDLE_MS);
    };

    const onClick = (e: MouseEvent): void => {
      if (e.button !== 0) return;
      const href = linkAt(e);
      const doc = (e.currentTarget as Element | null)?.ownerDocument ?? document;
      if (href) openSlideshowLink(doc, href);
      else controller.next();
    };

    return () => {
      const s = state.value;
      const fit = fitSlideToScreen(
        size.value.w,
        size.value.h,
        width.value,
        height.value,
        size.value.dpr
      );
      const position = t('slideshow.position', { current: s.index + 1, total: s.slideCount });
      return (
        <div
          ref={root}
          class="pptx-show"
          role="dialog"
          aria-modal="true"
          aria-label={t('slideshow.ariaLabel')}
          tabindex={-1}
          data-screen={s.screen}
          data-idle={String(!bar.value)}
          data-link={String(link.value)}
          onMousemove={onMove}
          onClick={onClick}
        >
          <div
            ref={slide}
            class="pptx-show__slide"
            data-slide-index={s.index}
            style={{
              left: px(fit.left),
              top: px(fit.top),
              width: px(fit.width),
              height: px(fit.height),
            }}
          />
          {s.screen === 'black' || s.screen === 'white' ? (
            <div class="pptx-show__blank" data-color={s.screen} />
          ) : null}
          {s.screen === 'end' ? <div class="pptx-show__end">{t('slideshow.end')}</div> : null}
          {s.typed ? (
            <div class="pptx-show__goto" role="status">
              {t('slideshow.goTo', { number: s.typed })}
            </div>
          ) : null}
          <div
            class="pptx-show__bar"
            role="toolbar"
            aria-label={t('slideshow.controls')}
            data-visible={String(bar.value)}
            onClick={(e: MouseEvent) => e.stopPropagation()}
            onMouseenter={() => {
              overBar = true;
            }}
            onMouseleave={() => {
              overBar = false;
            }}
          >
            <button
              type="button"
              class="pptx-show__button"
              aria-label={t('slideshow.previous')}
              title={t('slideshow.previous')}
              onClick={() => controller.previous()}
            >
              <Icon name="keyboard_arrow_left" className="" />
            </button>
            <span class="pptx-show__position" aria-live="polite">
              {position}
            </span>
            <button
              type="button"
              class="pptx-show__button"
              aria-label={t('slideshow.next')}
              title={t('slideshow.next')}
              onClick={() => controller.next()}
            >
              <Icon name="keyboard_arrow_right" className="" />
            </button>
            <button
              type="button"
              class="pptx-show__button"
              aria-label={t('slideshow.exit')}
              title={t('slideshow.exit')}
              onClick={() => controller.exit()}
            >
              <Icon name="close" className="" />
            </button>
          </div>
        </div>
      );
    };
  },
});

/**
 * The slide show of the editor in `PptxEditorRoot`, the twin of React's `PptxEditor.Slideshow`.
 * Mount it once inside the viewport. It shows nothing until a show starts:
 * `useSlideshow().start(...)`, or the keys F5 (from the start) and Shift+F5 or Ctrl/Cmd+Enter
 * (from the current slide). @public
 */
export const PptxEditorSlideshow = defineComponent({
  name: 'PptxEditorSlideshow',
  setup() {
    const editor = usePptxEditor();
    const show = useSlideshow();
    const anchor = ref<HTMLSpanElement | null>(null);
    // The show start keys, while focus is in this editor (or on nothing). The listener stays
    // registered while a show runs and reads the live state, so a start key pressed right after
    // a show ends is never lost.
    watch(
      editor,
      (instance, _old, onCleanup) => {
        if (!instance) return;
        onCleanup(
          listenSlideshowStartKeys(
            () => anchor.value,
            () => show.controller.value !== null,
            (from) => show.start(from)
          )
        );
      },
      { immediate: true }
    );
    return () => {
      const instance = editor.value;
      const controller = show.controller.value;
      if (!instance || !controller) return <span ref={anchor} hidden />;
      return (
        <SlideshowScreen
          key={showKey(controller, instance)}
          editor={instance}
          controller={controller}
          onClose={show.stop}
        />
      );
    };
  },
});
