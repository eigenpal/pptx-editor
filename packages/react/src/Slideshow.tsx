/** @jsxImportSource react */
// The slide show: full screen (the Fullscreen API, or a window-covering overlay when the
// browser refuses), one slide painted by the engine's `paintSlide`, the usual show keys, and an
// auto-hiding control bar. Navigation state is the engine's `SlideshowController`; this
// component only paints it and feeds input in. Full screen, the start keys, the show keys, the
// screen size, and links are core's (`slideshow-screen.ts`), shared with the Vue
// `PptxEditorSlideshow`.

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type MouseEvent as ReactMouseEvent,
} from 'react';
import {
  SLIDESHOW_IDLE_MS,
  enterSlideshowScreen,
  fitSlideToScreen,
  listenSlideshowKeys,
  listenSlideshowStartKeys,
  observeSlideshowScreen,
  followSlideshowLink,
  paintSlideshowSlide,
  slideshowLinkUnder,
  startSlideshowController,
  translateKeyOrText,
  type PptxEditorInstance,
  type SlideshowController,
  type SlideshowScreenSize,
} from '@pptx-editor.dev/core/editor';
import { Icon } from './chrome.tsx';
import {
  type UseSlideshowReturn,
  useEditorState,
  usePptxEditor,
  useSlideshow,
  useTranslation,
} from './context.ts';

/**
 * The slide show state `PptxEditor.Root` publishes: the running controller, and start and stop.
 * A show starts only on an open document with slides; viewing mode can present. @internal
 */
export function useSlideshowHandle(editor: PptxEditorInstance | null): UseSlideshowReturn {
  const [show, setShow] = useState<SlideshowController | null>(null);
  return useMemo(
    () => ({
      controller: show,
      start(from: 'start' | 'current') {
        const next = startSlideshowController(editor, from);
        if (next) setShow(next);
      },
      stop: () => setShow(null),
    }),
    [editor, show]
  );
}

/**
 * The slide show of the editor in `PptxEditor.Root`. Mount it once inside the viewport. It
 * shows nothing until a show starts: `useSlideshow().start(...)`, or the keys F5 (from the
 * start) and Shift+F5 or Ctrl/Cmd+Enter (from the current slide). @public
 */
export function Slideshow() {
  const editor = usePptxEditor();
  const show = useSlideshow();
  const anchor = useRef<HTMLSpanElement>(null);
  const { start } = show;

  // The show start keys, while focus is in this editor (or on nothing). The listener stays
  // registered while a show runs and reads the live state, so a start key pressed right after
  // a show ends is never lost in the gap before an effect could register it again.
  const running = useRef(show.controller !== null);
  running.current = show.controller !== null;
  const startRef = useRef(start);
  startRef.current = start;
  useEffect(() => {
    if (!editor) return;
    return listenSlideshowStartKeys(
      () => anchor.current,
      () => running.current,
      (from) => startRef.current(from)
    );
  }, [editor]);

  if (!editor || !show.controller) return <span ref={anchor} hidden />;
  return <SlideshowScreen editor={editor} controller={show.controller} onClose={show.stop} />;
}

function SlideshowScreen({
  editor,
  controller,
  onClose,
}: {
  editor: PptxEditorInstance;
  controller: SlideshowController;
  onClose: () => void;
}) {
  const { t } = useTranslation();
  const root = useRef<HTMLDivElement>(null);
  const slide = useRef<HTMLDivElement>(null);
  const subscribe = useCallback((cb: () => void) => controller.subscribe(cb), [controller]);
  const get = useCallback(() => controller.state, [controller]);
  const state = useSyncExternalStore(subscribe, get, get);
  const width = useEditorState((s) => s.slideWidth);
  const height = useEditorState((s) => s.slideHeight);
  const renderKey = useEditorState((s) => s.slides[state.index]?.key ?? '');
  const [size, setSize] = useState<SlideshowScreenSize>({ w: 0, h: 0, dpr: 1 });
  const [bar, setBar] = useState(false);
  const [link, setLink] = useState(false);
  const idle = useRef<ReturnType<typeof setTimeout> | null>(null);
  const overBar = useRef(false);

  // The deck can change under a running show (fonts load, a collaborator edits).
  useEffect(
    () => editor.subscribe(() => controller.update(editor.snapshot().slides)),
    [editor, controller]
  );

  // The show ended: the editor returns to the slide shown last.
  useEffect(() => {
    if (!state.done) return;
    editor.goToSlide(state.index);
    onClose();
  }, [state.done, state.index, editor, onClose]);

  // Full screen when the browser allows it; the fixed overlay covers the window otherwise
  // (core's `enterSlideshowScreen`). A layout effect, so focus returns in the same commit that
  // removes the show, not after the next paint.
  useLayoutEffect(() => {
    const el = root.current;
    if (!el) return;
    return enterSlideshowScreen(el);
  }, [controller]);

  // Every key goes to the show. A layout effect, so the listener leaves in the same commit as
  // the show: a start key pressed right after the show ends is not swallowed by a listener that
  // a passive cleanup has not removed.
  useLayoutEffect(() => listenSlideshowKeys(controller, () => root.current), [controller]);

  // The screen size, in CSS and device pixels.
  useLayoutEffect(() => {
    const el = root.current;
    if (!el) return;
    return observeSlideshowScreen(el, setSize);
  }, []);

  // Paint the slide on show with the editor's own painter.
  useLayoutEffect(() => {
    const el = slide.current;
    const session = editor.session;
    if (!el) return;
    if (!session || state.screen === 'end') {
      el.replaceChildren();
      return;
    }
    const svg = paintSlideshowSlide(el.ownerDocument, session, state.index, {
      frameLabel: (kind) =>
        t('canvas.placeholderObject', { kind: translateKeyOrText(t, `canvas.kinds.${kind}`) }),
    });
    el.replaceChildren(...(svg ? [svg] : []));
  }, [editor, state.index, state.screen, renderKey, t]);

  useEffect(
    () => () => {
      if (idle.current) clearTimeout(idle.current);
    },
    []
  );

  const fit = fitSlideToScreen(size.w, size.h, width, height, size.dpr);

  const linkAt = (e: ReactMouseEvent): string | null =>
    slideshowLinkUnder(editor.session, state, slide.current, width, height, e.clientX, e.clientY);

  const onMove = (e: ReactMouseEvent): void => {
    setBar(true);
    setLink(!!linkAt(e));
    if (idle.current) clearTimeout(idle.current);
    idle.current = setTimeout(() => {
      if (!overBar.current) setBar(false);
    }, SLIDESHOW_IDLE_MS);
  };

  const onClick = (e: ReactMouseEvent): void => {
    if (e.button !== 0) return;
    const href = linkAt(e);
    // A slide link moves the show; any other link opens in a new tab.
    if (href) followSlideshowLink(controller, e.currentTarget.ownerDocument, href);
    else controller.next();
  };

  const position = t('slideshow.position', {
    current: state.index + 1,
    total: state.slideCount,
  });
  return (
    <div
      ref={root}
      className="pptx-show"
      role="dialog"
      aria-modal="true"
      aria-label={t('slideshow.ariaLabel')}
      tabIndex={-1}
      data-screen={state.screen}
      data-idle={!bar}
      data-link={link}
      onMouseMove={onMove}
      onClick={onClick}
    >
      <div
        ref={slide}
        className="pptx-show__slide"
        data-slide-index={state.index}
        style={{ left: fit.left, top: fit.top, width: fit.width, height: fit.height }}
      />
      {state.screen === 'black' || state.screen === 'white' ? (
        <div className="pptx-show__blank" data-color={state.screen} />
      ) : null}
      {state.screen === 'end' ? <div className="pptx-show__end">{t('slideshow.end')}</div> : null}
      {state.typed ? (
        <div className="pptx-show__goto" role="status">
          {t('slideshow.goTo', { number: state.typed })}
        </div>
      ) : null}
      <div
        className="pptx-show__bar"
        role="toolbar"
        aria-label={t('slideshow.controls')}
        data-visible={bar}
        onClick={(e) => e.stopPropagation()}
        onMouseEnter={() => (overBar.current = true)}
        onMouseLeave={() => (overBar.current = false)}
      >
        <button
          type="button"
          className="pptx-show__button"
          aria-label={t('slideshow.previous')}
          title={t('slideshow.previous')}
          onClick={() => controller.previous()}
        >
          <Icon name="keyboard_arrow_left" className="" />
        </button>
        <span className="pptx-show__position" aria-live="polite">
          {position}
        </span>
        <button
          type="button"
          className="pptx-show__button"
          aria-label={t('slideshow.next')}
          title={t('slideshow.next')}
          onClick={() => controller.next()}
        >
          <Icon name="keyboard_arrow_right" className="" />
        </button>
        <button
          type="button"
          className="pptx-show__button"
          aria-label={t('slideshow.exit')}
          title={t('slideshow.exit')}
          onClick={() => controller.exit()}
        >
          <Icon name="close" className="" />
        </button>
      </div>
    </div>
  );
}
