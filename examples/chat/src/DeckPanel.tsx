// The deck panel beside the chat: a compact header with the deck's own actions, the editor's
// toolbar, and a workspace that follows the panel's shape. A wide, short panel (expanded)
// gets the packaged layout: the slide rail on the left. A tall or narrow panel, the usual
// shape beside a chat, stacks the slide at full width, the same rail as a row under it
// (`orientation="horizontal"`), and the speaker notes in the height that is left, so no
// space goes to empty margins.

import { useLayoutEffect, useRef, useState, type CSSProperties } from 'react';
import { Icon, PptxEditor, useEditorState, useSlideshow } from '@pptx-editor.dev/react';
import { COLLAPSE_CONTENT, EXPAND_CONTENT, SLIDES } from './icons.ts';
import { TEXT } from './text.ts';

/**
 * The room around the stacked slide, on each side. The workspace sets it as `--doc-slide-padding`, so
 * the editor's fit and the stacked canvas height below use the same value.
 */
const SLIDE_PADDING = 12;

type Shape = {
  readonly mode: 'wide' | 'stacked';
  readonly canvas: number | null;
  /** The panel is narrow: the rail row hides its move buttons. */
  readonly narrow: boolean;
};

/** The width under which the rail row hides its move buttons (styles.css uses the same). */
const NARROW = 560;

/**
 * The workspace's layout for its size: stacked unless the panel is wide and short. The mode
 * reads the width the panel is going to (`target`), so it changes once when Expand starts, not
 * halfway through the animation. A stacked slide takes only the height that fits the width;
 * flexbox gives the rest to the notes (styles.css `.deck-ws`).
 */
function shapeFor(width: number, target: number, height: number, ratio: number): Shape {
  const narrow = width > 0 && width <= NARROW;
  if (target >= 960 && target / Math.max(height, 1) >= 1.5)
    return { mode: 'wide', canvas: null, narrow };
  return {
    mode: 'stacked',
    canvas: Math.round((width - 2 * SLIDE_PADDING) * ratio + 2 * SLIDE_PADDING),
    narrow,
  };
}

function useShape(target: number): [React.RefObject<HTMLDivElement | null>, Shape] {
  const box = useRef<HTMLDivElement>(null);
  const ratio = useEditorState((s) => (s.slideWidth ? s.slideHeight / s.slideWidth : 9 / 16));
  const [size, setSize] = useState({ width: 0, height: 0 });
  useLayoutEffect(() => {
    const el = box.current;
    if (!el) return;
    const observer = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect;
      setSize((s) => (s.width === width && s.height === height ? s : { width, height }));
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);
  return [box, shapeFor(size.width, target, size.height, ratio)];
}

function Workspace({ target }: { target: number }) {
  const [box, shape] = useShape(target);
  const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
  // One open state for both rails, so hidden thumbnails stay hidden when the layout switches.
  const [railOpen, setRailOpen] = useState(true);
  const rail = { open: railOpen, onOpenChange: setRailOpen };
  const canvasStyle: CSSProperties | undefined =
    shape.canvas === null ? undefined : { flexBasis: shape.canvas };
  const stacked = shape.mode === 'stacked';
  const tokens = { '--doc-slide-padding': `${SLIDE_PADDING}px` } as CSSProperties;
  return (
    <div ref={box} className="deck-ws" data-mode={shape.mode} style={stacked ? tokens : undefined}>
      {!stacked && !sorter ? <PptxEditor.SlideRail {...rail} /> : null}
      <div className="deck-ws__main">
        {sorter ? <PptxEditor.SlideSorter /> : null}
        <div className="deck-ws__canvas" style={canvasStyle} hidden={sorter}>
          <PptxEditor.Content />
          <PptxEditor.FindBar />
          <PptxEditor.Loading overlay />
          <PptxEditor.FontNotice />
          {stacked ? null : <PptxEditor.ZoomPill />}
        </div>
        {stacked && !sorter ? (
          <PptxEditor.SlideRail {...rail} orientation="horizontal" className="deck-ws__rail">
            {/* A narrow row drops the move buttons, so more slides fit. Ctrl+Left and
                Ctrl+Right, and the slide menu, still move slides. */}
            <PptxEditor.SlideRail.MoveEarlier hidden={shape.narrow} />
            <PptxEditor.SlideRail.MoveLater hidden={shape.narrow} />
          </PptxEditor.SlideRail>
        ) : null}
        <PptxEditor.NotesPane className="deck-ws__notes" />
        <PptxEditor.EditNotice />
        <PptxEditor.OpenFailure />
      </div>
    </div>
  );
}

export function DeckPanel({
  title,
  expanded,
  sheet,
  targetWidth,
  onClose,
  onExpand,
  onDownload,
}: {
  title: string;
  expanded: boolean;
  /** The panel covers the whole window (a phone): Back replaces Expand and Close. */
  sheet: boolean;
  /** The panel's width once its animation ends. */
  targetWidth: number;
  onClose: () => void;
  onExpand: () => void;
  onDownload: () => void;
}) {
  const count = useEditorState((s) => s.slides.length);
  const edited = useEditorState((s) => s.canUndo);
  const canPresent = useEditorState((s) => s.status === 'ready' && s.slides.length > 0);
  const show = useSlideshow();
  return (
    <PptxEditor.Viewport className="deck">
      {/* The header's buttons are `Toolbar.Action`s: the editor's tooltip, focus guard, and
          disabled state, on the chat's own icons, with no private class names. */}
      <header className="deck__head">
        {sheet ? (
          <PptxEditor.Toolbar.Action
            label={TEXT.back}
            className="deck__action"
            icon={<Icon name="keyboard_arrow_left" className="" />}
            onSelect={onClose}
          />
        ) : null}
        <span className="deck__glyph" aria-hidden="true">
          <Icon paths={SLIDES} className="" />
        </span>
        <div className="deck__title">
          <span className="deck__name">{title}</span>
          <span className="deck__meta">
            {count ? TEXT.slides(count) : TEXT.presentation}
            {/* The word swaps in place; the dot and the word fade together. */}
            <span className="deck__edited" data-shown={edited || undefined}>
              {' · '}
              {TEXT.edited}
            </span>
          </span>
        </div>
        <div className="deck__actions">
          <PptxEditor.Toolbar.Action
            label={TEXT.present}
            className="deck__action"
            icon={<Icon name="slideshow" className="" />}
            disabled={!canPresent}
            onSelect={() => show.start('current')}
          />
          <PptxEditor.Toolbar.Action
            label={TEXT.download}
            className="deck__action"
            icon={<Icon name="file_download" className="" />}
            disabled={!canPresent}
            onSelect={onDownload}
          />
          {sheet ? null : (
            <>
              <PptxEditor.Toolbar.Action
                label={expanded ? TEXT.collapse : TEXT.expand}
                className="deck__action"
                icon={<Icon paths={expanded ? COLLAPSE_CONTENT : EXPAND_CONTENT} className="" />}
                active={expanded}
                onSelect={onExpand}
              />
              <PptxEditor.Toolbar.Action
                label={TEXT.closeDeck}
                className="deck__action"
                icon={<Icon name="close" className="" />}
                onSelect={onClose}
              />
            </>
          )}
        </div>
      </header>
      <div className="deck__dock">
        <PptxEditor.Toolbar className="deck__toolbar">
          <PptxEditor.Toolbar.ChangeCase hidden />
          <PptxEditor.Toolbar.Link hidden />
        </PptxEditor.Toolbar>
      </div>
      <Workspace target={targetWidth} />
      <PptxEditor.Slideshow />
    </PptxEditor.Viewport>
  );
}
