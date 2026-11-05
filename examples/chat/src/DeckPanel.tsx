// The deck panel beside the chat: a compact header with the deck's own actions, the editor's
// toolbar, and a workspace that follows the panel's shape. A wide, short panel (expanded)
// gets the packaged layout: the slide rail on the left. A tall or narrow panel, the usual
// shape beside a chat, stacks the slide at full width, the same rail as a row under it
// (`orientation="horizontal"`), and the speaker notes in the height that is left, so no
// space goes to empty margins.

import { useLayoutEffect, useRef, useState, type CSSProperties } from 'react';
import { Icon, PptxEditor, keepFocus, useEditorState, useSlideshow } from '@pptx-editor.dev/react';
import { COLLAPSE_CONTENT, EXPAND_CONTENT, SLIDES } from './icons.ts';
import { TEXT } from './text.ts';

/** The fitted slide's margin: `.pptx-surface` pads 24px on each side, and fit leaves it. */
const SURFACE_PAD = 48;

type Shape = { readonly mode: 'wide' | 'stacked'; readonly canvas: number | null };

/**
 * The workspace's layout for its size: stacked unless the panel is wide and short. The mode
 * reads the width the panel is going to (`target`), so it changes once when Expand starts, not
 * halfway through the animation. A stacked slide asks for the height that fits the width;
 * flexbox gives the rest to the rail and the notes (styles.css `.deck-ws`).
 */
function shapeFor(width: number, target: number, height: number, ratio: number): Shape {
  if (target >= 960 && target / Math.max(height, 1) >= 1.5) return { mode: 'wide', canvas: null };
  return { mode: 'stacked', canvas: Math.round((width - SURFACE_PAD) * ratio + SURFACE_PAD) };
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
  return (
    <div ref={box} className="deck-ws" data-mode={shape.mode}>
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
          <PptxEditor.SlideRail {...rail} orientation="horizontal" className="deck-ws__rail" />
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
      <header className="deck__head">
        {sheet ? (
          <button
            type="button"
            className="chat-icon-btn"
            aria-label={TEXT.back}
            onMouseDown={keepFocus}
            onClick={onClose}
          >
            <Icon name="keyboard_arrow_left" className="" />
          </button>
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
          <button
            type="button"
            className="chat-icon-btn"
            aria-label={TEXT.present}
            data-pptx-tooltip={TEXT.present}
            disabled={!canPresent}
            onMouseDown={keepFocus}
            onClick={() => show.start('current')}
          >
            <Icon name="slideshow" className="" />
          </button>
          <button
            type="button"
            className="chat-icon-btn"
            aria-label={TEXT.download}
            data-pptx-tooltip={TEXT.download}
            disabled={!canPresent}
            onMouseDown={keepFocus}
            onClick={onDownload}
          >
            <Icon name="file_download" className="" />
          </button>
          {sheet ? null : (
            <>
              <button
                type="button"
                className="chat-icon-btn"
                aria-label={expanded ? TEXT.collapse : TEXT.expand}
                aria-pressed={expanded}
                data-pptx-tooltip={expanded ? TEXT.collapse : TEXT.expand}
                onMouseDown={keepFocus}
                onClick={onExpand}
              >
                <Icon paths={expanded ? COLLAPSE_CONTENT : EXPAND_CONTENT} className="" />
              </button>
              <button
                type="button"
                className="chat-icon-btn"
                aria-label={TEXT.closeDeck}
                data-pptx-tooltip={TEXT.closeDeck}
                onMouseDown={keepFocus}
                onClick={onClose}
              >
                <Icon name="close" className="" />
              </button>
            </>
          )}
        </div>
      </header>
      <div className="deck__dock">
        <PptxEditor.Toolbar>
          <PptxEditor.Toolbar.ChangeCase hidden />
          <PptxEditor.Toolbar.Link hidden />
        </PptxEditor.Toolbar>
      </div>
      <Workspace target={targetWidth} />
      <PptxEditor.Slideshow />
    </PptxEditor.Viewport>
  );
}
