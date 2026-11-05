// The deck panel's size and state: open, expanded over the chat, a full-window sheet on a
// narrow window, and the width that the divider sets. The panel keeps a fraction of the
// window, so a window resize keeps the split, and the chat never gets narrower than it can
// read. A drag past the chat's least width resists, and a release far past it expands the
// panel, so a quick fling works as well as the Expand button.

import {
  useCallback,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent,
  type PointerEvent,
} from 'react';

/** The least widths: below them the chat or the slide stops being usable. */
const MIN_CHAT = 380;
const MIN_PANEL = 460;
/** The panel's first share of the window. A slide needs width more than a chat does. */
const DEFAULT_FRACTION = 0.6;
/** Narrower windows show the panel as a full-window sheet. */
const SHEET_BELOW = 860;
/** A release that leaves the chat narrower than this expands the panel. */
const EXPAND_BELOW = 220;

const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v));

/** Past a limit, the edge moves a fifth of the pointer's distance, as a rubber band. */
const resist = (v: number, lo: number, hi: number): number =>
  v < lo ? lo - (lo - v) * 0.2 : v > hi ? hi + (v - hi) * 0.2 : v;

export function usePanelLayout() {
  const shell = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(() => window.innerWidth);
  const [open, setOpen] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const [fraction, setFraction] = useState(DEFAULT_FRACTION);
  /** The panel width while a drag runs, before the release settles it. */
  const [dragWidth, setDragWidth] = useState<number | null>(null);
  const drag = useRef<{ id: number; right: number; wasExpanded: boolean } | null>(null);

  useLayoutEffect(() => {
    const el = shell.current;
    if (!el) return;
    const observer = new ResizeObserver(([entry]) => setWidth(entry.contentRect.width));
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const sheet = width < SHEET_BELOW;
  const maxPanel = Math.max(MIN_PANEL, width - MIN_CHAT);
  const settled = clamp(Math.round(fraction * width), MIN_PANEL, maxPanel);
  const panelWidth =
    sheet || expanded
      ? width
      : dragWidth === null
        ? settled
        : resist(dragWidth, MIN_PANEL, maxPanel);

  const onPointerDown = useCallback(
    (e: PointerEvent<HTMLDivElement>) => {
      if (e.button !== 0 || drag.current || sheet) return;
      e.preventDefault();
      e.currentTarget.setPointerCapture(e.pointerId);
      const right = shell.current?.getBoundingClientRect().right ?? window.innerWidth;
      drag.current = { id: e.pointerId, right, wasExpanded: expanded };
      setDragWidth(expanded ? width : settled);
      setExpanded(false);
    },
    [sheet, expanded, width, settled]
  );
  const onPointerMove = useCallback((e: PointerEvent<HTMLDivElement>) => {
    const d = drag.current;
    // A second finger does not take the drag over.
    if (!d || e.pointerId !== d.id) return;
    setDragWidth(d.right - e.clientX);
  }, []);
  const finish = useCallback(
    (e: PointerEvent<HTMLDivElement>) => {
      const d = drag.current;
      if (!d || e.pointerId !== d.id) return;
      drag.current = null;
      const raw = d.right - e.clientX;
      setDragWidth(null);
      if (width - raw < EXPAND_BELOW) setExpanded(true);
      else setFraction(clamp(raw, MIN_PANEL, maxPanel) / width);
    },
    [width, maxPanel]
  );
  /** A drag the browser cancels (a system gesture) puts the panel back where it was. */
  const cancel = useCallback((e: PointerEvent<HTMLDivElement>) => {
    const d = drag.current;
    if (!d || e.pointerId !== d.id) return;
    drag.current = null;
    setDragWidth(null);
    if (d.wasExpanded) setExpanded(true);
  }, []);
  const onKeyDown = useCallback(
    (e: KeyboardEvent<HTMLDivElement>) => {
      // Left and Right move the divider; Home gives the deck its least width, End the most
      // beside the chat, and Enter expands the deck or brings the chat back.
      const lo = MIN_PANEL / width;
      const hi = maxPanel / width;
      const step = e.key === 'ArrowLeft' ? 0.02 : e.key === 'ArrowRight' ? -0.02 : 0;
      if (e.key === 'Enter') setExpanded((x) => !x);
      else if (e.key === 'Home' || e.key === 'End') {
        setExpanded(false);
        setFraction(e.key === 'Home' ? lo : hi);
      } else if (step) {
        setExpanded(false);
        setFraction((f) => clamp(f + step, lo, hi));
      } else return;
      e.preventDefault();
    },
    [width, maxPanel]
  );

  const setOpenAndReset = useCallback((next: boolean) => {
    setOpen(next);
    if (!next) setExpanded(false);
  }, []);

  return {
    shell,
    open,
    setOpen: setOpenAndReset,
    expanded,
    toggleExpanded: useCallback(() => setExpanded((x) => !x), []),
    sheet,
    resizing: dragWidth !== null,
    panelWidth,
    fraction: panelWidth / Math.max(width, 1),
    /** The deck's least share of the window, for the divider's `aria-valuemin`. */
    minFraction: Math.min(1, MIN_PANEL / Math.max(width, 1)),
    handleProps: {
      onPointerDown,
      onPointerMove,
      onPointerUp: finish,
      onPointerCancel: cancel,
      onKeyDown,
      // A double click puts the divider back.
      onDoubleClick: () => (setExpanded(false), setFraction(DEFAULT_FRACTION)),
    },
  };
}
