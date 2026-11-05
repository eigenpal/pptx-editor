// The hooks a host uses to build its own chrome, after DOCX Editor's `useZoom`,
// `useEditorEvent`, `useEditorSnapshot`, and `useScopeClassName`. Each hook reads the editor
// through its one subscription (`subscribe` and `snapshot`) and calls back into the instance:
// no hook keeps a second copy of editor state. The Vue twins are in `customization-hooks.ts`
// of the Vue adapter.

import { useCallback, useContext, useEffect, useMemo, useRef, useState } from 'react';
import {
  ZOOM_LEVELS,
  onEditorEvent,
  stepEditorZoom,
  zoomStep,
  type EditorSnapshot,
  type PptxEditorEvents,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from './context.ts';
import { ChromeDarkContext, ScopedByAncestorContext } from './color-mode.ts';

/** What {@link useZoom} returns. @public */
export interface UseZoomResult {
  /** The zoom in force. 1 is 100 %. In fit mode it is the zoom that fits the slide. */
  readonly zoom: number;
  /** Whether the editor fits the slide to the window and follows its size. */
  readonly isFit: boolean;
  /** Set a fixed zoom, kept in the range 10 % to 400 %. Leaves fit mode. */
  readonly setZoom: (zoom: number) => void;
  /** Fit the slide to the window, and keep it fitted when the window changes size. */
  readonly fit: () => void;
  /** A fixed 100 %. */
  readonly reset: () => void;
  /** One step of `levels` in from the zoom in force. Does nothing at 400 %. */
  readonly zoomIn: () => void;
  /** One step of `levels` out from the zoom in force. Does nothing at 10 %. */
  readonly zoomOut: () => void;
  /** Whether `zoomIn` can step. */
  readonly canZoomIn: boolean;
  /** Whether `zoomOut` can step. */
  readonly canZoomOut: boolean;
  /** The steps the zoom buttons move through, so a custom control shows the same levels. */
  readonly levels: readonly number[];
}

interface ZoomSlice {
  readonly zoom: number;
  readonly fit: boolean;
}

const selectZoom = (s: EditorSnapshot): ZoomSlice => ({ zoom: s.zoom, fit: s.fitZoom });
const sameZoom = (a: ZoomSlice, b: ZoomSlice): boolean => a.zoom === b.zoom && a.fit === b.fit;

/**
 * Read and set the slide zoom, as DOCX Editor's `useZoom`. Outside `PptxEditor.Root`, and
 * before the editor mounts, it reports 100 % and every action does nothing, so a control can
 * render unconditionally.
 *
 * @example
 * ```tsx
 * const { zoom, isFit, fit, zoomIn, canZoomIn } = useZoom();
 * <button onClick={fit} aria-pressed={isFit}>Fit</button>
 * <button onClick={zoomIn} disabled={!canZoomIn}>+</button>
 * <span>{Math.round(zoom * 100)}%</span>
 * ```
 *
 * @public
 */
export function useZoom(): UseZoomResult {
  const editor = usePptxEditor();
  const { zoom, fit } = useEditorState(selectZoom, sameZoom);
  const setZoom = useCallback((next: number) => editor?.setZoom(next), [editor]);
  return useMemo(() => {
    const next = zoomStep(zoom, 1);
    const previous = zoomStep(zoom, -1);
    return {
      zoom,
      isFit: fit,
      setZoom,
      fit: () => editor?.setZoom('fit'),
      reset: () => setZoom(1),
      zoomIn: () => stepEditorZoom(editor, 1),
      zoomOut: () => stepEditorZoom(editor, -1),
      canZoomIn: !!editor && next !== null,
      canZoomOut: !!editor && previous !== null,
      levels: ZOOM_LEVELS,
    };
  }, [editor, zoom, fit, setZoom]);
}

/**
 * Subscribe to an editor event (`'change'`, `'selectionChange'`, `'rejection'`,
 * `'openFailure'`) for the lifetime of the component, as DOCX Editor's `useEditorEvent`. It
 * does nothing until the nearest `PptxEditor.Root` creates the editor, and subscribes again
 * when the instance changes. A new handler on each render does not subscribe again; the
 * latest one is called.
 *
 * @example
 * ```tsx
 * useEditorEvent('rejection', (rejection) => toast(rejection.reason));
 * ```
 *
 * @public
 */
export function useEditorEvent<E extends keyof PptxEditorEvents>(
  event: E,
  handler: PptxEditorEvents[E]
): void {
  const editor = usePptxEditor();
  const handlerRef = useRef(handler);
  handlerRef.current = handler;
  useEffect(() => {
    if (!editor) return undefined;
    const forward = ((...args: unknown[]) =>
      (handlerRef.current as (...a: unknown[]) => void)(...args)) as PptxEditorEvents[E];
    return onEditorEvent(editor, event, forward);
  }, [editor, event]);
}

/**
 * Re-render the caller each time `editor` publishes a change, as DOCX Editor's
 * `useEditorSnapshot`. Returns a counter that changes on each change, so it also serves as a
 * dependency. Prefer `useEditorState` with a selector: it re-renders only when its slice
 * changes.
 *
 * @public
 */
export function useEditorSnapshot(editor: PptxEditorInstance | null): number {
  const [revision, setRevision] = useState(0);
  useEffect(() => {
    if (!editor) return undefined;
    const bump = (): void => setRevision((n) => n + 1);
    const off = editor.subscribe(bump);
    bump();
    return off;
  }, [editor]);
  return revision;
}

/**
 * The scope classes a part you build yourself adds to its own root element, so the editor's
 * `--doc-*` tokens and chrome styles apply to it, as DOCX Editor's `useScopeClassName`. Inside
 * `PptxEditor.Viewport` it is `''`: the viewport already carries the scope. Outside it, it is
 * `'pptx-editor '`, or `'pptx-editor dark '` when the root's `colorMode` resolved dark. A part
 * that portals out of the viewport (to `document.body`) adds `pptx-editor` itself, with
 * `useEditorColorMode` for the theme.
 *
 * The value ends with a space, so it concatenates cleanly, and `''` collapses away.
 *
 * @example
 * ```tsx
 * const scope = useScopeClassName();
 * <aside className={`${scope}my-panel`} />
 * ```
 *
 * @public
 */
export function useScopeClassName(): '' | 'pptx-editor ' | 'pptx-editor dark ' {
  const scoped = useContext(ScopedByAncestorContext);
  const dark = useContext(ChromeDarkContext);
  if (scoped) return '';
  return dark ? 'pptx-editor dark ' : 'pptx-editor ';
}
