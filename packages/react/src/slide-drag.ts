// Drag to reorder in a slide list: the slide rail (a column, `y`) and the slide sorter (a grid
// read in rows, `x`). Dragging a selected slide moves the whole selection; another slide moves
// alone. The drop is one `moveSlides` command, so it is one undo step. With sections, the drop
// joins the section of the slide under the pointer: a drop before a section's first slide
// (below its header) puts the block into that section. The rules live in core
// (`pane-support.ts`), shared with the Vue adapter.

import { useCallback, useRef, useState, type DragEvent } from 'react';
import {
  dragOverSlide,
  dropSlideDrag,
  slideDropMarker,
  startSlideDrag,
  type SlideDrag,
} from '@pptx-editor.dev/core/editor';
import type { SlideSection } from '@pptx-editor.dev/core/contracts';
import { usePptxEditor } from './context.ts';

/**
 * The drag handlers and drop marker of each slide item of a list. `axis` is the direction in
 * which the list reads: `y` for the rail, `x` for the sorter's rows. `grouped` is the list's
 * sections (`groupSlideSections`), or null when it is flat. `cancel` drops a drag.
 */
export function useSlideDrag(
  axis: 'x' | 'y',
  selected: readonly number[],
  readOnly: boolean,
  grouped: readonly SlideSection[] | null = null
) {
  const editor = usePptxEditor();
  // The drag lives in a ref, as Vue's `shallowRef`: a quick drag fires `dragover` before React
  // re-renders, and a handler that read the last render's state would not call
  // `preventDefault()`, so the browser would refuse the drop. The state copy only repaints the
  // drop marker.
  const dragRef = useRef<SlideDrag | null>(null);
  const [drag, setDragState] = useState<SlideDrag | null>(null);
  const setDrag = useCallback((next: SlideDrag | null) => {
    dragRef.current = next;
    setDragState(next);
  }, []);
  const itemProps = (index: number) => ({
    'data-drop': slideDropMarker(drag, index),
    draggable: !readOnly,
    onDragStart: (e: DragEvent<HTMLElement>) => {
      e.dataTransfer.effectAllowed = 'move';
      setDrag(startSlideDrag(editor, selected, index));
    },
    onDragOver: (e: DragEvent<HTMLElement>) => {
      const current = dragRef.current;
      if (!current) return;
      e.preventDefault();
      setDrag(dragOverSlide(current, index, axis, e.currentTarget.getBoundingClientRect(), e));
    },
    onDragEnd: () => setDrag(null),
    onDrop: (e: DragEvent<HTMLElement>) => {
      e.preventDefault();
      const current = dragRef.current;
      if (!current || !editor) return;
      dropSlideDrag(editor, current, grouped);
      setDrag(null);
    },
  });
  const cancel = useCallback(() => setDrag(null), [setDrag]);
  /** The drop marker of slide `index`, without the handlers (a memoized item reads it). */
  const marker = (index: number) => slideDropMarker(drag, index);
  return { itemProps, cancel, marker };
}
