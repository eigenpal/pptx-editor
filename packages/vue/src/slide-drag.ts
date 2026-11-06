// Drag to reorder in a slide list: the Vue twin of the React adapter's `slide-drag.ts`. The
// slide rail reads as a column (`y`), the slide sorter as a grid read in rows (`x`). Dragging a
// selected slide moves the whole selection; another slide moves alone. The drop is one
// `moveSlides` command, so it is one undo step; the rules live in core (`pane-support.ts`).

import { shallowRef, toValue, type MaybeRefOrGetter } from 'vue';
import {
  dragOverSlide,
  dropSlideDrag,
  slideDropMarker,
  startSlideDrag,
  type SlideDrag,
} from '@pptx-editor.dev/core/editor';
import type { SlideSection } from '@pptx-editor.dev/core/contracts';
import { usePptxEditor } from './composables.ts';

/** The attributes and handlers of one slide item. @internal */
export interface SlideDragItemProps {
  readonly 'data-drop': 'before' | 'after' | undefined;
  readonly draggable: boolean;
  readonly onDragstart: (e: DragEvent) => void;
  readonly onDragover: (e: DragEvent) => void;
  readonly onDragend: () => void;
  readonly onDrop: (e: DragEvent) => void;
}

/** What `useSlideDrag` returns. @internal */
export interface UseSlideDragReturn {
  /** The drag attributes and handlers of slide `index`. */
  readonly itemProps: (index: number) => SlideDragItemProps;
  /** Drop a drag in progress. */
  readonly cancel: () => void;
  /** The drop marker of slide `index`, without the handlers (a rail item reads it). */
  readonly marker: (index: number) => SlideDragItemProps['data-drop'];
}

/**
 * The drag handlers and drop marker of each slide item of a list, as React's `useSlideDrag`.
 * `grouped` is the list's sections (`groupSlideSections`), or null when it is flat. @internal
 */
export function useSlideDrag(
  axis: MaybeRefOrGetter<'x' | 'y'>,
  selected: MaybeRefOrGetter<readonly number[]>,
  readOnly: MaybeRefOrGetter<boolean>,
  grouped: MaybeRefOrGetter<readonly SlideSection[] | null> = null
): UseSlideDragReturn {
  const editor = usePptxEditor();
  const drag = shallowRef<SlideDrag | null>(null);
  const itemProps = (index: number): SlideDragItemProps => ({
    'data-drop': slideDropMarker(drag.value, index),
    draggable: !toValue(readOnly),
    onDragstart: (e) => {
      if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
      drag.value = startSlideDrag(editor.value, toValue(selected), index);
    },
    onDragover: (e) => {
      if (!drag.value) return;
      e.preventDefault();
      const box = (e.currentTarget as HTMLElement).getBoundingClientRect();
      drag.value = dragOverSlide(drag.value, index, toValue(axis), box, e);
    },
    onDragend: () => {
      drag.value = null;
    },
    onDrop: (e) => {
      e.preventDefault();
      if (!drag.value || !editor.value) return;
      dropSlideDrag(editor.value, drag.value, toValue(grouped));
      drag.value = null;
    },
  });
  return {
    itemProps,
    cancel: () => {
      drag.value = null;
    },
    marker: (index) => slideDropMarker(drag.value, index),
  };
}
