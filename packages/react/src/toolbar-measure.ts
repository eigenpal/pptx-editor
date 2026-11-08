// The toolbar's measuring and focus helpers, kept apart from `Toolbar.tsx` (its 1,000-line cap):
// roving focus here, and the room and separator measures from core. After DOCX Editor's
// `packages/core/src/editor/toolbar-measure.ts` and `useToolbarOverflow`. DOCX Editor commit
// adf0e178.

import { useLayoutEffect, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import {
  syncToolbarTabStop,
  toolbarArrowKey,
  toolbarFocusTracker,
} from '@pptx-editor.dev/core/editor';

// The room and separator measures live in core, shared with the Vue twin.
export { separatorCost, toolbarRoom } from '@pptx-editor.dev/core/editor';

/**
 * One Tab stop and arrow keys for a toolbar element (`toolbar-focus.ts` in core): the control
 * that last had focus keeps the Tab stop, and the arrows move between the controls, disabled
 * ones too. When the focused control leaves the bar with its group (the selection changed),
 * the focus stays in the bar (`toolbarFocusTracker`).
 */
export function useRovingToolbar(ref: { readonly current: HTMLElement | null }) {
  const [tracker] = useState(toolbarFocusTracker);
  useLayoutEffect(() => {
    const bar = ref.current;
    if (!bar) return;
    syncToolbarTabStop(bar, bar.ownerDocument.activeElement);
    tracker.restore(bar);
  });
  return {
    onFocus: (event: { readonly target: EventTarget }) => {
      if (!ref.current) return;
      tracker.focus(ref.current, event.target);
      syncToolbarTabStop(ref.current, event.target as Element);
    },
    onBlur: (event: {
      readonly target: EventTarget;
      readonly relatedTarget: EventTarget | null;
    }) => {
      if (ref.current) tracker.blur(ref.current, event.target, event.relatedTarget);
    },
    onKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>) => {
      const bar = ref.current;
      if (!bar) return;
      const rtl = getComputedStyle(bar).direction === 'rtl';
      if (toolbarArrowKey(bar, event.nativeEvent, rtl)) event.preventDefault();
    },
  };
}
