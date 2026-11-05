// Roving focus for a toolbar element: one Tab stop and arrow keys (`toolbar-focus.ts` in
// core), the Vue twin of React's `useRovingToolbar` (`toolbar-measure.ts`). The room and
// separator measures live in core (`toolbarRoom`, `separatorCost`), shared by both adapters.

import { onBeforeUnmount, onMounted, onUpdated, type Ref } from 'vue';
import {
  syncToolbarTabStop,
  toolbarArrowKey,
  toolbarFocusTracker,
} from '@pptx-editor.dev/core/editor';

/** What `useRovingToolbar` returns: the handlers for the bar element. @internal */
export interface UseRovingToolbarReturn {
  readonly onFocusin: (event: FocusEvent) => void;
  readonly onFocusout: (event: FocusEvent) => void;
  readonly onKeydown: (event: KeyboardEvent) => void;
}

/**
 * One Tab stop and arrow keys for the bar in `root`: the control that last had focus keeps the
 * Tab stop, and the arrows move between the controls, disabled ones too. A control re-renders
 * on its own (it turns on or off) and host content changes without a bar render, so the bar's
 * subtree is watched as well as its own renders. When the focused control leaves the bar with
 * its group (the selection changed), the focus stays in the bar (`toolbarFocusTracker`). @internal
 */
export function useRovingToolbar(root: Readonly<Ref<HTMLElement | null>>): UseRovingToolbarReturn {
  const tracker = toolbarFocusTracker();
  const sync = (): void => {
    const el = root.value;
    if (!el) return;
    syncToolbarTabStop(el, el.ownerDocument.activeElement);
    tracker.restore(el);
  };
  let observer: MutationObserver | null = null;
  onMounted(() => {
    sync();
    if (root.value && typeof MutationObserver === 'function') {
      observer = new MutationObserver(sync);
      observer.observe(root.value, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['aria-disabled', 'hidden'],
      });
    }
  });
  onUpdated(sync);
  onBeforeUnmount(() => observer?.disconnect());
  return {
    onFocusin: (event) => {
      if (!root.value) return;
      tracker.focus(root.value, event.target);
      syncToolbarTabStop(root.value, event.target as Element);
    },
    onFocusout: (event) => {
      if (root.value) tracker.blur(root.value, event.target, event.relatedTarget);
    },
    onKeydown: (event) => {
      const el = root.value;
      if (!el) return;
      const rtl = getComputedStyle(el).direction === 'rtl';
      if (toolbarArrowKey(el, event, rtl)) event.preventDefault();
    },
  };
}
