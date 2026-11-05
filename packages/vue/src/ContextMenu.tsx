/** @jsxImportSource vue */
// A right-click menu at a point: the Vue twin of `ContextMenu` in the React adapter's
// `ContextMenu.tsx`, after DOCX Editor's context menu. The menu opens at the pointer and flips
// to stay inside the window (`placeInViewport`), focus moves into it and returns on close, and
// the rows share the popover menu's keyboard rules (`menu-keyboard.ts`). The rows are
// `MenuRow`, `MenuSeparator`, and `MenuSubmenu` (`menu-row.tsx`).

import {
  Teleport,
  defineComponent,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  focusEdge,
  focusMenuRowAtPointer,
  focusPanelFromPointer,
  focusSurfaceOf,
  handleMenuKey,
  listenContextMenuDismiss,
  panelKeyPressed,
  panelItems,
  placeInViewport,
  popoverScope,
} from '@pptx-editor.dev/core/editor';
import { keepFocus } from './chrome.tsx';
import { usePptxEditor } from './composables.ts';

/** A point in client (viewport) pixels. @public */
export interface ContextMenuPoint {
  /** The horizontal position in CSS pixels from the left of the window. */
  readonly x: number;
  /** The vertical position in CSS pixels from the top of the window. */
  readonly y: number;
}

const contextMenuProps = {
  /** Where the menu opens: the pointer, or the focused element for a keyboard open. */
  at: { type: Object as PropType<ContextMenuPoint>, required: true as const },
  /** Closes the menu: the host stops rendering it. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** The accessible name of the menu. */
  label: { type: String, required: true as const },
  /** Opened from the keyboard (Shift+F10, the Menu key): focus goes to the first row. */
  keyboard: { type: Boolean, default: false },
  /** The element the menu opened on; its editor scope carries the theme tokens. */
  owner: { type: Object as PropType<HTMLElement | null>, default: undefined },
  /** A class name added to the menu, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/**
 * Props of `ContextMenu`, the twin of React's `ContextMenuProps`. `className` (or a host
 * `class`) lands on the menu; the default slot is its rows. @public
 */
export type ContextMenuProps = ExtractPublicPropTypes<typeof contextMenuProps>;

/**
 * A context menu at a point. Closed by Escape (focus returns to where it was), Tab, a click
 * outside, or a window resize. The arrow keys, Home, and End move through the rows. @public
 */
export const ContextMenu = defineComponent({
  name: 'ContextMenu',
  // The menu renders in the editor scope (a teleport), so the host's `class` goes onto it here.
  inheritAttrs: false,
  props: contextMenuProps,
  setup(props, { attrs, slots }) {
    const editor = usePptxEditor();
    const panel = ref<HTMLDivElement | null>(null);
    const pos = shallowRef<{ left: number; top: number; maxHeight?: number } | null>(null);
    const close = (): void => props.onClose();
    const place = (): void => {
      const el = panel.value;
      if (!el) return;
      pos.value = placeInViewport(
        props.at,
        el.offsetWidth,
        // The full height, past a cap from an earlier place (its rows scroll).
        el.scrollHeight + el.offsetHeight - el.clientHeight,
        window.innerWidth,
        window.innerHeight
      );
    };
    // Where focus was when the menu opened: it goes back there on close.
    let back: HTMLElement | null = null;
    // A press outside, Escape, a blur, a resize, or a scroll closes it (core's
    // `listenContextMenuDismiss`).
    let stopDismiss = (): void => {};
    onMounted(() => {
      back = (panel.value?.ownerDocument.activeElement as HTMLElement | null) ?? null;
      place();
      stopDismiss = listenContextMenuDismiss(() => panel.value, close);
    });
    watch(() => [props.at.x, props.at.y], place, { flush: 'post' });
    // Focus moves in once the panel is placed: a hidden panel cannot take focus.
    watch(
      () => pos.value !== null,
      (placed) => {
        const el = panel.value;
        if (!el || !placed) return;
        // A pointer open focuses the panel, with no ring until a key is pressed.
        if (!props.keyboard || !focusEdge(panelItems(el), 'first')) focusPanelFromPointer(el);
      },
      { flush: 'post' }
    );
    onBeforeUnmount(() => {
      stopDismiss();
      const el = panel.value;
      if (!el) return;
      // Focus returns only when it is still in the menu: a click elsewhere keeps its focus.
      const active = el.ownerDocument.activeElement;
      const lost = !active || active === el.ownerDocument.body || el.contains(active);
      if (!lost) return;
      // On the slide, focus goes where typing works: the text input while text is selected.
      if (back?.closest('.pptx-surface') && editor.value && focusSurfaceOf(editor.value)) return;
      if (back?.isConnected) back.focus({ preventScroll: true });
    });
    const onKeydown = (e: KeyboardEvent): void => {
      const el = panel.value;
      panelKeyPressed(el);
      if (!el?.contains(e.target as Node)) return;
      if (e.key === 'Escape' || e.key === 'Tab') {
        e.preventDefault();
        e.stopPropagation();
        close();
        return;
      }
      if (handleMenuKey(el, e.key)) {
        e.preventDefault();
        e.stopPropagation();
      }
    };
    return () => (
      <Teleport to={popoverScope(props.owner)}>
        <div
          ref={panel}
          class={['pptx-menu pptx-contextmenu', attrs.class, props.className]}
          role="menu"
          aria-label={props.label}
          tabindex={-1}
          style={
            pos.value
              ? {
                  left: `${pos.value.left}px`,
                  top: `${pos.value.top}px`,
                  ...(pos.value.maxHeight === undefined
                    ? {}
                    : { maxHeight: `${pos.value.maxHeight}px` }),
                }
              : { visibility: 'hidden', left: `${props.at.x}px`, top: `${props.at.y}px` }
          }
          onMousedown={keepFocus}
          // The pointer moves the focused row, so one row looks active.
          onMousemove={(e: MouseEvent) => focusMenuRowAtPointer(panel.value, e)}
          onContextmenu={(e: MouseEvent) => e.preventDefault()}
          onKeydown={onKeydown}
        >
          {slots.default?.()}
        </div>
      </Teleport>
    );
  },
});
