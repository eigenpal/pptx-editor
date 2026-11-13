/** @jsxImportSource vue */
// The slide canvas's context menu request: the Vue twin of the React adapter's
// `CanvasContextMenu.tsx`. The engine surface stops the browser's own menu and selects what is
// under the pointer first (`surface-context-menu.ts`); this composable reads the selection,
// opens a menu request for it (shapes, pictures, text being edited, or empty slide area), and
// renders the `popups.contextMenu` entry (or the entry of a `PptxEditorContextMenu` placed in
// the root): the packaged `PptxEditorContextMenu` by default, a
// host renderer, or nothing for `false`. `PptxEditorContent` binds it.

import { h, shallowRef, watch, type VNodeChild } from 'vue';
import {
  canvasMenuKind,
  canvasMenuPoint,
  isContextMenuKey,
  type CanvasMenuKind,
} from '@pptx-editor.dev/core/editor';
import type { ContextMenuPoint } from './ContextMenu.tsx';
import { usePptxEditor } from './composables.ts';
import {
  ContextMenuTargetProvider,
  PptxEditorContextMenu,
  type ContextMenuRequest,
} from './context-menu-parts.tsx';
import { useCanvasMenuEntry } from './popups.ts';

interface CanvasMenuState {
  readonly at: ContextMenuPoint;
  readonly kind: CanvasMenuKind;
  readonly keyboard: boolean;
  readonly owner: HTMLElement;
}

/** What `useCanvasContextMenu` returns. @internal */
export interface UseCanvasContextMenuReturn {
  /** Bind to the canvas element's `contextmenu` event. */
  readonly onContextmenu: (e: MouseEvent) => void;
  /** Bind to the canvas element's `keydown` event: Shift+F10 and the Menu key. */
  readonly onKeydown: (e: KeyboardEvent) => void;
  /** Render next to the canvas: the open menu, or nothing. */
  readonly layer: () => VNodeChild;
}

/**
 * The canvas context menu. Returns the handlers the canvas host binds (right-click, Shift+F10,
 * and the Menu key) and the layer that renders the open menu. @internal
 */
export function useCanvasContextMenu(): UseCanvasContextMenuReturn {
  const editor = usePptxEditor();
  const entry = useCanvasMenuEntry();
  const menu = shallowRef<ContextMenuRequest | null>(null);
  // An entry that turns `false` (a placed menu turned `disabled`) closes the open menu.
  watch(
    () => entry() === false,
    (off) => {
      if (off) menu.value = null;
    }
  );
  // Each request mounts the menu anew, so a part reads a fixed request (`useContextMenuRequest`).
  let opened = 0;
  let keyboardAt = 0;
  const open = (owner: HTMLElement, at: ContextMenuPoint | null): void => {
    const instance = editor.value;
    if (entry() === false || !instance || instance.snapshot().status !== 'ready') return;
    const kind = canvasMenuKind(instance.snapshot());
    const state: CanvasMenuState = {
      // From the keyboard: open at the caret, the selection, or the middle of the canvas.
      at: canvasMenuPoint(owner, at, instance.snapshot().selection.kind === 'text'),
      kind,
      keyboard: at === null,
      owner,
    };
    opened++;
    const request: ContextMenuRequest = {
      ...state,
      close: () => {
        if (menu.value === request) menu.value = null;
      },
    };
    menu.value = request;
  };
  const onContextmenu = (e: MouseEvent): void => {
    e.preventDefault();
    // A Menu key press that also fires `contextmenu` opens the menu once.
    if (Date.now() - keyboardAt < 500) return;
    const pointer = e.clientX !== 0 || e.clientY !== 0;
    open(e.currentTarget as HTMLElement, pointer ? { x: e.clientX, y: e.clientY } : null);
  };
  const onKeydown = (e: KeyboardEvent): void => {
    if (!isContextMenuKey(e)) return;
    e.preventDefault();
    e.stopPropagation();
    keyboardAt = Date.now();
    open(e.currentTarget as HTMLElement, null);
  };
  const layer = (): VNodeChild => {
    const target = menu.value;
    const renderer = entry();
    if (!target || renderer === false) return null;
    return h(
      ContextMenuTargetProvider,
      { key: opened, value: target },
      {
        default: () => (typeof renderer === 'function' ? renderer({}) : h(PptxEditorContextMenu)),
      }
    );
  };
  return { onContextmenu, onKeydown, layer };
}
