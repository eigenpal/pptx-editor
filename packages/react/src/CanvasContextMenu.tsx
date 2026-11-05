/** @jsxImportSource react */
// The slide canvas's context menu request. The engine surface stops the browser's own menu and
// selects what is under the pointer first (`surface-context-menu.ts`); this hook reads the
// selection, opens a menu request for it (shapes, text being edited, or empty slide area), and
// renders the `popups.contextMenu` entry (or the entry of a `PptxEditor.ContextMenu` placed in
// the root): the packaged `PptxEditor.ContextMenu` by default, a
// host renderer, or nothing for `false`. The rows live in `context-menu-parts.tsx`.

import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
} from 'react';
import {
  canvasMenuKind,
  canvasMenuPoint,
  isContextMenuKey,
  type CanvasMenuKind,
} from '@pptx-editor.dev/core/editor';
import type { ContextMenuPoint } from './ContextMenu.tsx';
import { usePptxEditor } from './context.ts';
import {
  ContextMenuTargetProvider,
  PptxEditorContextMenu,
  type ContextMenuRequest,
} from './context-menu-parts.tsx';
import { useCanvasMenuEntry } from './popups.tsx';

interface CanvasMenuState {
  readonly at: ContextMenuPoint;
  readonly kind: CanvasMenuKind;
  readonly keyboard: boolean;
  readonly owner: HTMLElement;
}

/**
 * The canvas context menu. Returns the handlers the canvas host binds (right-click, Shift+F10,
 * and the Menu key) and the layer that renders the open menu.
 */
export function useCanvasContextMenu(): {
  readonly onContextMenu: (e: ReactMouseEvent<HTMLElement>) => void;
  readonly onKeyDown: (e: ReactKeyboardEvent<HTMLElement>) => void;
  readonly layer: ReactNode;
} {
  const editor = usePptxEditor();
  const entry = useCanvasMenuEntry();
  const [menu, setMenu] = useState<CanvasMenuState | null>(null);
  // An entry that turns `false` (a placed menu turned `disabled`) closes the open menu.
  useEffect(() => {
    if (entry === false) setMenu(null);
  }, [entry]);
  const keyboardAt = useRef(0);
  const open = (owner: HTMLElement, at: ContextMenuPoint | null): void => {
    if (entry === false || !editor || editor.snapshot().status !== 'ready') return;
    const kind = canvasMenuKind(editor.snapshot());
    // From the keyboard: open at the selection, or at the middle of the canvas.
    const point = canvasMenuPoint(owner, at);
    setMenu({ at: point, kind, keyboard: at === null, owner });
  };
  const onContextMenu = (e: ReactMouseEvent<HTMLElement>): void => {
    e.preventDefault();
    // A Menu key press that also fires `contextmenu` opens the menu once.
    if (Date.now() - keyboardAt.current < 500) return;
    const pointer = e.clientX !== 0 || e.clientY !== 0;
    open(e.currentTarget, pointer ? { x: e.clientX, y: e.clientY } : null);
  };
  const onKeyDown = (e: ReactKeyboardEvent<HTMLElement>): void => {
    if (!isContextMenuKey(e)) return;
    e.preventDefault();
    e.stopPropagation();
    keyboardAt.current = Date.now();
    open(e.currentTarget, null);
  };
  const target = useMemo<ContextMenuRequest | null>(
    () => (menu ? { ...menu, close: () => setMenu(null) } : null),
    [menu]
  );
  const layer =
    target && entry !== false ? (
      <ContextMenuTargetProvider value={target}>
        {typeof entry === 'function' ? entry({}) : <PptxEditorContextMenu />}
      </ContextMenuTargetProvider>
    ) : null;
  return { onContextMenu, onKeyDown, layer };
}
