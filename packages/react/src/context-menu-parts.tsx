/** @jsxImportSource react */
// `PptxEditor.ContextMenu`: the slide canvas's context menu as library parts, after DOCX
// Editor's `DocxEditor.ContextMenu` (docx-editor/docs/site/content/customization/menus.mdx,
// packages/react/src/editor/contextmenu/).
//
// The canvas opens a menu request (right-click, Shift+F10, the Menu key) for what is under the
// pointer: shapes, text being edited, or empty slide area. The menu renders core's
// `CANVAS_CONTEXT_MENUS` rows for that kind. `ContextMenu.Slot slotId` replaces or hides one
// row in place, other children add rows at the end, and `preset={false}` keeps only the
// children. The `popups.contextMenu` registry entry decides what the canvas renders.

import {
  Children,
  Fragment,
  createContext,
  isValidElement,
  useContext,
  type ReactNode,
} from 'react';
import {
  CANVAS_CONTEXT_MENUS,
  arrangeMenuRows,
  canvasMenuLabelKey,
  warnMenuPart,
  type CanvasMenuKind,
} from '@pptx-editor.dev/core/editor';
import {
  ContextMenu,
  MenuCloseContext,
  MenuSeparator,
  type ContextMenuPoint,
} from './ContextMenu.tsx';
import { useTranslation } from './context.ts';
import { MenuEntries } from './menu-rows.tsx';
import { MenuSlotItem, type MenuSlotItemProps } from './menu-parts.tsx';

/** What a canvas context menu opened on, and where. @public */
export interface ContextMenuRequest {
  /**
   * Shapes, pictures only (`picture`, the picture menu), text being edited, or empty
   * slide area. A selection of pictures only reports `picture`, not `shapes`.
   */
  readonly kind: CanvasMenuKind;
  /** Where the menu opens, in client pixels. */
  readonly at: ContextMenuPoint;
  /** Opened from the keyboard: focus goes to the first row. */
  readonly keyboard: boolean;
  /** The canvas element the menu opened on. */
  readonly owner: HTMLElement;
  /** Close the menu. */
  readonly close: () => void;
}

const TargetContext = createContext<ContextMenuRequest | null>(null);

/** Provides the open menu request to `PptxEditor.ContextMenu`. @internal */
export const ContextMenuTargetProvider = TargetContext.Provider;

/**
 * The open canvas context menu, for a host's own rows: what it opened on, where, and how to
 * close it. Null outside an open menu. @public
 */
export function useContextMenuRequest(): ContextMenuRequest | null {
  return useContext(TargetContext);
}

/**
 * The element the open context menu opened on (the slide canvas), or null while no menu is
 * open. The same hook as DOCX Editor's. @public
 */
export function useContextMenuTarget(): HTMLElement | null {
  return useContext(TargetContext)?.owner ?? null;
}

/** Props for `PptxEditor.ContextMenu`. @public */
export interface PptxEditorContextMenuProps {
  /** A class name added to the menu. */
  readonly className?: string;
  /**
   * `true` (default): the built-in rows for the target, with `ContextMenu.Slot` children
   * replacing their rows in place and other children added at the end. `false`: only the
   * children.
   */
  readonly preset?: boolean;
  /** `ContextMenu.Slot` children change built-in rows; other children add rows. */
  readonly children?: ReactNode;
}

type NodeProps = Record<string, unknown> & { children?: ReactNode };

const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/**
 * The slide canvas's context menu. It renders while the canvas has a menu request open: put it
 * in the `popups.contextMenu` entry to change the rows.
 *
 * @example
 * ```tsx
 * const popups: PptxEditorPopups = {
 *   contextMenu: (props) => (
 *     <PptxEditor.ContextMenu {...props}>
 *       <PptxEditor.ContextMenu.Slot slotId="format.painter" hidden />
 *       <PptxEditor.ContextMenu.Item label="Ask legal" icon="info" onSelect={askLegal} />
 *     </PptxEditor.ContextMenu>
 *   ),
 * };
 * ```
 *
 * @public
 */
export function PptxEditorContextMenu({
  className,
  preset = true,
  children,
}: PptxEditorContextMenuProps): ReactNode {
  const { t } = useTranslation();
  const target = useContextMenuRequest();
  if (!target) {
    warnMenuPart('PptxEditor.ContextMenu', 'outsideContextMenu');
    return null;
  }
  let content: ReactNode = children;
  if (preset) {
    // A hidden row for a row this target lacks (a text row on a shape) is no extra row.
    const {
      entries: shown,
      overrides,
      extras,
    } = arrangeMenuRows(CANVAS_CONTEXT_MENUS[target.kind], flatten(children), (child) =>
      isValidElement<NodeProps>(child) && child.type === MenuSlotItem
        ? (child.props as unknown as MenuSlotItemProps)
        : null
    );
    content = (
      <>
        <MenuEntries entries={shown} overrides={overrides} />
        {extras.length ? <MenuSeparator /> : null}
        {extras}
      </>
    );
  }
  const label = t(canvasMenuLabelKey(target.kind));
  return (
    <ContextMenu
      at={target.at}
      onClose={target.close}
      label={label}
      keyboard={target.keyboard}
      owner={target.owner}
      {...(className ? { className } : {})}
    >
      <MenuCloseContext.Provider value={target.close}>{content}</MenuCloseContext.Provider>
    </ContextMenu>
  );
}
