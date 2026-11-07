/** @jsxImportSource react */
// `PptxEditor.ContextMenu`: the slide canvas's context menu as library parts, after DOCX
// Editor's `DocxEditor.ContextMenu` (docx-editor/docs/site/content/customization/menus.mdx,
// packages/react/src/editor/contextmenu/).
//
// The canvas opens a menu request (right-click, Shift+F10, the Menu key) for what is under the
// pointer: shapes, text being edited, or empty slide area. The menu renders core's
// `CANVAS_CONTEXT_MENUS` rows for that kind. `ContextMenu.Slot slotId` replaces or hides one
// row in place, other children add rows at the end, and `preset={false}` keeps only the
// children. The `popups.contextMenu` registry entry, or a menu placed in the root, decides
// what the canvas renders.

import type { ChromeTranslate } from '@pptx-editor.dev/core/editor';
import {
  Children,
  Fragment,
  createContext,
  isValidElement,
  useContext,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  type ReactNode,
} from 'react';
import {
  CANVAS_CONTEXT_MENUS,
  CONTEXT_MENU_NAMED_PARTS,
  arrangeMenuRows,
  canvasMenuLabelKey,
  canvasMenuRows,
  warnMenuPart,
  type CanvasMenuKind,
} from '@pptx-editor.dev/core/editor';
import {
  ContextMenu,
  MenuCloseContext,
  MenuSeparator,
  type ContextMenuPoint,
} from './ContextMenu.tsx';
import { TranslationContext, usePptxEditor, useTranslation } from './context.ts';
import { ToolbarContext } from './headless-hooks.tsx';
import { useCanvasMenuRegistry, type PptxEditorPopups } from './popups.tsx';
import { MenuEntries } from './menu-rows.tsx';
import { defineMenuRowParts, menuSlotRowOf } from './menu-parts.tsx';

/** The named rows of the canvas context menu (`ContextMenu.Cut`, …), from core's table. */
export const CONTEXT_MENU_ROW_PARTS = defineMenuRowParts(CONTEXT_MENU_NAMED_PARTS);

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
  /**
   * Your translation function for this compound and its parts, as DOCX Editor's `t`. Without
   * it, the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /** A class name added to the menu. */
  readonly className?: string;
  /**
   * `true` (default): the built-in rows for the target, with `ContextMenu.Slot` children
   * replacing their rows in place and other children added at the end. `false`: only the
   * children.
   */
  readonly preset?: boolean;
  /**
   * `true` shows no canvas menu, as DOCX Editor's `disabled`. An open menu closes. Default
   * `false`.
   */
  readonly disabled?: boolean;
  /** Called with `true` when the menu opens and `false` when it closes. */
  readonly onOpenChange?: (open: boolean) => void;
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
 * The slide canvas's context menu. Place it in `PptxEditor.Root` or `PptxEditor.Viewport`, as
 * DOCX Editor's `DocxEditor.ContextMenu`, or render it from the `popups.contextMenu` entry. A
 * placed menu gives the canvas the same entry, and an explicit `popups.contextMenu` entry wins.
 *
 * @example
 * ```tsx
 * <PptxEditor.Root>
 *   <PptxEditor.Viewport>
 *     <PptxEditor.Content />
 *     <PptxEditor.ContextMenu onOpenChange={setMenuOpen}>
 *       <PptxEditor.ContextMenu.Slot slotId="format.painter" hidden />
 *       <PptxEditor.ContextMenu.Item label="Ask legal" icon="info" onSelect={askLegal} />
 *     </PptxEditor.ContextMenu>
 *   </PptxEditor.Viewport>
 * </PptxEditor.Root>
 * ```
 *
 * @public
 */
export function PptxEditorContextMenu(props: PptxEditorContextMenuProps): ReactNode {
  const target = useContextMenuRequest();
  const registry = useCanvasMenuRegistry();
  const placed = !target && registry !== null;
  usePlacedCanvasMenu(placed ? registry : null, props);
  useOpenChange(target && !props.disabled ? target : null, props);
  if (target) return <OpenContextMenu target={target} {...props} />;
  if (!placed) warnMenuPart('PptxEditor.ContextMenu', 'outsideContextMenu');
  return null;
}

/**
 * A menu placed in the root gives the canvas its entry: the same menu, with the `t` and the
 * locale of where it was placed, or `false` while it is `disabled`.
 */
function usePlacedCanvasMenu(
  registry: ReturnType<typeof useCanvasMenuRegistry>,
  props: PptxEditorContextMenuProps
): void {
  const token = useRef({}).current;
  const host = useContext(ToolbarContext);
  const translate = useContext(TranslationContext);
  const { disabled, className, preset, onOpenChange, children } = props;
  const entry = useMemo(
    () =>
      disabled
        ? (false as const)
        : () => (
            <ToolbarContext.Provider value={host}>
              <TranslationContext.Provider value={translate}>
                <PptxEditorContextMenu
                  {...(className !== undefined ? { className } : {})}
                  {...(preset !== undefined ? { preset } : {})}
                  {...(onOpenChange ? { onOpenChange } : {})}
                >
                  {children}
                </PptxEditorContextMenu>
              </TranslationContext.Provider>
            </ToolbarContext.Provider>
          ),
    [disabled, className, preset, onOpenChange, children, host, translate]
  );
  useLayoutEffect(() => {
    if (registry) registry.place(token, entry);
  }, [registry, token, entry]);
  useLayoutEffect(() => (registry ? () => registry.remove(token) : undefined), [registry, token]);
}

/** `onOpenChange` hears each open and close of the menu, once. */
function useOpenChange(
  target: ContextMenuRequest | null,
  { onOpenChange }: PptxEditorContextMenuProps
): void {
  const handler = useRef(onOpenChange);
  handler.current = onOpenChange;
  const open = target !== null;
  useEffect(() => {
    if (!open) return undefined;
    handler.current?.(true);
    return () => handler.current?.(false);
  }, [open]);
}

function OpenContextMenu({
  target,
  className,
  preset = true,
  disabled,
  children,
}: PptxEditorContextMenuProps & { readonly target: ContextMenuRequest }): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  // `disabled` turning on closes an open menu.
  useEffect(() => {
    if (disabled) target.close();
  }, [disabled, target]);
  if (disabled) return null;
  let content: ReactNode = children;
  if (preset) {
    // A hidden row for a row this target lacks (a text row on a shape) is no extra row.
    const {
      entries: shown,
      overrides,
      extras,
    } = arrangeMenuRows(CANVAS_CONTEXT_MENUS[target.kind], flatten(children), menuSlotRowOf);
    content = (
      <>
        {/* New comment shows only while a comments pane can open a draft. */}
        <MenuEntries entries={canvasMenuRows(shown, editor)} overrides={overrides} />
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

/**
 * The `popups` the sugar passes to the root: the host's entries, with `contextMenu` and
 * `hyperlinkPopup` mapped onto the same entries. A host entry wins. @internal
 */
export function sugarPopups(
  popups: PptxEditorPopups | undefined,
  hyperlinkPopup: boolean | undefined,
  contextMenu: boolean | PptxEditorContextMenuProps | undefined,
  /** The menu the entry renders: `PptxEditor.ContextMenu`. */
  Menu: (props: PptxEditorContextMenuProps) => ReactNode
): PptxEditorPopups | undefined {
  const linkOff = hyperlinkPopup === false && popups?.hyperlink === undefined;
  const menuEntry =
    popups?.contextMenu !== undefined || contextMenu === undefined || contextMenu === true
      ? undefined
      : contextMenu === false
        ? (false as const)
        : (props: PptxEditorContextMenuProps) => <Menu {...props} {...contextMenu} />;
  if (!linkOff && menuEntry === undefined) return popups;
  return {
    ...popups,
    ...(linkOff ? { hyperlink: false as const } : {}),
    ...(menuEntry !== undefined ? { contextMenu: menuEntry } : {}),
  };
}
