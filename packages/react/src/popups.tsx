/** @jsxImportSource react */
// The popups registry, after DOCX Editor's (docx-editor/packages/react/src/editor/popup-config.tsx,
// packages/react/src/editor/popup-renderer.ts, packages/react/src/editor/dialog-host.tsx).
// `PptxEditor.Root` mounts the editor's dialogs. A host changes one through `popups`: a render
// callback replaces the packaged dialog, and `false` turns off automatic rendering. The operation
// behind a dialog never depends on it. DOCX Editor commit adf0e178.
//
// Every packaged opener (the Format panel, the rail and canvas menus) and any host chrome open
// a dialog with `useEditorDialog(key).open()`, so a host renderer receives the request from all
// of them. As in DOCX Editor's dialog host, focus returns to the opener when the dialog closes,
// and a dialog closes when the editor opens another document.

import {
  createContext,
  createElement,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ComponentType,
  type ReactNode,
} from 'react';
import { BackgroundDialog, type BackgroundDialogProps } from './BackgroundDialog.tsx';
import { ImageAltTextDialog, type ImageAltTextDialogProps } from './ImageAltTextDialog.tsx';
import {
  ImagePropertiesDialog,
  type ImagePropertiesDialogProps,
} from './ImagePropertiesDialog.tsx';
import { PageSetupDialog, type PageSetupDialogProps } from './PageSetupDialog.tsx';
import { ParagraphDialog, type ParagraphDialogProps } from './ParagraphDialog.tsx';
import { FontDialog, type FontDialogProps } from './FontDialog.tsx';
import {
  BulletsAndNumberingDialog,
  type BulletsAndNumberingDialogProps,
} from './BulletsAndNumberingDialog.tsx';
import { ColumnsDialog, type ColumnsDialogProps } from './ColumnsDialog.tsx';
import { EquationDialog, type EquationDialogProps } from './EquationDialog.tsx';
import { markReadingPopup } from '@pptx-editor.dev/core/editor';
import { usePptxEditor } from './context.ts';
import type { PptxEditorContextMenuProps } from './context-menu-parts.tsx';
import type { RailContextMenuProps } from './RailContextMenu.tsx';
import type { SectionContextMenuProps } from './RailSections.tsx';

/** What a replacement link popup receives. @public */
export interface HyperlinkPopupProps {
  /**
   * The element the popup opens under: the toolbar's link button, else the element the request
   * came from (the slide for Ctrl/Cmd+K, a menu row's opener).
   */
  readonly anchor: HTMLElement | null;
  /** The link of the selection now, or an empty string. */
  readonly url: string;
  /**
   * `reading` after a click into linked text (the target, Copy, Edit, Unlink), `editing` for the
   * URL field. Absent means `editing`. `useHyperlinkPopup()` gives the state and the actions. A
   * renderer receives `reading` only when registered with `definePopup(C, { reading: true })`.
   */
  readonly mode?: 'reading' | 'editing';
  /** Close the popup. Apply the link with `editor.exec({ type: 'setLink', url })`. */
  readonly onClose: () => void;
  /** A class name added to the popup panel, for a host renderer that forwards these props. */
  readonly className?: string;
}

/** What a replacement find bar receives. @public */
export interface FindPopupProps {
  /** Opened to replace (Ctrl/Cmd+H), not only to find. */
  readonly replace: boolean;
  /** Close the bar. Run `editor.find` and `editor.replace` for the search itself. */
  readonly onClose: () => void;
}

/**
 * A render callback for an automatic popup. `false` disables automatic rendering. Same shape
 * as DOCX Editor's `DocxEditorPopup`. @public
 */
export type PptxEditorPopup<Props extends object> = false | ((props: Props) => ReactNode);

/**
 * Register an ordinary component in the `popups` map while keeping its props and hooks.
 * Define the component outside the parent render function to keep its state across renders.
 *
 * @example
 * ```tsx
 * function AcmePageSetup({ open, onClose }: PageSetupDialogProps) {
 *   return open ? <AcmeDialog title="Slide size" onDismiss={onClose} /> : null;
 * }
 *
 * <PptxEditor.Root popups={{ pageSetup: definePopup(AcmePageSetup) }}>…</PptxEditor.Root>
 * ```
 *
 * @public
 */
export function definePopup<Props extends object>(
  component: ComponentType<Props>,
  options?: DefinePopupOptions
): (props: Props) => ReactNode {
  const render = (props: Props): ReactNode => createElement(component, props);
  return options?.reading ? markReadingPopup(render) : render;
}

/** Options of `definePopup`. @public */
export interface DefinePopupOptions {
  /**
   * For `popups.hyperlink`: the renderer shows reading mode too, so a click into linked text
   * opens it with `mode: 'reading'`. Without it a renderer opens for the URL field alone.
   */
  readonly reading?: boolean;
}

/**
 * Render overrides for the dialogs the editor mounts. `false` disables one. A render callback
 * receives the session props (`open`, `onClose`) and can forward them to the packaged dialog
 * with its own options.
 *
 * @example
 * ```tsx
 * const popups: PptxEditorPopups = {
 *   pageSetup: (props) => <PageSetupDialog {...props} theme={theme} onThemeChange={setTheme} />,
 *   background: false,
 * };
 * ```
 *
 * @public
 */
export interface PptxEditorPopups {
  /** File > Page setup: slide size, appearance, and fonts. */
  pageSetup?: PptxEditorPopup<PageSetupDialogProps>;
  /** Format background: slide background fill. */
  background?: PptxEditorPopup<BackgroundDialogProps>;
  /** Format > Paragraph: alignment, indents, spacing before and after, and line spacing. */
  paragraph?: PptxEditorPopup<ParagraphDialogProps>;
  /**
   * The link popup (Ctrl/Cmd+K, the toolbar's link control, Insert > Link). `PptxEditor.Root`
   * hosts it, so it opens with no toolbar on the page. `false` turns links off in the chrome:
   * the link control and rows show as disabled; `setLink` keeps working from the API.
   */
  hyperlink?: PptxEditorPopup<HyperlinkPopupProps>;
  /**
   * The find and replace bar (Ctrl/Cmd+F, Ctrl/Cmd+H), rendered where `PptxEditor.FindBar`
   * stands. `false` turns it off; `editor.find` keeps working from the API.
   */
  find?: PptxEditorPopup<FindPopupProps>;
  /**
   * The slide canvas's context menu. A callback usually renders `PptxEditor.ContextMenu` with
   * host rows; `false` shows no menu. A `PptxEditor.ContextMenu` placed in the root or the
   * viewport gives the same entry, and this entry wins over it.
   */
  contextMenu?: PptxEditorPopup<PptxEditorContextMenuProps>;
  /**
   * The context menu of a slide in the slide rail and of a tile in the slide sorter. A callback
   * renders the host's menu, or `RailContextMenu` with its own options; `false` shows no menu.
   */
  railContextMenu?: PptxEditorPopup<RailContextMenuProps>;
  /** Format > Font: font, effects, offset, character spacing, and kerning. */
  font?: PptxEditorPopup<FontDialogProps>;
  /** Format > Bullets and Numbering: bullet or number style, size, color, and start number. */
  bulletsAndNumbering?: PptxEditorPopup<BulletsAndNumberingDialogProps>;
  /** Home > Columns > More Columns: the number of text columns and the spacing between them. */
  columns?: PptxEditorPopup<ColumnsDialogProps>;
  /**
   * The Alt Text dialog of the selected element (Picture Format and the context menu), as DOCX
   * Editor's `imageAltText`. `false` turns it off; `setAltText` keeps working from the API.
   */
  imageAltText?: PptxEditorPopup<ImageAltTextDialogProps>;
  /**
   * The Format Picture dialog: size, rotation, and crop position of the selected picture, as
   * DOCX Editor's `imageProperties`. `false` turns it off; `setImageProperties` keeps working.
   */
  imageProperties?: PptxEditorPopup<ImagePropertiesDialogProps>;
  /**
   * Insert > Equation: a gallery and the linear form of an equation, which inserts an equation
   * or edits the selected one. `false` turns it off; `insertEquation` and `setEquation` keep
   * working from the API.
   */
  equation?: PptxEditorPopup<EquationDialogProps>;
  /**
   * The menu of a section header in the slide rail and in the slide sorter (Rename, Remove,
   * Move, Collapse all). A callback renders the host's menu, or `SectionContextMenu` with its
   * own options; `false` shows no menu. The section commands keep working from the API.
   */
  sectionContextMenu?: PptxEditorPopup<SectionContextMenuProps>;
}

/** The modal dialogs the registry mounts and `useEditorDialog` opens. @public */
export type PptxEditorDialogKey =
  | 'pageSetup'
  | 'background'
  | 'paragraph'
  | 'font'
  | 'bulletsAndNumbering'
  | 'columns'
  | 'imageAltText'
  | 'imageProperties'
  | 'equation';

/** What `useEditorDialog` returns. @public */
export interface UseEditorDialogReturn {
  /**
   * Show the dialog. Focus returns to `returnFocusTo`, or to the element that has focus now,
   * when it closes. Does nothing when the popup entry is `false`.
   */
  open(returnFocusTo?: HTMLElement | null): void;
  /** Hide the dialog. */
  close(): void;
  /** Whether the dialog shows now. */
  readonly isOpen: boolean;
  /** False when the popup entry is `false`: hide or disable the control that opens it. */
  readonly isAvailable: boolean;
}

/** The canvas menu entry that a `PptxEditor.ContextMenu` placed in the root gives. */
type CanvasMenuEntry = PptxEditorPopup<PptxEditorContextMenuProps>;

/**
 * The `PptxEditor.ContextMenu` placed in the root or the viewport: it gives the canvas its menu
 * entry, as a `popups.contextMenu` entry does. The last one placed counts. @internal
 */
export interface CanvasMenuRegistry {
  get(): CanvasMenuEntry | undefined;
  place(token: object, entry: CanvasMenuEntry): void;
  remove(token: object): void;
  subscribe(listener: () => void): () => void;
}

function createCanvasMenuRegistry(): CanvasMenuRegistry {
  let current: { readonly token: object; readonly entry: CanvasMenuEntry } | null = null;
  const listeners = new Set<() => void>();
  const notify = (): void => {
    for (const listener of listeners) listener();
  };
  return {
    get: () => current?.entry,
    place(token, entry) {
      if (current?.token === token && current.entry === entry) return;
      current = { token, entry };
      notify();
    },
    remove(token) {
      if (current?.token !== token) return;
      current = null;
      notify();
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => void listeners.delete(listener);
    },
  };
}

interface DialogHost {
  readonly active: PptxEditorDialogKey | null;
  readonly popups: PptxEditorPopups | undefined;
  readonly canvasMenus: CanvasMenuRegistry;
  open(key: PptxEditorDialogKey, returnFocusTo?: HTMLElement | null): void;
  close(): void;
}

const DialogHostContext = createContext<DialogHost | null>(null);

/** The element that should take focus back, or null when it is gone or is the page itself. */
function focusTarget(el: Element | null | undefined): HTMLElement | null {
  // Check against the element's own window: `instanceof HTMLElement` is false across windows.
  const view = el?.ownerDocument.defaultView;
  if (!el || !view || !(el instanceof view.HTMLElement) || el === el.ownerDocument.body)
    return null;
  // A menu row unmounts with its menu: focus goes back to the editor's slide instead.
  if (el.closest('.pptx-menu'))
    return el.closest('.pptx-editor')?.querySelector<HTMLElement>('.pptx-surface__input') ?? null;
  return el;
}

/** Hosts the registry for one editor root, and renders the dialog that is open. */
export function DialogProvider({
  popups,
  children,
}: {
  popups?: PptxEditorPopups;
  children?: ReactNode;
}) {
  const editor = usePptxEditor();
  const [active, setActive] = useState<PptxEditorDialogKey | null>(null);
  const popupsRef = useRef(popups);
  popupsRef.current = popups;
  const opener = useRef<HTMLElement | null>(null);
  const [canvasMenus] = useState(createCanvasMenuRegistry);

  const open = useCallback((key: PptxEditorDialogKey, returnFocusTo?: HTMLElement | null) => {
    if (popupsRef.current?.[key] === false) return;
    opener.current =
      returnFocusTo === undefined
        ? focusTarget(typeof document === 'undefined' ? null : document.activeElement)
        : returnFocusTo;
    setActive(key);
  }, []);
  const close = useCallback(() => setActive(null), []);

  // Focus returns to the opener after the dialog unmounts, as in DOCX Editor's dialog host.
  useEffect(() => {
    if (active !== null) return;
    const target = opener.current;
    opener.current = null;
    if (target?.isConnected) target.focus({ preventScroll: true });
  }, [active]);

  // A dialog belongs to the document it opened on: close it when another document opens.
  useEffect(() => {
    if (!editor) return;
    let session = editor.session;
    return editor.subscribe(() => {
      if (editor.session === session) return;
      session = editor.session;
      setActive(null);
    });
  }, [editor]);

  // A popup entry that becomes `false` closes its open dialog.
  useEffect(() => {
    if (active !== null && popups?.[active] === false) setActive(null);
  }, [popups, active]);

  const host = useMemo<DialogHost>(
    () => ({ active, popups, canvasMenus, open, close }),
    [active, popups, canvasMenus, open, close]
  );
  return (
    <DialogHostContext.Provider value={host}>
      {children}
      <ActiveDialog host={host} />
    </DialogHostContext.Provider>
  );
}

/** One registry entry: the host's renderer, nothing for `false`, else the packaged dialog. */
function renderEntry<Props extends object>(
  entry: PptxEditorPopup<Props> | undefined,
  Packaged: ComponentType<Props>,
  props: Props
): ReactNode {
  if (entry === false) return null;
  return entry ? entry(props) : createElement(Packaged, props);
}

function ActiveDialog({ host }: { host: DialogHost }): ReactNode {
  const { active, popups, close } = host;
  const props = { open: true, onClose: close };
  if (active === 'pageSetup') return renderEntry(popups?.pageSetup, PageSetupDialog, props);
  if (active === 'background') return renderEntry(popups?.background, BackgroundDialog, props);
  if (active === 'paragraph') return renderEntry(popups?.paragraph, ParagraphDialog, props);
  if (active === 'font') return renderEntry(popups?.font, FontDialog, props);
  if (active === 'bulletsAndNumbering')
    return renderEntry(popups?.bulletsAndNumbering, BulletsAndNumberingDialog, props);
  if (active === 'columns') return renderEntry(popups?.columns, ColumnsDialog, props);
  if (active === 'imageAltText')
    return renderEntry(popups?.imageAltText, ImageAltTextDialog, props);
  if (active === 'imageProperties')
    return renderEntry(popups?.imageProperties, ImagePropertiesDialog, props);
  if (active === 'equation') return renderEntry(popups?.equation, EquationDialog, props);
  return null;
}

/** Open a modal registry dialog from a menu row or a toolbar control. @internal */
export function useDialogOpener(): (
  key: PptxEditorDialogKey,
  returnFocusTo?: HTMLElement | null
) => void {
  const open = useContext(DialogHostContext)?.open;
  return useCallback(
    (key: PptxEditorDialogKey, returnFocusTo?: HTMLElement | null) => open?.(key, returnFocusTo),
    [open]
  );
}

/**
 * The registry entry for `key` in the enclosing `PptxEditor.Root`: a renderer, `false`, or
 * undefined for the packaged part. The parts that own a non-modal popup (the link control, the
 * find bar, the canvas, the rail) read it. @internal
 */
export function usePopupEntry<K extends keyof PptxEditorPopups>(key: K): PptxEditorPopups[K] {
  return useContext(DialogHostContext)?.popups?.[key];
}

const noCanvasMenus = (): (() => void) => () => {};
const noCanvasMenu = (): undefined => undefined;

/**
 * The canvas menu entry of the enclosing `PptxEditor.Root`: its `popups.contextMenu` entry,
 * else the entry of a `PptxEditor.ContextMenu` placed in the root, else undefined for the
 * packaged menu. An explicit `popups.contextMenu` entry wins. @internal
 */
export function useCanvasMenuEntry(): CanvasMenuEntry | undefined {
  const host = useContext(DialogHostContext);
  const registry = host?.canvasMenus;
  const placed = useSyncExternalStore(
    registry?.subscribe ?? noCanvasMenus,
    registry?.get ?? noCanvasMenu,
    registry?.get ?? noCanvasMenu
  );
  const explicit = host?.popups?.contextMenu;
  return explicit !== undefined ? explicit : placed;
}

/** Where a `PptxEditor.ContextMenu` placed in the root registers, or null outside a root. @internal */
export function useCanvasMenuRegistry(): CanvasMenuRegistry | null {
  return useContext(DialogHostContext)?.canvasMenus ?? null;
}

/** The whole `popups` map of the enclosing `PptxEditor.Root`. @internal */
export function usePopups(): PptxEditorPopups | undefined {
  return useContext(DialogHostContext)?.popups;
}

/**
 * Open, close, and read one of the editor's dialogs. The packaged openers use the same call,
 * so a host renderer in `popups` receives every request. Outside `PptxEditor.Root` it reports
 * the dialog as unavailable. It takes the dialog's key; the per-dialog hooks
 * (`usePageSetupDialog`, `useBackgroundDialog`, `useParagraphDialog`, `useFontDialog`,
 * `useBulletsAndNumberingDialog`, `useColumnsDialog`) read an open dialog's draft, as in DOCX Editor.
 *
 * @example
 * ```tsx
 * function PageSetupButton() {
 *   const pageSetup = useEditorDialog('pageSetup');
 *   if (!pageSetup.isAvailable) return null;
 *   return <button onClick={() => pageSetup.open()}>Page setup</button>;
 * }
 * ```
 *
 * @public
 */
export function useEditorDialog(key: PptxEditorDialogKey): UseEditorDialogReturn {
  const host = useContext(DialogHostContext);
  const isOpen = host?.active === key;
  const isAvailable = host !== null && host.popups?.[key] !== false;
  const open = host?.open;
  const close = host?.close;
  return useMemo(
    () => ({
      open: (returnFocusTo?: HTMLElement | null) => open?.(key, returnFocusTo),
      close: () => close?.(),
      isOpen,
      isAvailable,
    }),
    [open, close, key, isOpen, isAvailable]
  );
}
