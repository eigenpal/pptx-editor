/** @jsxImportSource react */
// The React side of core's `ChromeActionHost`: how a menu row or a toolbar control opens a
// registry dialog, starts the slide show, opens the find bar or the link popup, and asks for a
// picture. Core decides which of these a slot needs (`runChromeRow`); this module only connects
// them to the adapter. `PptxEditor.Root` publishes one host and renders the link popup, as DOCX
// Editor's root hosts its hyperlink dialog, so every opener works with no toolbar on the page.

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useId,
  useMemo,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from 'react';
import {
  SLOT_POPUP_KEYS,
  createHyperlinkPopupStore,
  focusSurfaceOf,
  hyperlinkEntryReads,
  linkPopupAnnouncement,
  listenLinkClicks,
  listenLinkPanelKey,
  linkPopupReturnsFocus,
  pickImageInto,
  requestLinkPopup,
  type ChromeActionHost,
  type ChromeSlotId,
  type HyperlinkPopupStore,
} from '@pptx-editor.dev/core/editor';
import { ChromeActionHostContext, usePptxEditor, useSlideshow, useTranslation } from './context.ts';
import { requestFindBar } from './FindBar.tsx';
import { HyperLink, HyperlinkPopupContext as ControllerContext } from './HyperLink.tsx';
import { useDialogOpener, usePopupEntry, usePopups, type PptxEditorPopups } from './popups.tsx';
import {
  HyperlinkPopupContext,
  HyperlinkPopupViewsContext,
  useHyperlinkPopupBinding,
} from './useHyperlinkPopup.ts';

// Ask the `PptxEditor.Root` of an editor to open its link popup, as Ctrl/Cmd+K does. The
// request and the anchor rules are core's (`overlays.ts`), shared with the Vue adapter.
export { requestLinkPopup };

/** The reason key a `false` popups entry gives the slots it turns off. */
const OFF_REASONS: Readonly<Record<string, string>> = {
  hyperlink: 'link.unavailable',
  find: 'find.unavailable',
  pageSetup: 'dialog.unavailable',
  background: 'dialog.unavailable',
  imageAltText: 'dialog.unavailable',
  imageProperties: 'dialog.unavailable',
  equation: 'dialog.unavailable',
  paragraph: 'dialog.unavailable',
  font: 'dialog.unavailable',
  bulletsAndNumbering: 'dialog.unavailable',
  columns: 'dialog.unavailable',
};

/**
 * Why `slot` cannot run with these `popups`, or null. With `hyperlink: false`, a host panel
 * that reads `useHyperlinkPopup()` (`linkViews`) still opens for the link controls.
 */
function popupReason(
  popups: PptxEditorPopups | undefined,
  slot: ChromeSlotId,
  dialog: string | undefined,
  linkViews: number
): string | null {
  const key = dialog ?? SLOT_POPUP_KEYS[slot];
  if (!key || popups?.[key as keyof PptxEditorPopups] !== false) return null;
  if (key === 'hyperlink' && linkViews > 0) return null;
  return OFF_REASONS[key] ?? null;
}

/** How many find bars a root has mounted, so Find shows as disabled while there is none. */
const FindBarsContext = createContext<{
  readonly count: number;
  readonly register: () => () => void;
} | null>(null);

/** Count `PptxEditor.FindBar` in its root while it is mounted. @internal */
export function useRegisterFindBar(): void {
  const register = useContext(FindBarsContext)?.register;
  useEffect(() => register?.(), [register]);
}

/** The host this root's own contexts give, never an enclosing root's. */
function useLocalChromeActionHost(findBars: number | null, linkViews = 0): ChromeActionHost {
  const editor = usePptxEditor();
  const openDialog = useDialogOpener();
  const show = useSlideshow();
  const popups = usePopups();
  return useMemo<ChromeActionHost>(
    () => ({
      openDialog,
      startSlideshow: (from) => show.start(from),
      openFind: (replace) => requestFindBar(editor, replace),
      openLink: () => requestLinkPopup(editor),
      pickImage: (target) => {
        if (editor) pickImageInto(editor, document, target);
      },
      unavailable: (slot, dialog) =>
        popupReason(popups, slot, dialog, linkViews) ??
        // Find shows its matches on the slide canvas, which the slide sorter does not show.
        (SLOT_POPUP_KEYS[slot] === 'find' && editor?.snapshot().presentationView === 'slideSorter'
          ? 'disabled.findSorter'
          : null) ??
        // Find needs a find bar to open: nothing answers the request without one.
        (findBars === 0 && SLOT_POPUP_KEYS[slot] === 'find' ? 'find.noBar' : null),
    }),
    [editor, openDialog, show, popups, findBars, linkViews]
  );
}

/** The adapter actions behind dialog, clipboard, show, find, link, and picture slots. */
export function useChromeActionHost(): ChromeActionHost {
  const published = useContext(ChromeActionHostContext);
  const local = useLocalChromeActionHost(null);
  return published ?? local;
}

/**
 * Publishes the root's action host, so `useEditorCommand` runs adapter slots. A root inside
 * another root publishes its own host. @internal
 */
export function ChromeActionHostProvider({ children }: { children?: ReactNode }) {
  const [findBars, setFindBars] = useState(0);
  const finds = useMemo(
    () => ({
      count: findBars,
      register: () => {
        setFindBars((n) => n + 1);
        return () => setFindBars((n) => n - 1);
      },
    }),
    [findBars]
  );
  // The root's one link popup state, which the packaged popup and `useHyperlinkPopup` share.
  const [linkStore] = useState(createHyperlinkPopupStore);
  const [linkViews, setLinkViews] = useState(0);
  const attachLinkView = useCallback(() => {
    setLinkViews((n) => n + 1);
    let attached = true;
    return () => {
      if (!attached) return;
      attached = false;
      setLinkViews((n) => n - 1);
    };
  }, []);
  const host = useLocalChromeActionHost(findBars, linkViews);
  return (
    <FindBarsContext.Provider value={finds}>
      <ChromeActionHostContext.Provider value={host}>
        <HyperlinkPopupContext.Provider value={linkStore}>
          <HyperlinkPopupViewsContext.Provider value={attachLinkView}>
            <LinkPopupBridge store={linkStore}>
              {children}
              <LinkPopupHost store={linkStore} views={linkViews} />
            </LinkPopupBridge>
          </HyperlinkPopupViewsContext.Provider>
        </HyperlinkPopupContext.Provider>
      </ChromeActionHostContext.Provider>
    </FindBarsContext.Provider>
  );
}

/** Publish the current controller to the compound parts and host-rendered children. */
function LinkPopupBridge({ store, children }: { store: HyperlinkPopupStore; children: ReactNode }) {
  const state = useSyncExternalStore(store.subscribe, store.getState, store.getState);
  const errorId = useId();
  const controller = store.controller();
  return (
    <ControllerContext.Provider value={controller ? { controller, state, errorId } : null}>
      {children}
    </ControllerContext.Provider>
  );
}

/** The root's shared popup, including reading gestures and accessible announcements. */
function LinkPopupHost({ store, views }: { store: HyperlinkPopupStore; views: number }): ReactNode {
  const editor = usePptxEditor();
  const entry = usePopupEntry('hyperlink');
  const { t } = useTranslation();
  useHyperlinkPopupBinding(store, true);
  const state = useSyncExternalStore(store.subscribe, store.getState, store.getState);
  const controller = store.controller();
  const off = entry === false;
  const reads = hyperlinkEntryReads(entry);
  useEffect(() => {
    if (off || !reads || !editor || !controller) return;
    const stopClicks = listenLinkClicks(editor, (at) => controller.openReading(at));
    const stopKey = listenLinkPanelKey(editor, () => controller.getState().mode === 'reading');
    return () => {
      stopClicks();
      stopKey();
    };
  }, [editor, controller, off, reads]);
  useEffect(() => {
    if (off && views === 0) store.close();
  }, [off, views, store]);
  // Check focus synchronously while the closing panel is still mounted.
  useEffect(() => {
    if (!controller || !editor) return;
    let open = controller.getState().mode !== 'closed';
    return controller.subscribe(() => {
      const now = controller.getState().mode !== 'closed';
      if (open && !now && linkPopupReturnsFocus(editor)) focusSurfaceOf(editor);
      open = now;
    });
  }, [controller, editor]);
  if (!controller) return null;
  let content: ReactNode = null;
  if (state.mode !== 'closed' && !off)
    content =
      typeof entry === 'function' ? (
        entry({ anchor: state.anchor, url: state.url, mode: state.mode, onClose: store.close })
      ) : (
        <HyperLink />
      );
  return (
    <>
      {content}
      <span className="pptx-sr-only" role="status" aria-live="polite">
        {linkPopupAnnouncement(state, t, editor?.snapshot().customShows)}
      </span>
    </>
  );
}
