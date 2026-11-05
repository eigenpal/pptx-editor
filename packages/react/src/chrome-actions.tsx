/** @jsxImportSource react */
// The React side of core's `ChromeActionHost`: how a menu row or a toolbar control opens a
// registry dialog, starts the slide show, opens the find bar or the link popup, and asks for a
// picture. Core decides which of these a slot needs (`runChromeRow`); this module only connects
// them to the adapter. `PptxEditor.Root` publishes one host and renders the link popup, as DOCX
// Editor's root hosts its hyperlink dialog, so every opener works with no toolbar on the page.

import {
  createContext,
  useContext,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from 'react';
import {
  SLOT_POPUP_KEYS,
  createHyperlinkPopupController,
  focusSurfaceOf,
  hyperlinkEntryReads,
  linkPopupAnnouncement,
  listenLinkClicks,
  listenLinkPanelKey,
  linkPopupReturnsFocus,
  listenLinkRequests,
  pickImageInto,
  requestLinkPopup,
  type ChromeActionHost,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import {
  ChromeActionHostContext,
  useEditorCommand,
  usePptxEditor,
  useSlideshow,
  useTranslation,
} from './context.ts';
import { requestFindBar } from './FindBar.tsx';
import { HyperLink, HyperlinkPopupContext } from './HyperLink.tsx';
import { useDialogOpener, usePopupEntry, usePopups, type PptxEditorPopups } from './popups.tsx';

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
  paragraph: 'dialog.unavailable',
  font: 'dialog.unavailable',
  bulletsAndNumbering: 'dialog.unavailable',
  columns: 'dialog.unavailable',
};

/** Why `slot` cannot run with these `popups`, or null. */
function popupReason(
  popups: PptxEditorPopups | undefined,
  slot: ChromeSlotId,
  dialog?: string
): string | null {
  const key = dialog ?? SLOT_POPUP_KEYS[slot];
  if (!key || popups?.[key as keyof PptxEditorPopups] !== false) return null;
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
function useLocalChromeActionHost(findBars: number | null): ChromeActionHost {
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
        popupReason(popups, slot, dialog) ??
        // Find shows its matches on the slide canvas, which the slide sorter does not show.
        (SLOT_POPUP_KEYS[slot] === 'find' && editor?.snapshot().presentationView === 'slideSorter'
          ? 'disabled.findSorter'
          : null) ??
        // Find needs a find bar to open: nothing answers the request without one.
        (findBars === 0 && SLOT_POPUP_KEYS[slot] === 'find' ? 'find.noBar' : null),
    }),
    [editor, openDialog, show, popups, findBars]
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
  const host = useLocalChromeActionHost(findBars);
  return (
    <FindBarsContext.Provider value={finds}>
      <ChromeActionHostContext.Provider value={host}>
        {children}
        <LinkPopupHost />
      </ChromeActionHostContext.Provider>
    </FindBarsContext.Provider>
  );
}

/**
 * The link popup of one root: `popups.hyperlink`, or the packaged `PptxEditor.HyperLink`. It
 * owns the root's one popup state (core's controller), opened in edit mode by every link
 * request and in reading mode by a click into linked text.
 */
function LinkPopupHost(): ReactNode {
  const editor = usePptxEditor();
  const entry = usePopupEntry('hyperlink');
  const cmd = useEditorCommand('text.link');
  const { t } = useTranslation();
  const errorId = useId();
  const controller = useMemo(
    () => (editor ? createHyperlinkPopupController(editor) : null),
    [editor]
  );
  useEffect(() => () => controller?.dispose(), [controller]);
  const state = useSyncExternalStore(
    (notify) => controller?.subscribe(notify) ?? (() => {}),
    () => controller?.getState() ?? null,
    () => controller?.getState() ?? null
  );
  const off = entry === false;
  const enabled = cmd.isEnabled;
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;
  useEffect(() => {
    if (off || !editor || !controller) return;
    // Requests for this root's editor only: another editor on the page opens its own. The
    // anchor is core's `linkPopupAnchor`.
    return listenLinkRequests(
      editor,
      () => enabledRef.current,
      (at) => controller.openEditing(at)
    );
  }, [editor, controller, off]);
  // A click into linked text opens the reading panel under the caret, also for a viewer: it
  // reads, copies, and opens the link, and the panel shows Edit and Unlink only when the link
  // command can run. A host renderer gets reading mode only when it asks for it.
  const reads = hyperlinkEntryReads(entry);
  useEffect(() => {
    if (off || !reads || !editor || !controller) return;
    const stopClicks = listenLinkClicks(editor, (at) => controller.openReading(at));
    // F6 moves focus from the slide text into the open reading panel.
    const stopKey = listenLinkPanelKey(editor, () => controller.getState().mode === 'reading');
    return () => {
      stopClicks();
      stopKey();
    };
  }, [editor, controller, off, reads]);
  // Turning links off closes an open popup.
  useEffect(() => {
    if (off) controller?.close();
  }, [off, controller]);
  // On close, focus goes back to the slide (the text input while text is selected), as DOCX
  // Editor's HyperLink does; focus that moved elsewhere on purpose stays there. Checked at the
  // moment of closing, while focus is still in the panel.
  useEffect(() => {
    if (!controller || !editor) return;
    let open = controller.getState().mode !== 'closed';
    return controller.subscribe(() => {
      const now = controller.getState().mode !== 'closed';
      if (open && !now && linkPopupReturnsFocus(editor)) focusSurfaceOf(editor);
      open = now;
    });
  }, [controller, editor]);
  if (!controller || !state || off) return null;
  const value = { controller, state, errorId };
  // A polite live region, always mounted so a screen reader hears the panel open and the copy.
  const live = (
    <span className="pptx-sr-only" role="status" aria-live="polite">
      {linkPopupAnnouncement(state, t)}
    </span>
  );
  let content: ReactNode = null;
  if (state.mode !== 'closed')
    content =
      typeof entry === 'function' ? (
        entry({
          anchor: state.anchor,
          url: state.url,
          mode: state.mode,
          onClose: () => controller.close(),
        })
      ) : (
        <HyperLink />
      );
  return (
    <HyperlinkPopupContext.Provider value={value}>
      {content}
      {live}
    </HyperlinkPopupContext.Provider>
  );
}
