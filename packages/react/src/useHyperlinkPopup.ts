// The link popup's behavior with no UI, after DOCX Editor's `useHyperlinkPopup`. The state and
// the actions are core's (`createHyperlinkPopupStore` in `overlays.ts`), shared with the Vue
// adapter; these hooks subscribe to one store and bind it to the editor's link requests.

import {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import {
  createHyperlinkPopupStore,
  hyperlinkPopupActions,
  listenHyperlinkPopup,
  type HyperlinkPopupActions,
  type HyperlinkPopupState,
  type HyperlinkPopupStore,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, usePptxEditor } from './context.ts';
import {
  HyperlinkPopupContext as ControllerContext,
  useHyperlinkPopup as useControllerPopup,
} from './HyperLink.tsx';

/** What `useHyperlinkPopup` returns, as DOCX Editor's hook of the same name. @public */
export interface UseHyperlinkPopupResult extends HyperlinkPopupActions {
  /** Whether the popup shows, what it shows, and the draft. */
  readonly state: HyperlinkPopupState;
}

/** The link popup store of the enclosing `PptxEditor.Root`. @internal */
export const HyperlinkPopupContext = createContext<HyperlinkPopupStore | null>(null);

/**
 * Counts the views of the root's link popup state: a host panel that calls
 * `useHyperlinkPopup()`. With `popups.hyperlink: false`, the link controls work while one is
 * mounted. @internal
 */
export const HyperlinkPopupViewsContext = createContext<(() => () => void) | null>(null);

/** Bind `store` to the editor, and while `active`, to its link requests. @internal */
export function useHyperlinkPopupBinding(store: HyperlinkPopupStore, active: boolean): void {
  const editor = usePptxEditor();
  const enabled = useEditorCommand('text.link').isEnabled;
  const enabledRef = useRef(enabled);
  enabledRef.current = enabled;
  useEffect(() => {
    store.setEditor(editor);
    return () => store.setEditor(null);
  }, [store, editor]);
  useEffect(() => {
    if (!editor || !active) return undefined;
    // Requests for this editor only: another editor on the page opens its own popup.
    return listenHyperlinkPopup(store, editor, () => enabledRef.current);
  }, [store, editor, active]);
}

/** The state and actions of `store`, with a stable identity while the state does not change. */
function useStoreResult(store: HyperlinkPopupStore): UseHyperlinkPopupResult {
  const state = useSyncExternalStore(store.subscribe, store.getState, store.getState);
  const actions = useMemo(() => hyperlinkPopupActions(store), [store]);
  return useMemo(() => ({ state, ...actions }), [state, actions]);
}

/**
 * The link popup's behavior, for your own link UI, as DOCX Editor's `useHyperlinkPopup`.
 *
 * Inside `PptxEditor.Root` it is the shared state that the packaged popup shows, so your panel,
 * the toolbar's link control, Insert > Link, and Ctrl/Cmd+K agree. With
 * `popups.hyperlink: false`, the link controls open your panel while it is mounted. Outside a
 * root it is a standalone instance, as `useHyperlinkPopupInstance()`.
 *
 * @example
 * ```tsx
 * function LinkPanel() {
 *   const { state, setUrl, commitEdit, close } = useHyperlinkPopup();
 *   if (state.mode === 'closed') return null;
 *   return (
 *     <form onSubmit={(e) => { e.preventDefault(); commitEdit(); }}>
 *       <input value={state.draft} onChange={(e) => setUrl(e.target.value)} />
 *       <button type="button" onClick={close}>Cancel</button>
 *     </form>
 *   );
 * }
 * ```
 *
 * @public
 */
export function useHyperlinkPopup(): UseHyperlinkPopupResult {
  const provided = useContext(HyperlinkPopupContext);
  const controllerContext = useContext(ControllerContext);
  const controllerResult = useControllerPopup();
  const attach = useContext(HyperlinkPopupViewsContext);
  useEffect(() => (provided ? attach?.() : undefined), [provided, attach]);
  // Every render calls the same hooks: `provided` only decides whether the own store listens.
  const [own] = useState(createHyperlinkPopupStore);
  useHyperlinkPopupBinding(own, provided === null && controllerContext === null);
  const result = useStoreResult(provided ?? own);
  // A standalone packaged popup can still provide the existing controller context.
  return !provided && controllerContext ? controllerResult : result;
}

/**
 * A link popup of its own, as DOCX Editor's `useHyperlinkPopupInstance`. While `active`, it
 * opens for the link requests of the editor. Inside a root, prefer `useHyperlinkPopup()`: the
 * root's own popup also answers the requests.
 *
 * @public
 */
export function useHyperlinkPopupInstance(active = true): UseHyperlinkPopupResult {
  const [store] = useState(createHyperlinkPopupStore);
  useHyperlinkPopupBinding(store, active);
  return useStoreResult(store);
}
