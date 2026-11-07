// Requests to show a tab of the side panel: the toolbar and ribbon slots `arrange.selectionPane`
// (Selection pane) and `shape.properties` (Format shape) send them. The host owns the side
// panel, so the side panel and the host listen: the panel turns to the tab, and the host shows
// the panel. While nothing listens, the slots are off with the `disabled.noSidePanel` reason.
//
// This module imports nothing: the refusal scan of the automation hosts follows imports.

/** The tabs of the side panel, in order. @internal */
export const SIDE_PANEL_TABS = ['format', 'selection', 'comments'] as const;

/** A tab of the side panel: Format, Selection (the selection pane), or Comments. @public */
export type SidePanelTab = (typeof SIDE_PANEL_TABS)[number];

type PaneListener = (pane: SidePanelTab) => void;

/** For each editor: the functions that answer a pane request. */
const listeners = new WeakMap<object, Set<PaneListener>>();
/** Who follows the listeners coming and going: the toolbar slots read them. */
const watchers = new Set<() => void>();
let version = 0;
const changed = (): void => {
  version += 1;
  for (const watch of [...watchers]) watch();
};

/**
 * Listen for requests to show a side panel tab for `editor`. A mounted side panel listens and
 * turns to the tab. Listen in your host too, to show the side panel when it is closed. While at
 * least one listener is registered, the toolbar's Selection pane and Format shape controls are
 * on. Returns the function that stops listening. @public
 */
export function listenPaneRequests(editor: object, open: (pane: SidePanelTab) => void): () => void {
  let set = listeners.get(editor);
  if (!set) listeners.set(editor, (set = new Set()));
  const own: PaneListener = (pane) => open(pane);
  set.add(own);
  changed();
  return () => {
    if (set.delete(own)) changed();
  };
}

/**
 * Ask to show the side panel tab `pane` for `editor`. Every listener hears it; the host decides
 * what an answer means. Returns false when nothing listens. A listener that throws does not stop
 * the others; the first error is thrown after all of them ran. @public
 */
export function requestPane(editor: object, pane: SidePanelTab): boolean {
  const set = listeners.get(editor);
  if (!set || set.size === 0) return false;
  let failure: { error: unknown } | null = null;
  for (const listener of [...set]) {
    try {
      listener(pane);
    } catch (error) {
      failure ??= { error };
    }
  }
  if (failure) throw failure.error;
  return true;
}

/** Whether a side panel or a host listens for pane requests of `editor`. @internal */
export function hasPaneListener(editor: object | null | undefined): boolean {
  return !!editor && !!listeners.get(editor)?.size;
}

/** Follow listeners that start or stop, for any editor. Returns the undo. @internal */
export function subscribePaneListeners(watch: () => void): () => void {
  watchers.add(watch);
  return () => void watchers.delete(watch);
}

/** A number that changes when a listener starts or stops. @internal */
export function paneListenersVersion(): number {
  return version;
}
