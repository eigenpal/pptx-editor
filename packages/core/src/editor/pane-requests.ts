// Requests to show a tab of the side panel: the toolbar and ribbon slots `arrange.selectionPane`
// (Selection pane) and the Format pane slots send them. The host owns the side panel, so the
// side panel and the host listen: the panel turns to the tab, and the host shows the panel.

// This module imports nothing: the refusal scan of the automation hosts follows imports.

/** The tabs of the side panel, in order. @internal */
export const SIDE_PANEL_TABS = ['format', 'selection', 'comments'] as const;

/** A tab of the side panel: Format, Selection (the selection pane), or Comments. @public */
export type SidePanelTab = (typeof SIDE_PANEL_TABS)[number];

type PaneListener = (pane: SidePanelTab) => void;

/** For each editor: the functions that answer a pane request. */
const listeners = new WeakMap<object, Set<PaneListener>>();

/**
 * Listen for requests to show a side panel tab for `editor`. A mounted side panel listens to
 * turn to the tab. Listen in your host too, to show the side panel when it is closed. Returns
 * the function that stops listening. @public
 */
export function listenPaneRequests(editor: object, open: (pane: SidePanelTab) => void): () => void {
  let set = listeners.get(editor);
  if (!set) listeners.set(editor, (set = new Set()));
  const own: PaneListener = (pane) => open(pane);
  set.add(own);
  return () => void set.delete(own);
}

/**
 * Ask to show the side panel tab `pane` for `editor`. Returns false when nothing listens: no side
 * panel is mounted and the host does not answer. @public
 */
export function requestPane(editor: object, pane: SidePanelTab): boolean {
  const set = listeners.get(editor);
  if (!set || set.size === 0) return false;
  for (const listener of [...set]) listener(pane);
  return true;
}
