// Requests to show a tab of the side panel: the toolbar and ribbon slots `arrange.selectionPane`
// (Selection pane) and `shape.properties` (Format shape) send them. The host owns the side
// panel, so the side panel and the host listen: the panel turns to the tab, and the host shows
// the panel. While nothing listens, the slots are off with the `disabled.noSidePanel` reason.
// The comment draft requests at the end follow the same pattern.
//
// This module imports nothing: the refusal scan of the automation hosts follows imports.

/** The tabs of the side panel, in order. @internal */
export const SIDE_PANEL_TABS = [
  'format',
  'chart',
  'diagram',
  'selection',
  'animations',
  'comments',
] as const;

/**
 * A tab of the side panel: Format, Chart Data, Diagram (the text of the diagram's shapes),
 * Selection (the selection pane), Animations, or Comments. @public
 */
export type SidePanelTab = (typeof SIDE_PANEL_TABS)[number];

/**
 * The icon of each side panel tab (a name in `ICON_PATHS`). The tab row shows the icons, with the
 * label of the selected tab beside its icon, so five tabs fit the panel in every locale; each tab
 * keeps its full label as its accessible name and tooltip. @internal
 */
export const SIDE_PANEL_TAB_ICONS: Readonly<Record<SidePanelTab, string>> = {
  format: 'tune',
  chart: 'insert_chart',
  diagram: 'edit_note',
  selection: 'list_alt',
  animations: 'play_arrow',
  comments: 'comment',
};

/** How a side panel answers a request: whether it moves the focus to the tab. @public */
export interface PaneRequest {
  /**
   * False keeps the focus where it is: Insert > Chart opens the Chart tab, and the keyboard
   * stays on the slide, so Undo and Delete act on the new chart.
   */
  readonly focus: boolean;
}

type PaneListener = (pane: SidePanelTab, request: PaneRequest) => void;

/** For each editor: the functions that answer a pane request. */
const listeners = new WeakMap<object, Set<PaneListener>>();
/** Who follows the listeners coming and going: the toolbar slots read them. */
const watchers = new Set<() => void>();
/** The last request of each editor, until a side panel takes it (`takePaneRequest`). */
const recent = new WeakMap<
  object,
  { readonly pane: SidePanelTab; readonly focus: boolean; readonly at: number }
>();
/** A side panel that mounts this soon after a request answers it: the host just opened it. */
const RECENT_MS = 1000;
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
export function listenPaneRequests(
  editor: object,
  open: (pane: SidePanelTab, request: PaneRequest) => void
): () => void {
  let set = listeners.get(editor);
  if (!set) listeners.set(editor, (set = new Set()));
  const own: PaneListener = (pane, request) => open(pane, request);
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
export function requestPane(
  editor: object,
  pane: SidePanelTab,
  options: { readonly focus?: boolean } = {}
): boolean {
  const set = listeners.get(editor);
  if (!set || set.size === 0) return false;
  const request: PaneRequest = { focus: options.focus !== false };
  recent.set(editor, { pane, focus: request.focus, at: Date.now() });
  let failure: { error: unknown } | null = null;
  for (const listener of [...set]) {
    try {
      listener(pane, request);
    } catch (error) {
      failure ??= { error };
    }
  }
  if (failure) throw failure.error;
  return true;
}

/**
 * After a command ran: a new chart opens the Chart tab, where its sample data waits to be
 * replaced. The focus stays on the slide, so Undo, Delete, and the arrow keys act on the new
 * chart. Returns `result`. @internal
 */
export function afterInsertCommand<R extends { readonly ok: boolean }>(
  editor: object,
  command: { readonly type: string },
  result: R
): R {
  if (result.ok && command.type === 'insertChart') requestPane(editor, 'chart', { focus: false });
  return result;
}

/**
 * The tab of a request that no side panel has taken yet, made in the last second, and whether
 * it moves the focus, or null. A side panel takes it when it turns to a tab, and when it mounts:
 * a host that shows a closed panel for a request mounts the panel after the request. @internal
 */
export function takePaneRequest(
  editor: object | null | undefined
): { readonly pane: SidePanelTab; readonly focus: boolean } | null {
  if (!editor) return null;
  const last = recent.get(editor);
  recent.delete(editor);
  return last && Date.now() - last.at < RECENT_MS ? { pane: last.pane, focus: last.focus } : null;
}

/**
 * Whether a slot is a pane slot, which follows the pane listeners and the comment draft
 * openers (`review.addComment`). @internal
 */
export function isPaneSlot(slot: string): boolean {
  return (
    slot === 'arrange.selectionPane' ||
    slot === 'shape.properties' ||
    slot === 'chart.editData' ||
    slot === 'diagram.editText' ||
    slot === 'review.addComment'
  );
}

/**
 * The tabs the side panel shows, in order: Chart Data only while a chart is selected, and
 * Diagram only while a diagram is selected, or while the tab is open (it then says to select
 * one). @internal
 */
export function visibleSidePanelTabs(
  chartSelected: boolean,
  open: SidePanelTab,
  diagramSelected = false
): readonly SidePanelTab[] {
  return SIDE_PANEL_TABS.filter(
    (tab) =>
      (tab !== 'chart' || chartSelected || open === 'chart') &&
      (tab !== 'diagram' || diagramSelected || open === 'diagram')
  );
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

// A new comment draft: the New comment row of the canvas context menus (`review.addComment`)
// asks a comments pane to open one. A mounted pane registers its opener. A request while no pane
// is mounted asks for the Comments tab and waits, as a pane request does, for the pane that a
// side panel or the host then mounts.

/** For each editor: the openers of the mounted comments panes, the latest last. */
const draftOpeners = new WeakMap<object, (() => void)[]>();
/** The time of each editor's draft request that no pane has taken yet. */
const waitingDrafts = new WeakMap<object, number>();

/**
 * Register the function that opens a new comment draft for `editor`. A mounted comments pane
 * registers one; the New comment row of the canvas context menus calls the latest. A request
 * made at most a second before, while no pane was mounted, reaches the opener now: the host
 * showed the pane for it. Returns the function that unregisters. @public
 */
export function registerCommentDraft(editor: object, open: () => void): () => void {
  let list = draftOpeners.get(editor);
  if (!list) draftOpeners.set(editor, (list = []));
  const own = (): void => open();
  list.push(own);
  changed();
  const at = waitingDrafts.get(editor);
  if (at !== undefined) {
    waitingDrafts.delete(editor);
    if (Date.now() - at < RECENT_MS) own();
  }
  return () => {
    const i = list.indexOf(own);
    if (i < 0) return;
    list.splice(i, 1);
    changed();
  };
}

/**
 * Whether a draft request can reach a comments pane: one is mounted, or a side panel or host
 * listens for pane requests and can show the **Comments** tab. @internal
 */
export function canRequestCommentDraft(editor: object | null | undefined): boolean {
  return !!editor && (!!draftOpeners.get(editor)?.length || hasPaneListener(editor));
}

/**
 * Ask for a new comment draft. A mounted comments pane opens it. Otherwise the request asks for
 * the **Comments** tab (`requestPane`), and the pane that mounts within a second opens it.
 * Returns false when nothing can answer. @public
 */
export function requestCommentDraft(editor: object): boolean {
  const open = draftOpeners.get(editor)?.at(-1);
  if (open) {
    open();
    return true;
  }
  if (!hasPaneListener(editor)) return false;
  waitingDrafts.set(editor, Date.now());
  return requestPane(editor, 'comments');
}
