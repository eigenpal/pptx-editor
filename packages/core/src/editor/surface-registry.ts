// Which element each editor paints into, so chrome outside that element (a menu bar above the
// viewport, a menu in a portal) reaches the right editor: its text input for Cut and Copy, and
// keyboard focus after a menu or popup closes. The focused element cannot tell: two editors on
// a page, or a menu outside `.pptx-editor`.

interface SurfaceEntry {
  readonly root: HTMLElement;
  readonly focus: () => void;
  readonly repaint: () => void;
}

const surfaces = new WeakMap<object, SurfaceEntry>();
const hostListeners = new WeakMap<object, Set<() => void>>();

/** Record (or, with null, forget) the element `editor` paints into. @internal */
export function setSurfaceRoot(
  editor: object,
  element: HTMLElement | null,
  focus: () => void = () => element?.focus({ preventScroll: true }),
  repaint: () => void = () => {}
): void {
  const before = surfaces.get(editor)?.root.ownerDocument ?? null;
  if (element) surfaces.set(editor, { root: element, focus, repaint });
  else surfaces.delete(editor);
  if ((element?.ownerDocument ?? null) === before) return;
  for (const listener of [...(hostListeners.get(editor) ?? [])]) listener();
}

/**
 * The document that `editor`'s slide is mounted in: an iframe's when the host mounted the
 * editor there, else the page's own. Before the slide mounts, the page's document. Null with no
 * DOM, as during server rendering. @public
 */
export function editorHostDocument(editor: object): Document | null {
  const root = surfaces.get(editor)?.root;
  if (root) return root.ownerDocument;
  return typeof document === 'undefined' ? null : document;
}

/**
 * Call `listener` each time `editor`'s slide moves to another document, as DOCX Editor's
 * `hostChange` event. Returns the function that stops. @public
 */
export function onEditorHostChange(editor: object, listener: () => void): () => void {
  let set = hostListeners.get(editor);
  if (!set) hostListeners.set(editor, (set = new Set()));
  set.add(listener);
  return () => {
    set.delete(listener);
  };
}

/** The element `editor` paints into, or null. @internal */
export function surfaceRootOf(editor: object): HTMLElement | null {
  return surfaces.get(editor)?.root ?? null;
}

/** The text input of the surface `editor` paints into, or null. @internal */
export function surfaceInputOf(editor: object): HTMLElement | null {
  return surfaces.get(editor)?.root.querySelector<HTMLElement>('.pptx-surface__input') ?? null;
}

/**
 * Give keyboard focus back to `editor`'s slide: its text input while text is selected, else the
 * slide. False when the editor paints nowhere. @internal
 */
export function focusSurfaceOf(editor: object): boolean {
  const entry = surfaces.get(editor);
  if (!entry?.root.isConnected) return false;
  entry.focus();
  return true;
}

/**
 * Repaint `editor`'s slide with no change event: the adapters call it when the UI locale
 * changes, so the placeholder prompts and frame labels take the new strings. False when the
 * editor paints nowhere. @internal
 */
export function repaintSurfaceOf(editor: object): boolean {
  const entry = surfaces.get(editor);
  if (!entry) return false;
  entry.repaint();
  return true;
}
