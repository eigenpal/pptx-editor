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

/** Record (or, with null, forget) the element `editor` paints into. @internal */
export function setSurfaceRoot(
  editor: object,
  element: HTMLElement | null,
  focus: () => void = () => element?.focus({ preventScroll: true }),
  repaint: () => void = () => {}
): void {
  if (element) surfaces.set(editor, { root: element, focus, repaint });
  else surfaces.delete(editor);
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
