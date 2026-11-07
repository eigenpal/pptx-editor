// Copy, cut, and paste while the keyboard focus is in the editor but not on the slide: on a
// slide in the slide list after a click there, or on a toolbar button. The slide surface
// handles its own clipboard events (`surface-input.ts`); these events never reach it, so
// they are taken here, at the document, and given to the same handlers. A paste after a click
// on another slide puts the copied shapes on that slide.
//
// In the slide list, Copy and Cut take the selected slides, and Paste puts copied slides after
// the current slide.
//
// Events from a field (a dialog's input, the notes, the find bar, a section name), a dialog, or a
// menu keep their browser behavior, and so does an event outside the editor's own element.

/** What the binding asks of the editor: the surface's own clipboard handlers. */
export interface EditorClipboardHost {
  copy(data: DataTransfer, cut: boolean, slideList?: boolean): boolean;
  paste(data: DataTransfer, slideList?: boolean, chrome?: boolean): boolean;
  warmCopy?(): void;
}

/** The slide list and the slide sorter. */
const SLIDE_LISTS = '.pptx-rail, .pptx-sorter';

/** The element an editor's chrome and slide share. */
const EDITOR_ROOT = '.pptx-editor';

/**
 * Whether an element takes text input of its own, or is in a dialog or menu (a slide show, a
 * popover, a link editor): a copy or paste there is not the slide's.
 */
function isField(el: Element): boolean {
  return !!el.closest(
    'input, textarea, select, [contenteditable=""], [contenteditable="true"], [role="textbox"], ' +
      'dialog, [role="dialog"], [role="alertdialog"], [role="menu"]'
  );
}

/**
 * Whether a clipboard event is this editor's to handle here: its target is inside the editor's
 * element, outside the slide surface, and not a field.
 */
export function isEditorChromeEvent(surface: HTMLElement, target: EventTarget | null): boolean {
  const el = target instanceof Element ? target : null;
  if (!el || surface.contains(el)) return false;
  const editor = surface.closest(EDITOR_ROOT);
  if (!editor?.contains(el)) return false;
  // A nested editor's element is that editor's.
  if (el.closest(EDITOR_ROOT) !== editor) return false;
  return !isField(el);
}

/**
 * Listen for clipboard events in the rest of the editor around `surface`. Returns the function
 * that stops listening.
 */
export function bindEditorClipboard(surface: HTMLElement, host: EditorClipboardHost): () => void {
  const doc = surface.ownerDocument;
  const onClipboard = (e: ClipboardEvent): void => {
    if (e.defaultPrevented || !e.clipboardData || !isEditorChromeEvent(surface, e.target)) return;
    // In the slide list and the slide sorter, Copy and Cut take the selected slides, and Paste
    // puts in copied slides also while text on the current slide is edited.
    const slideList = e.target instanceof Element && !!e.target.closest(SLIDE_LISTS);
    const done =
      e.type === 'paste'
        ? host.paste(e.clipboardData, slideList, true)
        : host.copy(e.clipboardData, e.type === 'cut', slideList);
    if (done) e.preventDefault();
  };
  const onKeyDown = (e: KeyboardEvent): void => {
    if (
      (e.key === 'Meta' || e.key === 'Control') &&
      !e.repeat &&
      isEditorChromeEvent(surface, e.target)
    )
      host.warmCopy?.();
  };
  doc.addEventListener('copy', onClipboard);
  doc.addEventListener('cut', onClipboard);
  doc.addEventListener('paste', onClipboard);
  doc.addEventListener('keydown', onKeyDown);
  return () => {
    doc.removeEventListener('copy', onClipboard);
    doc.removeEventListener('cut', onClipboard);
    doc.removeEventListener('paste', onClipboard);
    doc.removeEventListener('keydown', onKeyDown);
  };
}
