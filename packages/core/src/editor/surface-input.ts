// DOM input wiring for the slide surface: the `beforeinput` dispatch, IME composition, keys,
// and clipboard events. Browser input becomes editor text input or commands; the DOM is never
// edited. Every `beforeinput` is prevented, and an input type the editor does not apply is
// refused with a typed reason, never dropped silently. Adapted from docx-editor's
// `surface-input.ts`.

import type { EditorCommand, EditorExecOptions, HistoryGroup } from '../contracts/index.ts';
import type { TextInputKind } from './text-input.ts';
import { bindEditorClipboard } from './editor-clipboard.ts';

/** What the input wiring asks of the editor. */
export interface SurfaceInputHost {
  textInput(kind: TextInputKind, data?: string): void;
  shortcut(key: string, event: KeyboardEvent): boolean;
  /**
   * Fill a copy or cut; false leaves the event to the browser. `slideList`: the event comes
   * from the slide list, and copies the selected slides.
   */
  copy(data: DataTransfer, cut: boolean, slideList?: boolean): boolean;
  /** Paste from a paste event; false leaves the event to the browser. */
  paste(data: DataTransfer): boolean;
  /** A copy may come next (Cmd or Ctrl went down, a menu opened): get ready for it. */
  warmCopy?(): void;
  /**
   * Run a command the browser asked for (Edit menu undo, format shortcuts). The steps of one
   * pointer gesture pass one `historyGroup`.
   */
  run(command: EditorCommand, options?: EditorExecOptions): void;
  /** Open a history group for the steps of one pointer gesture. */
  beginHistoryGroup(): HistoryGroup;
  /** Publish why an input was not applied. */
  refuse(reason: string, detail?: string): void;
}

/** Input types that map to one text input. */
const TEXT_INPUT: Readonly<Record<string, TextInputKind>> = {
  deleteContentBackward: 'deleteBackward',
  deleteContentForward: 'deleteForward',
  deleteWordBackward: 'deleteWordBackward',
  deleteWordForward: 'deleteWordForward',
  // Cmd+Backspace on macOS: the word before the caret, as Alt+Backspace.
  deleteSoftLineBackward: 'deleteWordBackward',
  deleteHardLineBackward: 'deleteWordBackward',
  deleteSoftLineForward: 'deleteLineForward',
  deleteHardLineForward: 'deleteLineForward',
  insertParagraph: 'paragraph',
  insertLineBreak: 'lineBreak',
};

/** Input types that carry text to insert, on the event or in its data transfer. */
const INSERTS = new Set([
  'insertText',
  'insertReplacementText',
  'insertFromPaste',
  'insertFromPasteAsQuotation',
  'insertFromDrop',
]);

/** Input types that map to a command. */
const COMMANDS: Readonly<Record<string, EditorCommand>> = {
  historyUndo: { type: 'undo' },
  historyRedo: { type: 'redo' },
  formatBold: { type: 'toggleBold' },
  formatItalic: { type: 'toggleItalic' },
  formatUnderline: { type: 'toggleUnderline' },
  formatStrikeThrough: { type: 'toggleStrike' },
};

/**
 * Input types of an IME composition. The composition's result arrives once, at composition
 * end; a browser that reports one of these outside a composition (a late commit event) is
 * ignored, so the text never goes in twice.
 */
const COMPOSITION_INPUT = new Set([
  'insertCompositionText',
  'deleteCompositionText',
  'insertFromComposition',
  'deleteByComposition',
]);

/** Apply one `beforeinput` event. Exported for tests. */
export function dispatchBeforeInput(host: SurfaceInputHost, e: InputEvent): void {
  const type = e.inputType;
  if (COMPOSITION_INPUT.has(type)) return;
  const kind = Object.prototype.hasOwnProperty.call(TEXT_INPUT, type) ? TEXT_INPUT[type] : null;
  if (kind) return host.textInput(kind);
  if (INSERTS.has(type)) {
    const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
    if (text) host.textInput('text', text);
    return;
  }
  if (Object.prototype.hasOwnProperty.call(COMMANDS, type)) return host.run(COMMANDS[type]);
  host.refuse('errors.inputUnsupported', type);
}

/** The class the input takes while an IME composes: it shows the text being composed. */
export const COMPOSING_CLASS = 'pptx-surface__input--composing';

/**
 * Whether a key belongs to an IME composition: during one, and the key that ends one (some
 * browsers send the Enter that commits a composition after `compositionend`, with key code
 * 229). Such a key never runs an editor shortcut, so Enter does not also split the paragraph.
 */
export function isCompositionKey(e: KeyboardEvent): boolean {
  return e.isComposing || e.keyCode === 229;
}

/** Listen for input on the surface. Returns the function that stops listening. */
export function bindSurfaceInput(
  input: HTMLTextAreaElement,
  root: HTMLElement,
  host: SurfaceInputHost
): () => void {
  let composing = false;
  // The input grows with the text being composed, so the text shows at the caret while the
  // IME composes it (the slide shows the text only once the composition ends).
  const fit = (): void => {
    input.style.width = '';
    if (composing && input.value) input.style.width = `${input.scrollWidth + 1}px`;
  };
  const endComposition = (): void => {
    composing = false;
    input.classList.remove(COMPOSING_CLASS);
    input.value = '';
    fit();
  };
  const onBeforeInput = (e: InputEvent): void => {
    // The IME owns the field until it finishes; the result arrives at composition end.
    if (composing) return;
    e.preventDefault();
    dispatchBeforeInput(host, e);
    input.value = '';
  };
  const onInput = (): void => {
    if (composing) fit();
  };
  const onCompositionStart = (): void => {
    composing = true;
    input.classList.add(COMPOSING_CLASS);
  };
  const onCompositionEnd = (e: CompositionEvent): void => {
    endComposition();
    if (e.data) host.textInput('text', e.data);
  };
  const onKeyDown = (e: KeyboardEvent): void => {
    if (isCompositionKey(e)) return;
    if ((e.key === 'Meta' || e.key === 'Control') && !e.repeat) host.warmCopy?.();
    if (host.shortcut(e.key, e)) {
      e.preventDefault();
      e.stopPropagation();
    }
  };
  const onCopy = (e: ClipboardEvent): void => {
    if (e.clipboardData && host.copy(e.clipboardData, false)) e.preventDefault();
  };
  const onCut = (e: ClipboardEvent): void => {
    if (e.clipboardData && host.copy(e.clipboardData, true)) e.preventDefault();
  };
  const onPaste = (e: ClipboardEvent): void => {
    if (e.clipboardData && host.paste(e.clipboardData)) {
      e.preventDefault();
      e.stopPropagation();
    }
  };
  input.addEventListener('beforeinput', onBeforeInput);
  input.addEventListener('input', onInput);
  input.addEventListener('compositionstart', onCompositionStart);
  input.addEventListener('compositionend', onCompositionEnd);
  input.addEventListener('keydown', onKeyDown);
  root.addEventListener('keydown', onKeyDown);
  root.addEventListener('copy', onCopy);
  root.addEventListener('cut', onCut);
  root.addEventListener('paste', onPaste);
  const onContextMenu = (): void => host.warmCopy?.();
  root.addEventListener('contextmenu', onContextMenu);
  // Copy and paste with the focus elsewhere in the editor (the slide list, the toolbar).
  const unbindOutside = bindEditorClipboard(root, host);
  return () => {
    unbindOutside();
    root.removeEventListener('contextmenu', onContextMenu);
    input.removeEventListener('beforeinput', onBeforeInput);
    input.removeEventListener('input', onInput);
    input.removeEventListener('compositionstart', onCompositionStart);
    input.removeEventListener('compositionend', onCompositionEnd);
    input.removeEventListener('keydown', onKeyDown);
    root.removeEventListener('keydown', onKeyDown);
    root.removeEventListener('copy', onCopy);
    root.removeEventListener('cut', onCut);
    root.removeEventListener('paste', onPaste);
  };
}
