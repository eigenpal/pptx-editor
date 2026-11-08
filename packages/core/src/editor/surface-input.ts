// DOM input wiring for the slide surface: the `beforeinput` dispatch, IME composition, keys,
// and clipboard events. Browser input becomes editor text input or commands; the DOM is never
// edited. Every `beforeinput` is prevented, and an input type the editor does not apply is
// refused with a typed reason, never dropped silently. After DOCX Editor
// `packages/core/src/editor/surface-input.ts` at e568148f8; the `insertReplacementText`
// handling (`replacementTarget`) is adapted from that file at f189d6503 (EigenPal).
//
// The input holds the caret's paragraph up to the caret (`TypingContext`), so the system's
// own text changes (capitals, quotes, dashes, spelling, dictation) see what the user typed.
// The surface sets no `autocorrect`, `autocapitalize`, or `spellcheck` on it, so the user's
// system settings apply. The input is transparent (`opacity: 0`), so a spelling mark on it
// never shows; while an IME composes, it shows only the composed text, scrolled to the caret.

import type { EditorCommand, EditorExecOptions, HistoryGroup } from '../contracts/index.ts';
import type { TextInputKind } from './text-input.ts';
import { bindEditorClipboard } from './editor-clipboard.ts';
import { compositionEdit, type ReplacementTarget, type TypingContext } from './replacement-text.ts';

/** What the input wiring asks of the editor. */
export interface SurfaceInputHost {
  /** Apply a text input; `exact`: a system replacement, typed with no AutoCorrect. */
  textInput(kind: TextInputKind, data?: string, exact?: boolean): void;
  /**
   * Apply a system text change: `replacement` replaces `target` when the model text before
   * the caret holds it, else goes in at the caret (`replacement-text.ts`).
   */
  replaceText?(target: ReplacementTarget | null, replacement: string): void;
  /** The text the input holds for the system, or null for none. */
  typingContext?(): TypingContext | null;
  shortcut(key: string, event: KeyboardEvent): boolean;
  /**
   * Fill a copy or cut; false leaves the event to the browser. `slideList`: the event comes
   * from the slide list, and copies the selected slides.
   */
  copy(data: DataTransfer, cut: boolean, slideList?: boolean): boolean;
  /**
   * Paste from a paste event; false leaves the event to the browser. `slideList`: the event
   * comes from the slide list; `chrome`: from outside the slide (`editor-clipboard.ts`).
   */
  paste(data: DataTransfer, slideList?: boolean, chrome?: boolean): boolean;
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

/** The text one static range covers, or null when it cannot be read. */
function staticRangeText(range: StaticRange): string | null {
  const { startContainer, endContainer } = range;
  if (startContainer === endContainer && startContainer.nodeType === 3 /* text */)
    return (startContainer.textContent ?? '').slice(range.startOffset, range.endOffset);
  const doc = startContainer.ownerDocument;
  if (!doc) return null;
  try {
    const live = doc.createRange();
    live.setStart(startContainer, range.startOffset);
    live.setEnd(endContainer, range.endOffset);
    return live.toString();
  } catch {
    return null;
  }
}

/**
 * The target of a replacement event, or null when it names none the model could act on. In
 * the input, the system selects the text it replaces over the mirrored `context`; a browser
 * that reports a target range instead gives its text, with no paragraph. One range only: a
 * real substitution targets one run of text, and several joined could match by chance.
 */
export function replacementTarget(
  e: InputEvent,
  input: { readonly value: string; readonly selectionStart: number; readonly selectionEnd: number },
  context: TypingContext | null
): ReplacementTarget | null {
  const { value, selectionStart: start, selectionEnd: end } = input;
  if (context && value === context.text && end > start && end <= value.length)
    return { text: value.slice(start, end), after: value.length - end, paragraph: context.key };
  const ranges = typeof e.getTargetRanges === 'function' ? e.getTargetRanges() : [];
  if (ranges.length !== 1) return null;
  const text = staticRangeText(ranges[0]!);
  return text ? { text, after: 0, paragraph: null } : null;
}

/** Apply one `beforeinput` event. Exported for tests. */
export function dispatchBeforeInput(
  host: SurfaceInputHost,
  e: InputEvent,
  target: ReplacementTarget | null = null
): void {
  const type = e.inputType;
  if (COMPOSITION_INPUT.has(type)) return;
  const kind = Object.prototype.hasOwnProperty.call(TEXT_INPUT, type) ? TEXT_INPUT[type] : null;
  if (kind) return host.textInput(kind);
  if (type === 'insertReplacementText') {
    // Autocorrect, dictation, and smart substitutions arrive this way, not from a paste.
    const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
    if (!text) return;
    if (host.replaceText) host.replaceText(target, text);
    else host.textInput('text', text);
    return;
  }
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

/** The surface's input binding. */
export interface SurfaceInputBinding {
  /** Put the caret's context in the input (`TypingContext`); none while an IME composes. */
  sync(): void;
  /** Stop listening. */
  destroy(): void;
}

/** Listen for input on the surface. */
export function bindSurfaceInput(
  input: HTMLTextAreaElement,
  root: HTMLElement,
  host: SurfaceInputHost
): SurfaceInputBinding {
  let composing = false;
  /** The context the input holds, and the value and width when a composition started. */
  let context: TypingContext | null = null;
  let composeFrom = { value: '', width: 0 };
  const sync = (): void => {
    if (composing) return;
    context = host.typingContext?.() ?? null;
    const text = context?.text ?? '';
    if (input.value !== text) input.value = text;
    // The caret of the input stays at the end of the context, as the model caret is. A new
    // value puts it there; only a focused input can have moved it.
    const end = text.length;
    const focused = input.ownerDocument.activeElement === input;
    if (focused && (input.selectionStart !== end || input.selectionEnd !== end))
      input.setSelectionRange(end, end);
  };
  // The input grows with the text being composed, so the text shows at the caret while the
  // IME composes it (the slide shows the text only once the composition ends). The context
  // before it scrolls out of view.
  const fit = (): void => {
    input.style.width = '';
    if (!composing || input.value === composeFrom.value) return;
    input.style.width = `${Math.max(1, input.scrollWidth - composeFrom.width) + 1}px`;
    input.scrollLeft = input.scrollWidth;
  };
  const endComposition = (): void => {
    composing = false;
    input.classList.remove(COMPOSING_CLASS);
    fit();
    sync();
  };
  const onBeforeInput = (e: InputEvent): void => {
    // The IME owns the field until it finishes; the result arrives at composition end.
    if (composing) return;
    e.preventDefault();
    const target =
      e.inputType === 'insertReplacementText' ? replacementTarget(e, input, context) : null;
    dispatchBeforeInput(host, e, target);
    sync();
  };
  const onInput = (): void => {
    if (composing) fit();
  };
  const onCompositionStart = (): void => {
    composing = true;
    input.classList.add(COMPOSING_CLASS);
    composeFrom = { value: input.value, width: input.value ? input.scrollWidth : 0 };
  };
  const onCompositionEnd = (e: CompositionEvent): void => {
    const before = composeFrom.value;
    const { value, selectionEnd } = input;
    const key = context?.key ?? null;
    endComposition();
    if (!e.data) return;
    // A keyboard can recompose a word the input already holds: only the change goes in.
    const edit = before ? compositionEdit(before, value, selectionEnd, e.data, key) : null;
    if (!edit || !host.replaceText) return host.textInput('text', e.data);
    if (edit.target.text === '' && edit.target.after === 0)
      return host.textInput('text', edit.text);
    host.replaceText(edit.target, edit.text);
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
  sync();
  const destroy = (): void => {
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
  return { sync, destroy };
}
