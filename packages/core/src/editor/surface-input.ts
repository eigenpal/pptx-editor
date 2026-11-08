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

import type {
  EditorCommand,
  EditorExecOptions,
  EditorSelection,
  HistoryGroup,
} from '../contracts/index.ts';
import type { TextInputKind } from './text-input.ts';
import { bindEditorClipboard } from './editor-clipboard.ts';
import {
  compositionEdit,
  fieldEdit,
  type ReplacementTarget,
  type TypingContext,
} from './replacement-text.ts';

/** What the input wiring asks of the editor. */
export interface SurfaceInputHost {
  /**
   * Apply a text input. `before`: a system replacement, typed with no AutoCorrect, whose undo
   * gives `before` back as the selection.
   */
  textInput(kind: TextInputKind, data?: string, before?: EditorSelection): void;
  /**
   * Apply a system text change: `replacement` replaces `target` when the model text before
   * the caret holds it, and goes in at the caret when there is no target
   * (`replacement-text.ts`). `typed`: `replacement` holds the new characters of an IME
   * composition; when the target no longer matches, they go in at the caret, so no typed
   * text is lost.
   */
  replaceText?(target: ReplacementTarget | null, replacement: string, typed?: boolean): void;
  /** The text the input holds for the system, or null for none. */
  typingContext?(): TypingContext | null;
  /** The text for a translation key: the input's description. */
  label?(key: string): string;
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

/** The input's value and selection, as `replacementTarget` reads them. */
interface InputField {
  readonly value: string;
  readonly selectionStart: number;
  readonly selectionEnd: number;
}

/** The range `[start, end)` of the input's context as a target, or null for none. */
function contextTarget(
  context: TypingContext | null,
  value: string,
  start: number,
  end: number
): ReplacementTarget | null {
  if (!context || value !== context.text || end <= start || end > value.length) return null;
  const between = value.slice(end);
  return { text: value.slice(start, end), after: between.length, between, paragraph: context.key };
}

/**
 * A selection in the input that does not lie over the synced context: it never matches the
 * model text (`after` is negative), so the editor refuses the input and changes nothing.
 */
const UNUSABLE: ReplacementTarget = { text: '', after: -1, paragraph: null };

/**
 * The text the system selected in the input over the mirrored `context`, or null when the
 * input selection is collapsed. Some systems select the text they correct, then send
 * `insertText`: that text is a target too. The input selection is only the system's own:
 * the binding collapses it before each key (`bindSurfaceInput`). A selection over a value
 * that is not the synced context cannot be used (`UNUSABLE`).
 */
export function inputTarget(input: InputField, context: TypingContext | null) {
  const { value, selectionStart: start, selectionEnd: end } = input;
  if (end <= start) return null;
  return contextTarget(context, value, start, end) ?? UNUSABLE;
}

/**
 * The target of a replacement event, or null when it names none the model could act on: the
 * text the system selected in the input, else one target range. A range in the input's text
 * (its offsets address the input's value) gives its distance to the caret; a range elsewhere
 * gives its text, with no paragraph. One range only: a real substitution targets one run of
 * text, and several joined could match by chance.
 */
export function replacementTarget(
  e: InputEvent,
  input: InputField,
  context: TypingContext | null
): ReplacementTarget | null {
  const selected = inputTarget(input, context);
  if (selected) return selected;
  const ranges = typeof e.getTargetRanges === 'function' ? e.getTargetRanges() : [];
  if (ranges.length !== 1) return null;
  const range = ranges[0]!;
  const node = range.startContainer;
  if (node === range.endContainer && node.nodeType === 3 && node.textContent === input.value) {
    const inContext = contextTarget(context, input.value, range.startOffset, range.endOffset);
    if (inContext) return inContext;
  }
  const text = staticRangeText(range);
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
  // Autocorrect, dictation, and smart substitutions arrive as `insertReplacementText`, or as
  // `insertText` over the text the system selected in the input. Not from a paste.
  if (type === 'insertReplacementText' || (type === 'insertText' && target)) {
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

/**
 * The width of a text in the input's font, in CSS pixels, or 0 where no canvas measures it
 * (then the composing overlay keeps its plain placement).
 */
function measurer(input: HTMLTextAreaElement): (text: string) => number {
  let context: CanvasRenderingContext2D | null | undefined;
  return (text) => {
    try {
      context ??= input.ownerDocument.createElement('canvas').getContext('2d');
      const view = input.ownerDocument.defaultView;
      if (!context || !view) return 0;
      context.font = view.getComputedStyle(input).font;
      return context.measureText(text).width;
    } catch {
      return 0;
    }
  };
}

/** A strong right-to-left letter: Hebrew, Arabic, Syriac, Thaana, N'Ko, and their forms. */
const RTL = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufefc]/u;

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
  // A screen reader reads the input's value: it says that the value is the text up to the
  // caret, not the whole text of the shape.
  const description = host.label?.('canvas.inputDescription');
  if (description) input.setAttribute('aria-description', description);
  // The input follows each caret move at once, not lazily at the next key: a mobile keyboard
  // reads the text before the caret when the caret moves, for its capitals and suggestions,
  // and an empty input reads to it as the start of a sentence.
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
  /**
   * Put the input caret back where `sync` left it, at the end of the value: a native key the
   * editor does not handle can move it or select with it (a line-wise move). The user sees
   * neither on the slide, so no input may start from there. Not during a composition. Every
   * edit that reads the input selection runs after this, unless the system made that
   * selection for the same event.
   */
  const collapse = (): void => {
    if (composing) return;
    const end = input.value.length;
    if (input.selectionStart !== end || input.selectionEnd !== end)
      input.setSelectionRange(end, end);
  };
  /**
   * The input selection counts only for the edit that the system sends with it, in the same
   * task: a system replacement selects its target and sends its `beforeinput` at once. A
   * selection that is still there when `selectionchange` (or `select`) runs, a task later,
   * came from something with no edit after it: a native key, a pointer, an emoji picker or
   * dictation that opens. The caret goes back before any later input can use it.
   */
  const onSelectionChange = (): void => {
    if (input.ownerDocument.activeElement === input) collapse();
  };
  const width = measurer(input);
  // The input grows with the text being composed, so the text shows at the caret while the
  // IME composes it (the slide shows the text only once the composition ends). The context
  // before it scrolls out of view. A keyboard that composes the word before the caret again
  // (`compositionEdit`) moves the input left over that word, so the composed text covers it.
  const fit = (): void => {
    input.style.width = '';
    input.style.marginLeft = '';
    input.style.direction = '';
    if (!composing || input.value === composeFrom.value) return;
    const from = composeFrom.value;
    const value = input.value;
    let shared = 0;
    while (shared < from.length && shared < value.length && from[shared] === value[shared])
      shared += 1;
    const back = shared < from.length ? width(from.slice(shared)) : 0;
    const shown = back > 0 ? width(value.slice(shared)) : 0;
    if (back > 0 && shown > 0) {
      input.style.width = `${Math.ceil(shown) + 1}px`;
      // In right-to-left text the word before the caret is on its right: the input stays at
      // the caret and grows right, and shows the end of its text, which is on its left.
      if (RTL.test(from.slice(shared))) {
        input.style.direction = 'rtl';
        input.scrollLeft = -input.scrollWidth;
        return;
      }
      input.style.marginLeft = `${-back}px`;
      input.scrollLeft = input.scrollWidth;
      return;
    }
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
    // An edit the browser does not let the editor cancel changes the field itself: the
    // `input` event after it applies that change once (`fieldEdit`), so nothing goes in twice.
    if (!e.cancelable) return;
    e.preventDefault();
    const target =
      e.inputType === 'insertReplacementText'
        ? replacementTarget(e, input, context)
        : e.inputType === 'insertText'
          ? inputTarget(input, context)
          : null;
    dispatchBeforeInput(host, e, target);
    sync();
  };
  const onInput = (): void => {
    if (composing) return fit();
    // An input that no `beforeinput` announced, or one that could not be prevented, changed
    // the field itself (an extension that writes the field, for example): its change goes
    // to the model, so the field and the model agree again.
    const before = context?.text ?? '';
    const edit =
      context && input.value !== before ? fieldEdit(before, input.value, context.key) : null;
    if (edit && host.replaceText) host.replaceText(edit.target, edit.text);
    collapse();
    sync();
  };
  const onCompositionStart = (): void => {
    // A composition that no key started (a native move came last) starts at the synced caret.
    collapse();
    composing = true;
    input.classList.add(COMPOSING_CLASS);
    composeFrom = { value: input.value, width: input.value ? input.scrollWidth : 0 };
    // The candidate window opens at the caret: the end of a long context.
    input.scrollLeft = input.scrollWidth;
  };
  const onCompositionEnd = (e: CompositionEvent): void => {
    const before = composeFrom.value;
    const { value, selectionEnd } = input;
    const key = context?.key ?? null;
    endComposition();
    const data = e.data ?? '';
    // A keyboard can recompose a word the input already holds: only the change goes in. A
    // composition that ends empty deletes what it covered, when the field lost it.
    const edit =
      before && (data || value !== before)
        ? compositionEdit(before, value, selectionEnd, data, key)
        : null;
    if (!edit || !host.replaceText) return void (data && host.textInput('text', data));
    if (edit.target.text === '' && edit.text === '') return; // Nothing changed.
    if (edit.target.text === '' && edit.target.after === 0)
      return host.textInput('text', edit.text);
    host.replaceText(edit.target, edit.text, true);
  };
  const onKeyDown = (e: KeyboardEvent): void => {
    // Before the key acts, also the first key of a composition (key code 229): a caret or a
    // selection that a native key left in the input is not the system's. The system makes
    // its own selection after this, for its own input.
    if (e.target === input) collapse();
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
  // A browser fires `selectionchange` at the document, at the input, or both; `select` covers
  // a selection with no `selectionchange`. The reset is the same for each.
  const doc = input.ownerDocument;
  doc.addEventListener('selectionchange', onSelectionChange);
  input.addEventListener('selectionchange', onSelectionChange);
  input.addEventListener('select', onSelectionChange);
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
    doc.removeEventListener('selectionchange', onSelectionChange);
    input.removeEventListener('selectionchange', onSelectionChange);
    input.removeEventListener('select', onSelectionChange);
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
