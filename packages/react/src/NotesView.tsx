/** @jsxImportSource react */
// The speaker notes field: one editable element that shows each run's bold, italic, underline,
// and strikethrough while you read and while you type. Typed
// text is a draft that the pane commits as one `setNotes` (one undo step); formats go through
// `formatNotes`. The core helpers (`renderNotes`, `notesSelection`) own the DOM, so React never
// renders into the editable element and the caret stays where the user put it.

import { useEffect, useLayoutEffect, useRef, type KeyboardEvent } from 'react';
import {
  notesFieldBeforeInput,
  notesFieldFocused,
  notesFieldKeyDown,
  notesFieldText,
  notesFormatShortcut,
  notesSelection,
  renderNotes,
  selectNotes,
  type NotesFormat,
  type NotesParagraph,
} from '@pptx-editor.dev/core/editor';

/** The keyboard shortcut of a notes format (Ctrl/Cmd + B, I, U), or null. */
export const formatShortcut: (e: KeyboardEvent) => NotesFormat | null = notesFormatShortcut;

export { notesFieldFocused };

/** A selection to put back after the field draws the committed notes. */
export interface NotesCaret {
  readonly start: number;
  readonly end: number;
}

/**
 * The notes field. `paragraphs` draw it whenever no draft is pending (`drafting` false); while a
 * draft is pending, the DOM is the user's text and stays as it is.
 */
export function NotesField({
  id,
  paragraphs,
  drafting,
  label,
  placeholder,
  readOnly,
  fieldRef,
  restore,
  onDraft,
  onCommit,
  onFormat,
  onHistory,
  onNewSlide,
  onSelection,
  onFocus,
}: {
  id?: string;
  paragraphs: readonly NotesParagraph[];
  drafting: boolean;
  label: string;
  placeholder: string;
  readOnly: boolean;
  fieldRef: React.RefObject<HTMLDivElement | null>;
  /** Where the selection goes after the next draw (a format keeps its range). */
  restore: React.RefObject<NotesCaret | null>;
  onDraft: (text: string) => void;
  onCommit: () => void;
  onFormat: (format: NotesFormat) => void;
  onHistory: (kind: 'undo' | 'redo') => void;
  /** New Slide (Ctrl+M), as on the canvas. */
  onNewSlide?: () => void;
  onSelection: () => void;
  onFocus?: () => void;
}) {
  const composing = useRef(false);
  // Draw the notes when they change and no draft holds the DOM; keep the caret where it was.
  // Only a focused field takes its selection back: a selection put into a field that lost the
  // focus (Tab, Escape) takes the focus back in Chromium, and the key is lost.
  useLayoutEffect(() => {
    const el = fieldRef.current;
    if (!el || drafting) return;
    const focused = notesFieldFocused(el);
    const keep = focused ? (restore.current ?? notesSelection(el)) : null;
    restore.current = null;
    renderNotes(el, paragraphs);
    if (keep) selectNotes(el, keep.start, keep.end);
    onSelection();
  }, [paragraphs, drafting, fieldRef, restore, onSelection]);

  // Typed text at a caret with an armed format goes into a run with that format. React's
  // `onBeforeInput` carries no `inputType`, so the native event is used.
  const draftRef = useRef(onDraft);
  draftRef.current = onDraft;
  useEffect(() => {
    const el = fieldRef.current;
    if (!el) return;
    const onBeforeInput = (e: Event): void =>
      notesFieldBeforeInput(e as InputEvent, el, {
        readOnly,
        onDraft: (text) => draftRef.current(text),
      });
    el.addEventListener('beforeinput', onBeforeInput);
    return () => el.removeEventListener('beforeinput', onBeforeInput);
  }, [fieldRef, readOnly]);

  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>): void => {
    if (composing.current) return;
    notesFieldKeyDown(e, e.currentTarget, { readOnly, onDraft, onFormat, onHistory, onNewSlide });
  };
  return (
    <div
      ref={fieldRef}
      id={id}
      className="pptx-notes__text"
      role="textbox"
      aria-multiline="true"
      aria-label={label}
      aria-placeholder={placeholder}
      aria-readonly={readOnly || undefined}
      data-placeholder={placeholder}
      contentEditable={readOnly ? 'false' : 'plaintext-only'}
      suppressContentEditableWarning
      spellCheck
      tabIndex={0}
      onKeyDown={onKeyDown}
      onInput={(e) => {
        if (!readOnly) onDraft(notesFieldText(e.currentTarget));
      }}
      onCompositionStart={() => (composing.current = true)}
      onCompositionEnd={(e) => {
        composing.current = false;
        if (!readOnly) onDraft(notesFieldText(e.currentTarget));
      }}
      onBlur={onCommit}
      onFocus={onFocus}
      onKeyUp={onSelection}
      onMouseUp={onSelection}
    />
  );
}
