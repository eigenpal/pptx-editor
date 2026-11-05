/** @jsxImportSource vue */
// The speaker notes field: the Vue twin of the React adapter's `NotesView.tsx`. One editable
// element shows each run's bold, italic, underline, and strikethrough while you read and while
// you type. Typed text is a draft that the pane commits as one `setNotes` (one undo step);
// formats go through `formatNotes`. The core helpers (`renderNotes`, `notesSelection`) own the
// DOM, so Vue never renders into the editable element and the caret stays where the user put
// it. The keys come from core (`notesFieldKeyDown`).

import { defineComponent, onMounted, watch, type PropType, type Ref } from 'vue';
import {
  notesFieldBeforeInput,
  notesFieldFocused,
  notesFieldKeyDown,
  notesFieldText,
  notesSelection,
  renderNotes,
  selectNotes,
  type NotesFormat,
  type NotesParagraph,
  type NotesRange,
} from '@pptx-editor.dev/core/editor';

/** A selection to put back after the field draws the committed notes. @internal */
export type NotesCaret = NotesRange;

/** A holder for the selection to put back after the next draw. @internal */
export interface NotesRestore {
  value: NotesCaret | null;
}

/**
 * The notes field. `paragraphs` draw it whenever no draft is pending (`drafting` false); while a
 * draft is pending, the DOM is the user's text and stays as it is. @internal
 */
export const NotesField = defineComponent({
  name: 'PptxNotesField',
  props: {
    id: { type: String, default: undefined },
    paragraphs: { type: Array as PropType<readonly NotesParagraph[]>, required: true },
    drafting: { type: Boolean, required: true },
    label: { type: String, required: true },
    placeholder: { type: String, required: true },
    readOnly: { type: Boolean, required: true },
    /** The pane's ref to the field element. */
    field: { type: Object as PropType<Ref<HTMLDivElement | null>>, required: true },
    /** Where the selection goes after the next draw (a format keeps its range). */
    restore: { type: Object as PropType<NotesRestore>, required: true },
    onDraft: { type: Function as PropType<(text: string) => void>, required: true },
    onCommit: { type: Function as PropType<() => void>, required: true },
    onFormat: { type: Function as PropType<(format: NotesFormat) => void>, required: true },
    onHistory: { type: Function as PropType<(kind: 'undo' | 'redo') => void>, required: true },
    /** New Slide (Ctrl+M), as on the canvas. */
    onNewSlide: { type: Function as PropType<() => void>, default: undefined },
    onSelection: { type: Function as PropType<() => void>, required: true },
    onFocus: { type: Function as PropType<() => void>, default: undefined },
  },
  setup(props) {
    let composing = false;
    // Draw the notes when they change and no draft holds the DOM; keep the caret where it was.
    // Only a focused field takes its selection back: a selection put into a field that lost the
    // focus (Tab, Escape) takes the focus back in Chromium, and the key is lost.
    const draw = (): void => {
      const el = props.field.value;
      if (!el || props.drafting) return;
      const focused = notesFieldFocused(el);
      const keep = focused ? (props.restore.value ?? notesSelection(el)) : null;
      props.restore.value = null;
      renderNotes(el, props.paragraphs);
      if (keep) selectNotes(el, keep.start, keep.end);
      props.onSelection();
    };
    onMounted(draw);
    watch(() => [props.paragraphs, props.drafting] as const, draw, { flush: 'post' });
    const onKeydown = (e: KeyboardEvent): void => {
      if (composing) return;
      notesFieldKeyDown(e, e.currentTarget as HTMLElement, {
        readOnly: props.readOnly,
        onDraft: props.onDraft,
        onFormat: props.onFormat,
        onHistory: props.onHistory,
        onNewSlide: props.onNewSlide,
      });
    };
    const draft = (e: Event): void => {
      if (!props.readOnly) props.onDraft(notesFieldText(e.currentTarget as HTMLElement));
    };
    return () => (
      <div
        ref={props.field}
        id={props.id}
        class="pptx-notes__text"
        role="textbox"
        aria-multiline="true"
        aria-label={props.label}
        aria-placeholder={props.placeholder}
        aria-readonly={props.readOnly || undefined}
        data-placeholder={props.placeholder}
        contenteditable={props.readOnly ? 'false' : 'plaintext-only'}
        spellcheck="true"
        tabindex={0}
        onKeydown={onKeydown}
        // Typed text at a caret with an armed format goes into a run with that format.
        onBeforeinput={(e: InputEvent) =>
          notesFieldBeforeInput(e, e.currentTarget as HTMLElement, {
            readOnly: props.readOnly,
            onDraft: props.onDraft,
          })
        }
        onInput={draft}
        onCompositionstart={() => {
          composing = true;
        }}
        onCompositionend={(e: CompositionEvent) => {
          composing = false;
          draft(e);
        }}
        onBlur={() => props.onCommit()}
        onFocus={() => props.onFocus?.()}
        onKeyup={() => props.onSelection()}
        onMouseup={() => props.onSelection()}
      />
    );
  },
});
