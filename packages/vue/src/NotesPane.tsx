/** @jsxImportSource vue */
// The speaker notes pane: the Vue twin of the React adapter's `NotesPane.tsx`, with the same
// markup, classes, ARIA, and keys. A collapsible field for the current slide's notes, with
// bold, italic, and underline. The field shows the notes formatting while you read and while
// you type (`NotesField`). The open state is shared with `useNotesPane` through the root.

import {
  computed,
  defineComponent,
  nextTick,
  onBeforeUnmount,
  ref,
  shallowRef,
  useId,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  NOTES_FORMAT_BUTTONS,
  NO_NOTES_FORMATS,
  armNotesFormat,
  createNotesFindPainter,
  followNotesArm,
  notesCanFormat,
  notesCommittedRange,
  notesFieldFocused,
  notesFieldParagraphs,
  notesFieldText,
  notesFormatRange,
  notesFormatsAt,
  notesSelection,
  onFindShown,
  sameNotesFormats,
  warnDuplicatePane,
  type NotesFormat,
  type NotesParagraph,
} from '@pptx-editor.dev/core/editor';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { rejectionMessage } from './EditNotice.tsx';
import { setNotesTarget } from './notes-target.ts';
import { usePaneState } from './pane-state.ts';
import { NotesField, type NotesRestore } from './NotesView.tsx';

/**
 * Uncommitted notes text and the slide it belongs to, with the formats the field shows for it,
 * so the commit stores what the user saw while typing.
 */
interface NotesDraft {
  readonly slideId: string;
  readonly text: string;
  readonly paragraphs?: readonly NotesParagraph[];
}

const notesPaneProps = {
  /** Whether the pane is open, controlled. Use with the `update:open` event (`v-model:open`). */
  open: { type: Boolean, default: undefined },
  /** Whether the pane is open first, uncontrolled. Default `true`. */
  defaultOpen: { type: Boolean, default: undefined },
  /**
   * Show the pane's bold, italic, and underline buttons. Default true. Ctrl/Cmd+B, I, and U
   * work in the notes field either way.
   */
  showFormatButtons: { type: Boolean, default: true },
  /** A class name added to the pane, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Inline styles for the pane, as React's `style`. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/**
 * Props of `PptxEditorNotesPane`, the twin of React's `NotesPaneProps`: the pane props, without
 * a width (the notes take the slide's width), and the format buttons. `className` (or a host
 * `class`) and `style` go on the pane. @public
 */
export type NotesPaneProps = ExtractPublicPropTypes<typeof notesPaneProps>;

type OpenProps = { readonly open: boolean | undefined; readonly defaultOpen: boolean | undefined };

/**
 * The open state of a notes pane: controlled by `open`, else the root's (so host chrome built
 * on `useNotesPane` agrees), else its own. One notes pane owns the root's state; a second one
 * in the same root warns once and keeps its own value.
 */
function useNotesOpen(props: OpenProps, emit: (event: 'update:open', open: boolean) => void) {
  const shared = usePaneState();
  const own = ref(props.defaultOpen ?? true);
  const owner = ref(false);
  const open = computed(() =>
    props.open !== undefined
      ? props.open
      : shared && owner.value
        ? (shared.notesOpen.value ?? props.defaultOpen ?? true)
        : own.value
  );
  const setOpen = (next: boolean): void => {
    if (props.open === undefined) {
      if (shared && owner.value) shared.setNotesOpen(next);
      else own.value = next;
    }
    emit('update:open', next);
  };
  // A toggle from `useNotesPane` goes through this pane, so a controlled pane tells its host.
  const handler = (next: boolean): void => setOpen(next);
  // The request slot decides ownership: the first pane takes it, and a later one warns once and
  // keeps its own state until the owner unmounts.
  const slot = shared?.notesRequest ?? null;
  if (slot?.current && slot.current !== handler) warnDuplicatePane('PptxEditorNotesPane');
  const claim = (): void => {
    if (slot && !slot.current) {
      slot.current = handler;
      owner.value = true;
    }
  };
  claim();
  // Also on each change of the root value: when the owner unmounts, the next change hands its
  // slot to a remaining pane.
  watch(() => shared?.notesOpen.value, claim);
  onBeforeUnmount(() => {
    if (slot?.current === handler) slot.current = null;
  });
  // The root follows the owning pane: its first value, and a controlled value.
  watch(
    [open, owner],
    () => {
      if (!shared || !owner.value) return;
      if (shared.notesOpen.value !== open.value) shared.setNotesOpen(open.value);
    },
    { immediate: true }
  );
  return { open, setOpen };
}

/**
 * Collapsible speaker notes for the current slide, the twin of React's `PptxEditor.NotesPane`.
 * The text shown is the document's; only the text typed since the field took focus is held
 * here, bound to the slide it was typed on. It is written to that slide (wherever the slide is
 * now) when the field loses focus, when the pointer goes down anywhere else, when the current
 * slide changes, and before a format or an undo. Emits `update:open`. @public
 */
export const PptxEditorNotesPane = defineComponent({
  name: 'PptxEditorNotesPane',
  props: notesPaneProps,
  emits: { 'update:open': (_open: boolean) => true },
  setup(props, { emit }) {
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const notes = useEditorState((s) => s.notes);
    const formatted = useEditorState((s) => s.notesFormatted);
    const slideId = useEditorState((s) => s.slides[s.currentSlide]?.id ?? null);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const pane = useNotesOpen(props, emit);
    const notesId = useId();
    const draft = shallowRef<NotesDraft | null>(null);
    const refusal = ref<string | null>(null);
    const formats = shallowRef(NO_NOTES_FORMATS);
    const canFormat = ref(false);
    const field = ref<HTMLDivElement | null>(null);
    let pending: NotesDraft | null = null;
    let bound: string | null = null;
    const restore: NotesRestore = { value: null };
    // An engine without formatted notes still shows the plain text.
    const paragraphs = computed<readonly NotesParagraph[]>(
      () =>
        formatted.value ??
        notes.value.split('\n').map((text) => ({
          runs: [{ text, bold: false, italic: false, underline: false, strike: false }],
        }))
    );

    /**
     * Write the draft to its own slide. A deleted slide takes its notes with it. `leaving` is
     * true when the field loses the focus: its selection is not put back.
     */
    const commit = (leaving = false): void => {
      const d = pending;
      pending = null;
      draft.value = null;
      const ed = editor.value;
      if (!ed || !d) return;
      // The field redraws from the written text, which `setNotes` cleans: while the field still
      // has the focus and holds the draft, its selection goes back at the offsets of the
      // written text. A field without the focus takes no selection back (see `NotesField`).
      const el = field.value;
      const selected =
        el && !leaving && !restore.value && notesFieldFocused(el) ? notesSelection(el) : null;
      if (el && selected) {
        const text = notesFieldText(el);
        if (text === d.text) restore.value = notesCommittedRange(text, selected);
      }
      const index = ed.snapshot().slides.findIndex((s) => s.id === d.slideId);
      if (index < 0) return;
      const r = ed.exec({
        type: 'setNotes',
        index,
        text: d.text,
        ...(d.paragraphs ? { paragraphs: d.paragraphs } : {}),
      });
      refusal.value = r.ok ? null : rejectionMessage(t, r);
    };

    // The current slide changed under a draft (New slide, Delete slide, the rail): commit it.
    watch(
      slideId,
      (id) => {
        if (pending && pending.slideId !== id) commit();
        bound = id;
      },
      { immediate: true }
    );

    // Commit before a click anywhere else acts, so a rail or toolbar command sees the notes.
    watch(draft, (d, _old, onCleanup) => {
      if (!d) return;
      const doc = field.value?.ownerDocument;
      const onPointerDown = (e: Event): void => {
        if (!field.value?.contains(e.target as Node)) commit();
      };
      doc?.addEventListener('pointerdown', onPointerDown, true);
      onCleanup(() => doc?.removeEventListener('pointerdown', onPointerDown, true));
    });

    // Closing, from this pane, the host, or `useNotesPane`, commits the draft first: the field
    // unmounts with no blur event. Unmount commits too.
    watch(
      () => pane.open.value,
      (open) => {
        if (!open) commit();
      },
      { flush: 'post' }
    );
    onBeforeUnmount(() => commit());

    /**
     * The B, I, and U buttons show the formats of the selected text. They take a pick only with
     * selected text or a caret in a word, as `applyFormat` needs.
     */
    const readFormats = (): void => {
      const el = field.value;
      // A caret that moved drops the format armed at it.
      if (el) followNotesArm(el);
      // A field without the focus formats nothing: its buttons and the toolbar show no format.
      const range = el && notesFieldFocused(el) ? notesSelection(el) : null;
      const next = el && range ? notesFormatsAt(el, range) : NO_NOTES_FORMATS;
      canFormat.value = !!el && notesCanFormat(notesFieldText(el), range);
      // Every key and pointer move reads the formats: render only when one of them changed.
      if (!sameNotesFormats(formats.value, next)) formats.value = next;
    };

    // The notes take the editing focus: no shape stays selected on the slide, and the
    // toolbar's bold, italic, underline, and strikethrough format the notes text.
    let focused = false;
    const leaveSlide = (): void => {
      const ed = editor.value;
      if (ed && ed.snapshot().selection.kind !== 'none') ed.setSelection({ kind: 'none' });
    };
    const active = ref(false);
    const target = () => ({
      formats: formats.value,
      readOnly: readOnly.value,
      canFormat: canFormat.value,
      apply: (f: NotesFormat) => applyFormat(f),
    });
    const onFocus = (): void => {
      focused = true;
      active.value = true;
      leaveSlide();
      if (editor.value) setNotesTarget(editor.value, target());
    };
    const onBlur = (): void => {
      focused = false;
      active.value = false;
      if (editor.value) setNotesTarget(editor.value, null);
      commit(true);
      // Without the focus, B, I, and U have nothing to format.
      canFormat.value = false;
      formats.value = NO_NOTES_FORMATS;
    };
    // While the field has the focus, every selection change (the keyboard, the pointer, a
    // script, the accessibility tree) updates the B, I, and U state. The focus can come back
    // with the same selection and fire no selectionchange, so the formats are read when the
    // field becomes active too.
    watch(
      active,
      (on, _old, onCleanup) => {
        const doc = field.value?.ownerDocument;
        if (!on || !doc) return;
        readFormats();
        doc.addEventListener('selectionchange', readFormats);
        onCleanup(() => doc.removeEventListener('selectionchange', readFormats));
      },
      { flush: 'post' }
    );
    // The find bar's match in the notes: painted while find shows it, and not while the field
    // has the focus or holds a draft. A step inside the notes moves no slide or selection, so
    // find tells the pane when it shows another match. The paint waits for the field to draw.
    const painter = createNotesFindPainter();
    let mounted = true;
    const findCount = ref(0);
    const selectionKind = useEditorState((s) => s.selection.kind);
    const revision = useEditorState((s) => s.revision);
    watch(
      editor,
      (ed, _old, onCleanup) => {
        if (ed)
          onCleanup(
            onFindShown(ed.find, () => {
              findCount.value += 1;
            })
          );
      },
      { immediate: true }
    );
    watch(
      [
        editor,
        () => pane.open.value,
        paragraphs,
        draft,
        active,
        findCount,
        selectionKind,
        revision,
        slideId,
      ],
      () =>
        void nextTick(() => {
          if (!mounted) return;
          painter.paint(
            editor.value,
            pane.open.value ? field.value : null,
            paragraphs.value,
            !!draft.value
          );
        }),
      { flush: 'post', immediate: true }
    );
    onBeforeUnmount(() => {
      mounted = false;
      painter.stop();
    });
    watch([editor, formats, readOnly, canFormat], () => {
      if (editor.value && focused) setNotesTarget(editor.value, target());
    });
    watch(editor, (_next, was) => {
      if (was) setNotesTarget(was, null);
    });
    onBeforeUnmount(() => {
      if (editor.value) setNotesTarget(editor.value, null);
    });

    const onDraft = (text: string): void => {
      const id = pending?.slideId ?? bound ?? slideId.value;
      if (!id) return;
      const el = field.value;
      pending = { slideId: id, text, ...(el ? { paragraphs: notesFieldParagraphs(el) } : {}) };
      draft.value = pending;
    };

    /**
     * Bold, italic, or underline on the selected notes text, as one undo step. With a caret in
     * a word, the word takes the format; the caret stays where it was. With any other caret,
     * the format is armed: the next typed text takes it, as on the canvas.
     */
    const applyFormat = (format: NotesFormat): boolean => {
      const el = field.value;
      const ed = editor.value;
      if (!ed || readOnly.value || !el) return false;
      const text = notesFieldText(el);
      const selected = notesSelection(el);
      const word = notesFormatRange(text, selected);
      if (!word && armNotesFormat(el, format)) {
        refusal.value = null;
        readFormats();
        return true;
      }
      if (!word) {
        refusal.value = t('disabled.selectNotesText');
        return false;
      }
      // Typed text goes in first (its own undo step). `setNotes` drops characters a slide
      // cannot hold, so the offsets are counted in the text it writes.
      const range = notesCommittedRange(text, word)!;
      const caret = notesCommittedRange(text, selected);
      commit();
      const index = ed.snapshot().slides.findIndex((s) => s.id === slideId.value);
      if (index < 0) return false;
      restore.value = caret;
      const r = ed.exec({ type: 'formatNotes', index, ...range, format });
      if (!r.ok) restore.value = null;
      refusal.value = r.ok ? null : rejectionMessage(t, r);
      return r.ok;
    };

    /** Undo and redo in the field act on the deck, after the draft is written. */
    const applyHistory = (kind: 'undo' | 'redo'): void => {
      const el = field.value;
      const ed = editor.value;
      if (!ed || !el) return;
      const range = notesCommittedRange(notesFieldText(el), notesSelection(el));
      commit();
      const before = ed.snapshot().notesFormatted;
      restore.value = range;
      ed.exec({ type: kind });
      // The notes did not change: no redraw takes the selection, so drop it.
      if (ed.snapshot().notesFormatted === before) restore.value = null;
      // An undo can bring back a shape selection; the cursor stays in the notes, so it goes.
      leaveSlide();
    };

    // Layouts and masters have no speaker notes: the slide master view hides the pane.
    const masterView = useEditorState((s) => s.presentationView === 'slideMaster');
    return () => {
      if (masterView.value) return null;
      const open = pane.open.value;
      const d = draft.value;
      return (
        <section
          class={['pptx-notes', props.className]}
          style={props.style}
          aria-label={t('notes.title')}
          data-open={open ? 'true' : 'false'}
        >
          <div class="pptx-notes__head" onMousedown={keepFocus}>
            <span>{t('notes.title')}</span>
            {open && !readOnly.value && props.showFormatButtons ? (
              <div class="pptx-notes__format" role="group" aria-label={t('notes.format')}>
                {NOTES_FORMAT_BUTTONS.map(([format, icon, key]) => (
                  <ToolButton
                    key={format}
                    label={t(key)}
                    icon={icon}
                    pressed={formats.value[format]}
                    disabled={!canFormat.value}
                    disabledReason={t('disabled.selectNotesText')}
                    onClick={() => applyFormat(format)}
                  />
                ))}
              </div>
            ) : null}
            <ToolButton
              label={open ? t('notes.collapse') : t('notes.expand')}
              icon={open ? 'keyboard_arrow_down' : 'keyboard_arrow_up'}
              expanded={open}
              controls={open ? notesId : undefined}
              onClick={() => {
                commit();
                pane.setOpen(!open);
              }}
            />
          </div>
          {open ? (
            <NotesField
              id={notesId}
              paragraphs={paragraphs.value}
              drafting={!!d && d.slideId === slideId.value}
              label={t('notes.title')}
              placeholder={t('notes.placeholder')}
              readOnly={readOnly.value}
              field={field}
              restore={restore}
              onDraft={onDraft}
              onCommit={onBlur}
              onFormat={applyFormat}
              onHistory={applyHistory}
              onNewSlide={() => {
                // The field moves to the new slide: no caret comes back.
                commit(true);
                editor.value?.exec({ type: 'addSlide' });
              }}
              onSelection={readFormats}
              onFocus={onFocus}
            />
          ) : null}
          {refusal.value ? (
            <p class="pptx-notes__refusal" role="alert">
              {refusal.value}
            </p>
          ) : null}
        </section>
      );
    };
  },
});
