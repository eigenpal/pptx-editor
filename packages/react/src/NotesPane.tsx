/** @jsxImportSource react */
// The speaker notes pane: a collapsible field for the current slide's notes, with bold, italic,
// and underline. The field shows the notes formatting while you read and while you type
// (`NotesField`).

import {
  useCallback,
  useContext,
  useEffect,
  useId,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  NOTES_FORMAT_BUTTONS,
  NO_NOTES_FORMATS,
  armNotesFormat,
  createNotesFindPainter,
  followNotesArm,
  notesCanFormat,
  notesCommittedRange,
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
import { rejectionMessage } from './EditNotice.tsx';
import { NotesField, notesFieldFocused, type NotesCaret } from './NotesView.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { setNotesTarget } from './notes-target.ts';
import { PaneStateContext } from './pane-state.ts';
import type { PaneProps } from './panes.tsx';

const NO_FORMATS = NO_NOTES_FORMATS;

/**
 * Uncommitted notes text and the slide it belongs to, with the formats the field shows for it,
 * so the commit stores what the user saw while typing.
 */
interface NotesDraft {
  readonly slideId: string;
  readonly text: string;
  readonly paragraphs?: readonly NotesParagraph[];
}

/**
 * Props of `NotesPane`: the pane props, without a width (the notes take the slide's width),
 * and the format buttons. @public
 */
export interface NotesPaneProps extends Omit<PaneProps, 'paneWidth'> {
  /**
   * Show the pane's bold, italic, and underline buttons. Default true. Ctrl/Cmd+B, I, and U
   * work in the notes field either way.
   */
  readonly showFormatButtons?: boolean;
}

/**
 * The open state of a notes pane: controlled by `open`, else the root's (so host chrome built
 * on `useNotesPane` agrees), else its own. One notes pane owns the root's state; a second one
 * in the same root warns once and keeps its own value.
 */
function useNotesOpen(props: NotesPaneProps) {
  const shared = useContext(PaneStateContext);
  const [own, setOwn] = useState(props.defaultOpen ?? true);
  const [owner, setOwner] = useState(false);
  const controlled = props.open !== undefined;
  const useShared = shared !== null && owner;
  const open = controlled
    ? props.open!
    : useShared
      ? (shared.notesOpen ?? props.defaultOpen ?? true)
      : own;
  const setOpenValue = (next: boolean): void => {
    if (!controlled) {
      if (useShared) shared.setNotesOpen(next);
      else setOwn(next);
    }
    props.onOpenChange?.(next);
  };
  // A toggle from `useNotesPane` goes through this pane, so a controlled pane tells its host.
  const request = useRef(setOpenValue);
  request.current = setOpenValue;
  const handler = useRef((next: boolean): void => request.current(next)).current;
  // The stable request slot decides ownership: the first pane takes it, and a later one warns
  // once and keeps its own state until the owner unmounts.
  const slot = shared?.notesRequest ?? null;
  useLayoutEffect(() => {
    if (slot?.current && slot.current !== handler) warnDuplicatePane('PptxEditor.NotesPane');
  }, [slot, handler]);
  // Also on each change of the root value: when the owner unmounts, the next change hands its
  // slot to a remaining pane.
  const rootValue = shared?.notesOpen;
  useLayoutEffect(() => {
    if (slot && !slot.current) {
      slot.current = handler;
      setOwner(true);
    }
  }, [slot, handler, rootValue]);
  useLayoutEffect(
    () => () => {
      if (slot?.current === handler) slot.current = null;
    },
    [slot, handler]
  );
  // The root follows the owning pane: its first value, and a controlled value.
  useLayoutEffect(() => {
    if (!shared || !owner) return;
    if (shared.notesOpen !== open) shared.setNotesOpen(open);
  });
  return [open, setOpenValue] as const;
}

/**
 * Collapsible speaker notes for the current slide. The text shown is the document's; only the
 * text typed since the field took focus is held here, bound to the slide it was typed on. It
 * is written to that slide (wherever the slide is now) when the field loses focus, when the
 * pointer goes down anywhere else, when the current slide changes, and before a format or an
 * undo. @public
 */
export function NotesPane(props: NotesPaneProps = {}) {
  const { className, showFormatButtons = true } = props;
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const notes = useEditorState((s) => s.notes);
  const formatted = useEditorState((s) => s.notesFormatted);
  const slideId = useEditorState((s) => s.slides[s.currentSlide]?.id ?? null);
  const readOnly = useEditorState((s) => s.mode === 'view');
  const [open, setOpen] = useNotesOpen(props);
  const notesId = useId();
  const [draft, setDraft] = useState<NotesDraft | null>(null);
  const [refusal, setRefusal] = useState<string | null>(null);
  const [formats, setFormats] = useState(NO_FORMATS);
  const [canFormat, setCanFormat] = useState(false);
  const field = useRef<HTMLDivElement>(null);
  const pending = useRef<NotesDraft | null>(null);
  const bound = useRef<string | null>(null);
  const restore = useRef<NotesCaret | null>(null);
  // An engine without formatted notes still shows the plain text.
  const paragraphs = useMemo<readonly NotesParagraph[]>(
    () =>
      formatted ??
      notes.split('\n').map((text) => ({
        runs: [{ text, bold: false, italic: false, underline: false, strike: false }],
      })),
    [formatted, notes]
  );

  /**
   * Write the draft to its own slide. A deleted slide takes its notes with it. `leaving` is
   * true when the field loses the focus: its selection is not put back.
   */
  const commit = useCallback(
    (leaving = false): void => {
      const d = pending.current;
      pending.current = null;
      setDraft(null);
      if (!editor || !d) return;
      // The field redraws from the written text, which `setNotes` cleans: while the field still
      // has the focus and holds the draft, its selection goes back at the offsets of the written
      // text. A field without the focus takes no selection back (see `NotesField`).
      const el = field.current;
      const selected =
        el && !leaving && !restore.current && notesFieldFocused(el) ? notesSelection(el) : null;
      if (el && selected) {
        const text = notesFieldText(el);
        if (text === d.text) restore.current = notesCommittedRange(text, selected);
      }
      const index = editor.snapshot().slides.findIndex((s) => s.id === d.slideId);
      if (index < 0) return;
      const r = editor.exec({
        type: 'setNotes',
        index,
        text: d.text,
        ...(d.paragraphs ? { paragraphs: d.paragraphs } : {}),
      });
      setRefusal(r.ok ? null : rejectionMessage(t, r));
    },
    [editor, t]
  );

  // The current slide changed under a draft (New slide, Delete slide, the rail): commit it.
  useEffect(() => {
    if (pending.current && pending.current.slideId !== slideId) commit();
    bound.current = slideId;
  }, [slideId, commit]);

  // Commit before a click anywhere else acts, so a rail or toolbar command sees the notes.
  useEffect(() => {
    if (!draft) return;
    const doc = field.current?.ownerDocument;
    const onPointerDown = (e: Event): void => {
      if (!field.current?.contains(e.target as Node)) commit();
    };
    doc?.addEventListener('pointerdown', onPointerDown, true);
    return () => doc?.removeEventListener('pointerdown', onPointerDown, true);
  }, [draft, commit]);

  // Closing, from this pane, the host, or `useNotesPane`, commits the draft first: the field
  // unmounts with no blur event. Unmount commits too.
  useLayoutEffect(() => {
    if (!open) commit();
  }, [open, commit]);
  useEffect(() => () => commit(), [commit]);

  /**
   * The B, I, and U buttons show the formats of the selected text. They take a pick only with
   * selected text or a caret in a word, as `applyFormat` needs.
   */
  const readFormats = useCallback((): void => {
    const el = field.current;
    // A caret that moved drops the format armed at it.
    if (el) followNotesArm(el);
    // A field without the focus formats nothing: its buttons and the toolbar show no format.
    const range = el && notesFieldFocused(el) ? notesSelection(el) : null;
    const next = el && range ? notesFormatsAt(el, range) : NO_FORMATS;
    setCanFormat(!!el && notesCanFormat(notesFieldText(el), range));
    // Every key and pointer move reads the formats: render only when one of them changed.
    setFormats((was) => (sameNotesFormats(was, next) ? was : next));
  }, []);

  // The notes take the editing focus: no shape stays selected on the slide,
  // and the toolbar's bold, italic, underline, and strikethrough format the notes text.
  const focused = useRef(false);
  const applyRef = useRef<(format: NotesFormat) => boolean>(() => false);
  const leaveSlide = useCallback((): void => {
    if (editor && editor.snapshot().selection.kind !== 'none')
      editor.setSelection({ kind: 'none' });
  }, [editor]);
  const [active, setActive] = useState(false);
  const onFocus = (): void => {
    focused.current = true;
    setActive(true);
    leaveSlide();
    if (editor)
      setNotesTarget(editor, { formats, readOnly, canFormat, apply: (f) => applyRef.current(f) });
  };
  const onBlur = (): void => {
    focused.current = false;
    setActive(false);
    if (editor) setNotesTarget(editor, null);
    commit(true);
    // Without the focus, B, I, and U have nothing to format.
    setCanFormat(false);
    setFormats(NO_FORMATS);
  };
  // While the field has the focus, every selection change (the keyboard, the pointer, a
  // script, the accessibility tree) updates the B, I, and U state. The focus can come back with
  // the same selection (window refocus, a script's focus()) and fire no selectionchange, so the
  // formats are read when the field becomes active too.
  useEffect(() => {
    const doc = field.current?.ownerDocument;
    if (!active || !doc) return;
    readFormats();
    doc.addEventListener('selectionchange', readFormats);
    return () => doc.removeEventListener('selectionchange', readFormats);
  }, [active, readFormats]);
  useEffect(() => {
    if (editor && focused.current)
      setNotesTarget(editor, { formats, readOnly, canFormat, apply: (f) => applyRef.current(f) });
  }, [editor, formats, readOnly, canFormat]);
  useEffect(
    () => () => {
      if (editor) setNotesTarget(editor, null);
    },
    [editor]
  );

  // The find bar's match in the notes: painted while find shows it, and not while the field
  // has the focus or holds a draft. A step inside the notes moves no slide or selection, so
  // find tells the pane when it shows another match.
  const painter = useMemo(() => createNotesFindPainter(), []);
  const [findCount, setFindCount] = useState(0);
  const selectionKind = useEditorState((s) => s.selection.kind);
  const revision = useEditorState((s) => s.revision);
  useEffect(
    () => (editor ? onFindShown(editor.find, () => setFindCount((n) => n + 1)) : undefined),
    [editor]
  );
  useEffect(() => {
    painter.paint(editor, open ? field.current : null, paragraphs, !!draft);
  }, [
    painter,
    editor,
    open,
    paragraphs,
    draft,
    active,
    findCount,
    selectionKind,
    revision,
    slideId,
  ]);
  useEffect(() => () => painter.stop(), [painter]);

  const onDraft = (text: string): void => {
    const id = pending.current?.slideId ?? bound.current ?? slideId;
    if (!id) return;
    const el = field.current;
    pending.current = {
      slideId: id,
      text,
      ...(el ? { paragraphs: notesFieldParagraphs(el) } : {}),
    };
    setDraft(pending.current);
  };

  /**
   * Bold, italic, or underline on the selected notes text, as one undo step. With a caret in a
   * word, the word takes the format; the caret stays where it was. With any other caret, the
   * format is armed: the next typed text takes it, as on the canvas.
   */
  const applyFormat = (format: NotesFormat): boolean => {
    const el = field.current;
    if (!editor || readOnly || !el) return false;
    const text = notesFieldText(el);
    const selected = notesSelection(el);
    const word = notesFormatRange(text, selected);
    if (!word && armNotesFormat(el, format)) {
      setRefusal(null);
      readFormats();
      return true;
    }
    if (!word) {
      setRefusal(t('disabled.selectNotesText'));
      return false;
    }
    // Typed text goes in first (its own undo step). `setNotes` drops characters a slide cannot
    // hold, so the offsets are counted in the text it writes.
    const range = notesCommittedRange(text, word)!;
    const caret = notesCommittedRange(text, selected);
    commit();
    const index = editor.snapshot().slides.findIndex((s) => s.id === slideId);
    if (index < 0) return false;
    restore.current = caret;
    const r = editor.exec({ type: 'formatNotes', index, ...range, format });
    if (!r.ok) restore.current = null;
    setRefusal(r.ok ? null : rejectionMessage(t, r));
    return r.ok;
  };
  applyRef.current = applyFormat;

  /** Undo and redo in the field act on the deck, after the draft is written. */
  const applyHistory = (kind: 'undo' | 'redo'): void => {
    const el = field.current;
    if (!editor || !el) return;
    const text = notesFieldText(el);
    const selected = notesSelection(el);
    const range = notesCommittedRange(text, selected);
    commit();
    const before = editor.snapshot().notesFormatted;
    restore.current = range;
    editor.exec({ type: kind });
    // The notes did not change: no redraw takes the selection, so drop it.
    if (editor.snapshot().notesFormatted === before) restore.current = null;
    // An undo can bring back a shape selection; the cursor stays in the notes, so it goes.
    leaveSlide();
  };

  return (
    <section
      className={['pptx-notes', className].filter(Boolean).join(' ')}
      aria-label={t('notes.title')}
      data-open={open ? 'true' : 'false'}
      style={props.style}
    >
      <div className="pptx-notes__head" onMouseDown={keepFocus}>
        <span>{t('notes.title')}</span>
        {open && !readOnly && showFormatButtons ? (
          <div className="pptx-notes__format" role="group" aria-label={t('notes.format')}>
            {NOTES_FORMAT_BUTTONS.map(([format, icon, key]) => (
              <ToolButton
                key={format}
                label={t(key)}
                icon={icon}
                pressed={formats[format]}
                disabled={!canFormat}
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
          {...(open ? { controls: notesId } : {})}
          onClick={() => {
            commit();
            setOpen(!open);
          }}
        />
      </div>
      {open ? (
        <NotesField
          id={notesId}
          paragraphs={paragraphs}
          drafting={!!draft && draft.slideId === slideId}
          label={t('notes.title')}
          placeholder={t('notes.placeholder')}
          readOnly={readOnly}
          fieldRef={field}
          restore={restore}
          onDraft={onDraft}
          onCommit={onBlur}
          onFormat={applyFormat}
          onHistory={applyHistory}
          onNewSlide={() => {
            // The field moves to the new slide: no caret comes back.
            commit(true);
            editor?.exec({ type: 'addSlide' });
          }}
          onSelection={readFormats}
          onFocus={onFocus}
        />
      ) : null}
      {refusal ? (
        <p className="pptx-notes__refusal" role="alert">
          {refusal}
        </p>
      ) : null}
    </section>
  );
}
