// The editor's side of the clipboard: copy the selected shapes, fill a copy or cut event, and
// paste from a paste event. The formats and their checks are in `clipboard.ts`, and a table's
// cell ranges in `table-clipboard.ts`; this module keeps the editor state they need (the last
// copy, the last parsed payload) out of the facade.

import type { ClipboardPayload, EditorSelection } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import {
  CLIPBOARD_MIME,
  clipboardText,
  copySelection,
  parseClipboardText,
  readClipboard,
} from './clipboard.ts';
import type { DocumentSession } from './document-session.ts';
import { selectedText } from './selection-state.ts';
import { copyCellsTo, tablePasteFrom } from './table-clipboard.ts';

/** What the clipboard reads from and asks of the editor. */
export interface ClipboardEditor {
  session(): DocumentSession | null;
  slide(): number;
  selection(): EditorSelection;
  layout(): SlideLayout | undefined;
  readOnly(): boolean;
  /** Publish a state change (Paste's enabled state reads the last copy). */
  emit(): void;
  /** Delete the selected text (a cut of text). */
  deleteText(): void;
  /** Delete the selected shapes as a cut (`RemovalNotice.cut`). */
  cutShapes(): void;
  /** Clear the selected cells (a cut of a cell range). */
  clearCells(): void;
  /** Fill table cells from the selection's top-left cell (`pasteCells`). */
  pasteCells(texts: string[][]): void;
  /** Insert text at the caret, as typing does. */
  insertText(text: string): void;
  refuse(reason: string): void;
  paste(payload: ClipboardPayload): void;
}

/** The clipboard handlers of one editor. */
export function createClipboardHost(editor: ClipboardEditor) {
  let lastCopied: ClipboardPayload | null = null;
  let lastParsed: { raw: string; payload: unknown } = { raw: '', payload: null };

  /** The selected shapes as a clipboard payload. */
  const copy = (): ClipboardPayload | null => {
    const session = editor.session();
    const selection = editor.selection();
    if (!session || selection.kind !== 'shapes') return null;
    const ref = session.info.slides[editor.slide()];
    if (!ref) return null;
    const payload = copySelection(
      session.store.package,
      ref.partName,
      editor.layout()?.items ?? [],
      selection,
      session.id
    );
    if (payload) {
      lastCopied = payload;
      // Paste's enabled state reads `lastCopied`: tell the chrome it changed.
      editor.emit();
    }
    return payload;
  };

  return {
    get lastCopied(): ClipboardPayload | null {
      return lastCopied;
    },
    copy,
    /** Fill a clipboard event: shapes as the editor's format plus plain text. */
    copyTo(data: DataTransfer, cut: boolean): boolean {
      const selection = editor.selection();
      // A cell range: its text as tab-separated text and as this editor's cells format.
      if (copyCellsTo(data, editor.layout(), selection)) {
        if (cut && !editor.readOnly()) editor.clearCells();
        return true;
      }
      if (selection.kind === 'text') {
        const text = selectedText(editor.layout(), selection);
        if (!text) return false;
        data.setData('text/plain', text);
        if (cut) editor.deleteText();
        return true;
      }
      const payload = copy();
      if (!payload) return false;
      data.setData(CLIPBOARD_MIME, JSON.stringify(payload));
      data.setData('text/plain', clipboardText(payload));
      if (cut) editor.cutShapes();
      return true;
    },
    /**
     * Paste shapes from a clipboard event. Text pastes into text editing are left alone, and so
     * is a clipboard that holds neither this editor's format nor the text of the last copy.
     */
    pasteFrom(data: DataTransfer): boolean {
      if (editor.readOnly()) return false;
      // Copied cells or tab-separated text fill cells; one copied cell goes in at the caret.
      const table = tablePasteFrom(data, editor.selection());
      if (table && 'cells' in table) return (editor.pasteCells(table.cells), true);
      if (table) return (editor.insertText(table.text), true);
      if (editor.selection().kind === 'text') return false;
      // A payload parsed once is cached by its text, so a paste series keeps one identity.
      const read = readClipboard(data, lastCopied, (raw) => {
        if (raw === lastParsed.raw) return lastParsed.payload;
        const payload = parseClipboardText(raw);
        if (payload) lastParsed = { raw, payload };
        return payload;
      });
      if (!read) return false;
      if ('refused' in read) return (editor.refuse('errors.clipboard'), true);
      editor.paste(read.payload as ClipboardPayload);
      return true;
    },
  };
}
