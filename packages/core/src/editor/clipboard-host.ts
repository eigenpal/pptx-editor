// The editor's side of the clipboard: copy the selected shapes or text, fill a copy or cut event,
// and paste from a paste event. The formats and their checks are in `clipboard.ts` (shapes),
// `clipboard-text.ts` and `clipboard-rich.ts` (text), `clipboard-picture.ts` (the picture other
// applications paste shapes as), and `table-clipboard.ts` (cell ranges); this module keeps the
// editor state they need (the last copy, the last parsed payload) out of the facade.
//
// What a copy writes, by selection:
// - Text: plain text, HTML with the formatting the slide shows, and the runs with their own
//   formatting in this editor's text format.
// - Shapes: this editor's shape format, their text as plain text, and HTML for other
//   applications: the text of text boxes, or a picture of the shapes.
//
// What a paste reads, in order: cells, this editor's text or shapes, HTML text, a picture file,
// and plain text. Text pasted with no text selected goes in a new text box.

import type { ClipboardPayload, EditorCommand, EditorSelection } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import {
  CLIPBOARD_MIME,
  clipboardText,
  copySelection,
  parseClipboardText,
  readClipboard,
} from './clipboard.ts';
import {
  paragraphSpecs,
  richFromHtml,
  richFromText,
  richHtml,
  type RichParagraph,
} from './clipboard-rich.ts';
import {
  cleanParagraphs,
  copyTextParagraphs,
  richFromSelection,
  TEXT_MIME,
} from './clipboard-text.ts';
import {
  createPictureCache,
  drawShapesPicture,
  pictureHtml,
  pictureMark,
  pngResolution,
  type ShapesPicture,
} from './clipboard-picture.ts';
import type { DocumentSession } from './document-session.ts';
import { selectedText } from './selection-state.ts';
import { copyCellsTo, tablePasteFrom } from './table-clipboard.ts';
import { readSlidesClipboard, slidesClipboard, SLIDES_MIME } from './slide-clipboard.ts';
import { textItemOf } from './table-cells.ts';

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
  /** Insert text at the caret, as typing does. */
  insertText(text: string): void;
  refuse(reason: string): void;
  /** The slides selected in the slide list, in order. */
  selectedSlides(): readonly number[];
  /** Run a command: a paste, a cleared cell range (a cut of cells), a picture. */
  exec(command: EditorCommand): void;
}

/** Picture types a paste inserts. */
const PICTURE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp']);
/** The largest picture a paste reads, as the picture command's media limit. */
const MAX_PICTURE = 50 * 1024 * 1024;

/** The clipboard handlers of one editor. */
export function createClipboardHost(editor: ClipboardEditor) {
  let lastCopied: ClipboardPayload | null = null;
  /** The picture of the last copy, to know it again in pasted HTML. */
  let lastPicture: ShapesPicture | null = null;
  let lastParsed: { raw: string; payload: unknown } = { raw: '', payload: null };
  /** Counts copies of any kind, so a late picture never replaces a later copy of text or slides. */
  let copies = 0;
  const pictures = createPictureCache();

  /** The key of the selected shapes' picture: the deck, its revision, the slide, the shapes. */
  const pictureKey = (): string | null => {
    const session = editor.session();
    const selection = editor.selection();
    if (!session || selection.kind !== 'shapes' || selection.group) return null;
    return `${session.id}:${session.store.revision}:${editor.slide()}:${selection.nodeIds.join(',')}`;
  };

  /** Start drawing the selected shapes' picture, so a copy finds it ready. */
  const warm = (): Promise<ShapesPicture | null> | null => {
    const key = pictureKey();
    const session = editor.session();
    const lay = editor.layout();
    if (!key || !session || !lay) return null;
    const selection = editor.selection();
    return pictures.warm(key, () => drawShapesPicture(session, lay, selection));
  };

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
      lastPicture = null;
      // Paste's enabled state reads `lastCopied`: tell the chrome it changed.
      editor.emit();
    }
    return payload;
  };

  /** The text of selected text boxes as formatted paragraphs, or null when one is not a text box. */
  const textBoxParagraphs = (): RichParagraph[] | null => {
    const lay = editor.layout();
    const selection = editor.selection();
    if (!lay || selection.kind !== 'shapes') return null;
    const out: RichParagraph[] = [];
    for (const id of selection.nodeIds) {
      const item = findItem(lay.items, id);
      const el = item?.kind === 'shape' ? item.element : undefined;
      if (!item || item.kind !== 'shape' || !el?.textBox || el.fill.type !== 'none' || !el.text)
        return null;
      const end = el.text.paragraphs.length - 1;
      const all = {
        kind: 'text' as const,
        nodeId: id,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: end, offset: item.text?.paragraphLengths[end] ?? 0 },
      };
      out.push(...richFromSelection(item, all));
    }
    return out.length ? out : null;
  };

  /**
   * HTML for copied shapes: the text of text boxes as text, which other applications paste as a
   * text box; else a picture of the shapes. Without a picture drawn yet, it is drawn and written
   * after the copy, where the browser allows it.
   */
  const shapesHtml = (data: DataTransfer, payload: ClipboardPayload): void => {
    const text = clipboardText(payload);
    const boxes = textBoxParagraphs();
    if (boxes) return data.setData('text/html', richHtml(boxes));
    const key = pictureKey();
    const ready = key ? pictures.ready(key) : null;
    if (ready) {
      lastPicture = ready;
      return data.setData('text/html', pictureHtml(ready, text));
    }
    const drawing = warm();
    if (drawing) writeLater(drawing, text, payload);
  };

  /** Write the picture when it is drawn, with the rest of the copy, through the async clipboard. */
  const writeLater = (
    drawing: Promise<ShapesPicture | null>,
    text: string,
    payload: ClipboardPayload
  ): void => {
    const clip = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    if (!clip?.write || typeof ClipboardItem === 'undefined') return;
    const mine = copies;
    const html = drawing.then((picture) => {
      // A later copy replaced this one: leave the clipboard to it.
      if (!picture || mine !== copies || lastCopied !== payload) throw new Error('stale');
      lastPicture = picture;
      return new Blob([pictureHtml(picture, text)], { type: 'text/html' });
    });
    const items: Record<string, Blob | Promise<Blob>> = {
      'text/html': html,
      'text/plain': new Blob([text], { type: 'text/plain' }),
    };
    const web = `web ${CLIPBOARD_MIME}`;
    if (ClipboardItem.supports?.(web))
      items[web] = new Blob([JSON.stringify(payload)], { type: web });
    // A browser that refuses keeps what the copy event wrote.
    clip.write([new ClipboardItem(items)]).catch(() => undefined);
  };

  /** Paste formatted paragraphs: at the caret, or in a new text box. */
  const pasteParagraphs = (paragraphs: readonly RichParagraph[]): void =>
    editor.exec({ type: 'pasteText', paragraphs: paragraphSpecs(paragraphs) });

  /** Insert a picture file from the clipboard, at the size its resolution gives. */
  const pastePicture = async (file: File): Promise<void> => {
    if (file.size === 0 || file.size > MAX_PICTURE) return editor.refuse('errors.clipboard');
    // The picture goes on the slide the paste was made on: another slide by then drops it.
    const at = `${editor.session()?.id}:${editor.slide()}`;
    const bytes = new Uint8Array(await file.arrayBuffer());
    let width = 640;
    let height = 480;
    try {
      const bitmap = await createImageBitmap(file, { imageOrientation: 'none' });
      width = bitmap.width;
      height = bitmap.height;
      bitmap.close();
    } catch {
      // A picture the browser cannot decode keeps the 4:3 default; the command checks it.
    }
    // The command counts pixels at 96 per inch: a 144 dpi screenshot is two-thirds as large.
    const scale = 96 / (pngResolution(bytes) ?? 96);
    if (`${editor.session()?.id}:${editor.slide()}` !== at) return;
    editor.exec({
      type: 'insertImage',
      bytes,
      mime: file.type,
      width: width * scale,
      height: height * scale,
    });
  };

  /** A picture file on the clipboard. */
  const pictureFile = (data: DataTransfer): File | undefined => {
    for (const item of Array.from(data.items ?? []))
      if (item.kind === 'file' && PICTURE_TYPES.has(item.type)) {
        const file = item.getAsFile();
        if (file) return file;
      }
    return Array.from(data.files ?? []).find((f) => PICTURE_TYPES.has(f.type));
  };

  return {
    get lastCopied(): ClipboardPayload | null {
      return lastCopied;
    },
    copy,
    warm,
    /** Fill a clipboard event with the slides selected in the slide list. */
    copySlidesTo(data: DataTransfer, cut: boolean): boolean {
      copies += 1;
      const session = editor.session();
      const slides = editor.selectedSlides();
      if (!session || !slides.length) return false;
      const json = slidesClipboard(session, slides);
      if (!json) return (editor.refuse('errors.clipboard'), true);
      data.setData(SLIDES_MIME, json);
      if (cut && !editor.readOnly()) editor.exec({ type: 'deleteSlides', indices: [...slides] });
      return true;
    },
    /** Fill a clipboard event: with the selected slides when it comes from the slide list. */
    copyTo(data: DataTransfer, cut: boolean, slideList = false): boolean {
      if (slideList) return this.copySlidesTo(data, cut);
      copies += 1;
      const selection = editor.selection();
      // A cell range: its text as tab-separated text and as this editor's cells format.
      if (copyCellsTo(data, editor.layout(), selection)) {
        if (cut && !editor.readOnly()) editor.exec({ type: 'delete' });
        return true;
      }
      if (selection.kind === 'text') {
        const text = selectedText(editor.layout(), selection);
        if (!text) return false;
        data.setData('text/plain', text);
        const session = editor.session();
        const part = session?.info.slides[editor.slide()]?.partName;
        const item = textItemOf(editor.layout(), selection);
        if (session && part) {
          const paragraphs = copyTextParagraphs(session.probeBatch(), part, selection);
          if (paragraphs.length)
            data.setData(TEXT_MIME, JSON.stringify({ version: 1, paragraphs }));
        }
        if (item) data.setData('text/html', richHtml(richFromSelection(item, selection)));
        if (cut) editor.deleteText();
        return true;
      }
      const payload = copy();
      if (!payload) return false;
      data.setData(CLIPBOARD_MIME, JSON.stringify(payload));
      data.setData('text/plain', clipboardText(payload));
      shapesHtml(data, payload);
      if (cut) editor.cutShapes();
      return true;
    },
    /**
     * Paste from a clipboard event. Plain text into text editing is left to the browser (it
     * arrives as typed text), and so is a clipboard with nothing the editor takes.
     */
    pasteFrom(data: DataTransfer): boolean {
      if (editor.readOnly()) return false;
      // Copied cells or tab-separated text fill cells; one copied cell goes in at the caret.
      const table = tablePasteFrom(data, editor.selection());
      if (table && 'cells' in table)
        return (editor.exec({ type: 'pasteCells', texts: table.cells }), true);
      if (table) return (editor.insertText(table.text), true);
      // Copied slides go in after the current slide.
      const slides = data.getData(SLIDES_MIME);
      if (slides) {
        const read = readSlidesClipboard(slides);
        if (!read) return (editor.refuse('errors.clipboard'), true);
        editor.exec({ type: 'pasteSlides', ...read });
        return true;
      }
      const editing = editor.selection().kind === 'text';
      // This editor's formatted text.
      const ownText = data.getData(TEXT_MIME);
      if (ownText) {
        const parsed = parseClipboardText(ownText) as { paragraphs?: unknown } | null;
        const paragraphs = cleanParagraphs(parsed?.paragraphs);
        if (!paragraphs) return (editor.refuse('errors.clipboard'), true);
        editor.exec({ type: 'pasteText', paragraphs });
        return true;
      }
      const html = data.getData('text/html');
      if (!editing) {
        // A payload parsed once is cached by its text, so a paste series keeps one identity.
        const read = readClipboard(data, lastCopied, (raw) => {
          if (raw === lastParsed.raw) return lastParsed.payload;
          const payload = parseClipboardText(raw);
          if (payload) lastParsed = { raw, payload };
          return payload;
        });
        if (read && 'refused' in read) return (editor.refuse('errors.clipboard'), true);
        if (read)
          return (editor.exec({ type: 'paste', payload: read.payload as ClipboardPayload }), true);
        // The picture of the last copy, written after the copy event: paste the shapes.
        if (lastCopied && lastPicture && html.includes(pictureMark(lastPicture)))
          return (editor.exec({ type: 'paste', payload: lastCopied }), true);
      }
      const rich = html ? richFromHtml(html) : null;
      const file = pictureFile(data);
      // HTML with text wins over a picture of it; a picture with no text is a picture.
      if (rich && rich.some((p) => p.runs.some((r) => r.text.trim()))) {
        pasteParagraphs(rich);
        return true;
      }
      if (file) {
        void pastePicture(file);
        return true;
      }
      if (editing) return false;
      const text = data.getData('text/plain');
      if (text.trim()) {
        pasteParagraphs(richFromText(text));
        return true;
      }
      return false;
    },
  };
}
