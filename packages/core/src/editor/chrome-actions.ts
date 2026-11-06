// How a menu row or a toolbar control runs, for every kind of slot: an engine command, a value
// command, or an action the adapter owns (a registry dialog, the clipboard, the slide show, the
// find bar, a file picker). The decision and the clipboard work are framework-free, so the
// React and Vue menus run a row the same way; each adapter supplies only the hooks into its
// own dialog host, slide show, and find bar.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { ChromeDialogKey, ChromeSlotId } from './chrome-controls.ts';
import {
  chromeControlForSlot,
  commandForSlotValue,
  optionRefusal,
  runToolbarCommand,
} from './chrome-controls.ts';
import { chromeCommandState, type ChromeCommandState } from './adapter-support.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { focusSurfaceOf, surfaceInputOf, surfaceRootOf } from './surface-registry.ts';
import { warnNoAltTextPopup, warnNoLinkPopup } from './dev-warnings.ts';
import { asCut } from './removal-notice.ts';
import { tablePaste } from './table-clipboard.ts';
import type { ToolbarShownValue, ToolbarValueMap } from './toolbar-values.ts';
import { canRenameSectionInPlace, runSectionSlot } from './section-renamers.ts';

/** What an adapter does for the slots it runs itself. @internal */
export interface ChromeActionHost {
  /** Open a `popups` dialog. */
  openDialog(key: ChromeDialogKey): void;
  /** Start the slide show. */
  startSlideshow(from: 'start' | 'current'): void;
  /** Open the find bar, to find or to replace. */
  openFind(replace: boolean): void;
  /** Open the link popup. */
  openLink(): void;
  /** Ask the host for an image file to insert, or to change the selected picture to. */
  pickImage(target?: 'replaceImage'): void;
  /**
   * Why the host cannot run `slot` now (its `popups` entry is `false`), or null. A row then
   * shows disabled with this reason instead of doing nothing. `dialog` names the registry dialog
   * of a dialog row (Format > Bullets and numbering…), whose own entry decides.
   */
  unavailable?(slot: ChromeSlotId, dialog?: ChromeDialogKey): string | null;
}

/** The `popups` entry each adapter-owned slot opens. @internal */
export const SLOT_POPUP_KEYS: Readonly<Partial<Record<ChromeSlotId, string>>> = {
  'text.link': 'hyperlink',
  'find.find': 'find',
  'find.replace': 'find',
  'file.pageSetup': 'pageSetup',
  'slide.background': 'background',
  'image.altText': 'imageAltText',
  'image.properties': 'imageProperties',
  'paragraph.dialog': 'paragraph',
  'text.dialog': 'font',
};

/** Whether the adapter runs `slot` itself (`runChromeRow` sends it to the host). @internal */
export function isAdapterSlot(slot: ChromeSlotId): boolean {
  const kind = chromeControlForSlot(slot)?.state?.kind;
  return (
    kind === 'dialog' ||
    kind === 'slideshow' ||
    kind === 'panel' ||
    kind === 'load' ||
    kind === 'replace' ||
    kind === 'clipboard' ||
    slot === 'text.link'
  );
}

/**
 * Whether an adapter slot needs the adapter's host (a dialog, the slide show, find, the link
 * popup). The clipboard and a picture run from core alone. @internal
 */
export function adapterSlotNeedsHost(slot: ChromeSlotId): boolean {
  const kind = chromeControlForSlot(slot)?.state?.kind;
  return kind === 'dialog' || kind === 'slideshow' || kind === 'panel' || slot === 'text.link';
}

/**
 * The state and runner of `useEditorCommand` for one slot, shared by the React and Vue hooks.
 * `host` is the adapter's host, or null where there is none (outside `PptxEditor.Root`, or an
 * adapter without one): a slot that needs a host is then disabled with `missingHost` as the
 * reason. `text.link` with a value (a URL, or null to unlink) runs `setLink` directly, and
 * `image.altText` with a string runs `setAltText`, as DOCX Editor's `image.altText`. The popup
 * availability (`unavailable`) and the host never apply to such a value: hiding the popup never
 * turns off the operation. They gate only the opener (`execute()` with no value). @internal
 */
export function editorCommandFor(
  editor: PptxEditorInstance | null,
  slot: ChromeSlotId,
  host: ChromeActionHost | null,
  missingHost: string
): { readonly state: ChromeCommandState; readonly execute: (value?: unknown) => boolean } {
  const adapter = isAdapterSlot(slot);
  // `text.link` and `image.altText` run a value from core (`execute(url)`,
  // `execute(description)`); only their popup needs the host.
  const valueSlot = slot === 'text.link' || slot === 'image.altText';
  const needsHost = adapter && adapterSlotNeedsHost(slot) && !valueSlot;
  const base = chromeCommandState(editor, slot);
  // The popup's availability gates only the opener (a toolbar button, a menu row).
  const gated = valueSlot ? base : chromeRowState(editor, slot, undefined, host ?? undefined);
  const state =
    needsHost && !host ? { ...base, isEnabled: false, disabledReason: missingHost } : gated;
  const execute = (value?: unknown): boolean => {
    if (!editor) return false;
    // Rename Section and Add Section open the name field in a rail or a sorter.
    if (isSectionSlot(slot))
      return chromeRowState(editor, slot).isEnabled && runSectionSlot(editor, slot);
    // A dialog slot given its fields (`execute(fields)` on `paragraph.dialog`) runs the
    // dialog's command directly, as `text.link` with a URL does.
    const fields =
      typeof value === 'object' &&
      value !== null &&
      chromeControlForSlot(slot)?.state?.kind === 'dialog';
    if (
      !adapter ||
      (slot === 'text.link' && value !== undefined) ||
      (slot === 'image.altText' && typeof value === 'string') ||
      fields
    )
      return runToolbarCommand(editor, slot, value).ok;
    if (!state.isEnabled) return false;
    // `execute()` with no URL opens the link popup: say so when none can open.
    if (slot === 'text.link' && (!host || host.unavailable?.(slot))) {
      warnNoLinkPopup();
      return false;
    }
    // `execute()` with no description opens the Alt Text dialog, which needs the host.
    if (slot === 'image.altText' && (!host || host.unavailable?.(slot))) {
      warnNoAltTextPopup();
      return false;
    }
    if (host) runChromeRow(editor, host, slot, typeof value === 'string' ? value : undefined);
    else if (chromeControlForSlot(slot)?.state?.kind === 'load') pickImageInto(editor);
    else if (chromeControlForSlot(slot)?.state?.kind === 'replace')
      pickImageInto(editor, document, 'replaceImage');
    else runClipboardSlot(editor, slot, null);
    return true;
  };
  return { state, execute };
}

/** The typed slots whose `ToolbarValueMap` type is a number; the registry keeps strings. */
const NUMERIC_SLOTS: ReadonlySet<string> = new Set([
  'font.size',
  'text.characterSpacing',
  'list.lineSpacing',
  'paragraph.columns',
  'shape.lineWidth',
]);

/** A registry or state value as its slot's type: a number for a numeric slot. */
function typedSlotValue(slot: string, raw: unknown): unknown {
  if (!NUMERIC_SLOTS.has(slot) || typeof raw !== 'string') return raw;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

/**
 * The choices of a typed value slot, from its registry options, for `useEditorValueCommand`'s
 * `options` (DOCX Editor's shape), as the slot's type. A slot with free values (`font.size`,
 * `shape.fill`) has none. @internal
 */
export function valueSlotOptions<K extends keyof ToolbarValueMap>(
  slot: K
): readonly ToolbarValueMap[K][] {
  // A dialog row ("Line spacing options…") opens a dialog: it is not a value.
  const options = (chromeControlForSlot(slot)?.options ?? []).filter((o) => !o.dialog);
  return options.map((option) => typedSlotValue(slot, option.value) as ToolbarValueMap[K]);
}

/**
 * The value a typed slot shows, as `ToolbarShownValue`: a number for a numeric slot, as DOCX
 * Editor's `useEditorValueCommand` converts it; null when mixed or unavailable. @internal
 */
export function valueSlotShown<K extends keyof ToolbarValueMap>(
  slot: K,
  raw: unknown
): ToolbarShownValue<K> | null {
  return (raw === undefined ? null : typedSlotValue(slot, raw)) as ToolbarShownValue<K> | null;
}

/** The state of a row: the slot's state, or, for a value row, the engine's answer for it. @internal */
export function chromeRowState(
  editor: PptxEditorInstance | null,
  slot: ChromeSlotId,
  value?: string,
  host?: Pick<ChromeActionHost, 'unavailable'>
): ChromeCommandState {
  const state = chromeCommandState(editor, slot);
  // A dialog row takes the slot's state, and its dialog's availability: the dialog runs the
  // command.
  const dialog =
    value === undefined
      ? undefined
      : chromeControlForSlot(slot)?.options?.find((o) => o.value === value)?.dialog;
  const off = host?.unavailable?.(slot, dialog) ?? null;
  if (off) return { ...state, isEnabled: false, disabledReason: off };
  // The name field is in a slide rail or a slide sorter: without one, Rename is off and says why.
  if (slot === 'section.rename' && state.isEnabled && !canRenameSectionInPlace(editor))
    return { ...state, isEnabled: false, disabledReason: 'disabled.noSectionPane' };
  if (value === undefined || !editor || dialog) return state;
  const command = commandForSlotValue(slot, value);
  if (!command) return { ...state, isEnabled: false, disabledReason: 'disabled.notWired' };
  const refused = optionRefusal(editor, command);
  return { ...state, isEnabled: refused === null, disabledReason: refused };
}

/**
 * Run a row of the menu bar or a context menu. A value row runs its value; a slot the adapter
 * owns goes to `host`; a value slot without a value (a gallery, a picker) has no action from a
 * menu and is left to its submenu, except New slide, which adds a slide. @internal
 */
export function runChromeRow(
  editor: PptxEditorInstance | null,
  host: ChromeActionHost,
  slot: ChromeSlotId,
  value?: string,
  surfaceInput?: HTMLElement | null
): void {
  if (!editor) return;
  const state = chromeControlForSlot(slot)?.state;
  if (!state) return;
  // A dialog row (Format > Bullets and numbering…, a gallery's last row) opens its dialog.
  const dialog = chromeControlForSlot(slot)?.options?.find((o) => o.value === value)?.dialog;
  if (value !== undefined && dialog) {
    host.openDialog(dialog);
    return;
  }
  if (value !== undefined) {
    runToolbarCommand(editor, slot, value);
    return;
  }
  switch (state.kind) {
    case 'dialog':
      host.openDialog(state.dialog);
      return;
    case 'slideshow':
      host.startSlideshow(state.from);
      return;
    case 'panel':
      host.openFind(slot === 'find.replace');
      return;
    case 'load':
      host.pickImage();
      return;
    case 'replace':
      host.pickImage('replaceImage');
      return;
    case 'clipboard':
      runClipboardSlot(editor, slot, surfaceInput ?? null);
      return;
    case 'value':
      if (slot === 'text.link') host.openLink();
      // New slide without a layout: the layout `newSlideLayout` picks (`addSlide`).
      else if (slot === 'slide.add') runToolbarCommand(editor, slot);
      return;
    case 'command':
      if (isSectionSlot(slot)) {
        if (chromeRowState(editor, slot).isEnabled) runSectionSlot(editor, slot);
      } else runToolbarCommand(editor, slot);
  }
}

/** Whether `slot` is Rename Section or Add Section, which chrome runs with the name field. */
const isSectionSlot = (slot: ChromeSlotId): slot is 'section.rename' | 'section.add' =>
  slot === 'section.rename' || slot === 'section.add';

/**
 * Cut, copy, paste, or select all. Copy and Cut go through the surface's own copy and cut
 * events, which write the system clipboard (the editor's format and plain text), so Ctrl+V
 * pastes what the menu copied. Where the browser refuses `execCommand`, shapes fall back to the
 * editor's own clipboard. Paste of shapes pastes the editor's last copy, because a browser lets
 * a page read the system clipboard only from Ctrl+V, never from a menu click; Paste into text
 * reads it with `navigator.clipboard.readText`, where the browser allows it. The input is
 * always this editor's own (`surfaceInputOf`), never the focused element's. @internal
 */
export function runClipboardSlot(
  editor: PptxEditorInstance,
  slot: ChromeSlotId,
  surfaceInput: HTMLElement | null
): void {
  if (slot === 'edit.selectAll') {
    editor.selectAll();
    return;
  }
  const sel = editor.snapshot().selection;
  // A cell range copies and pastes as text does: through the surface's own clipboard events.
  const text = sel.kind === 'text' || sel.kind === 'cells';
  const input = surfaceInputOf(editor) ?? surfaceInput;
  if (!text && (slot === 'edit.copy' || slot === 'edit.cut')) {
    // Shapes copy through the surface's copy event, which writes the system clipboard; where
    // the browser refuses, through the editor's own clipboard. Cut is that copy plus a delete,
    // and deletes only when the copy carried every selected shape: a chart, diagram, or object
    // frame the clipboard cannot carry is never lost.
    const before = editor.lastCopied;
    // A copy that carried nothing leaves the last copy as it was.
    const copied = copyThroughSurface(editor)
      ? editor.lastCopied !== before
        ? editor.lastCopied
        : null
      : editor.copy();
    const all = sel.kind === 'shapes' && copied?.shapes.length === sel.nodeIds.length;
    if (slot === 'edit.cut' && all) asCut(editor, () => editor.exec({ type: 'delete' }));
    return;
  }
  if (text) {
    if (slot === 'edit.paste') {
      void globalThis.navigator?.clipboard
        ?.readText()
        .then((data) => {
          if (!data) return;
          // Tab-separated text, or any text on a cell range, fills cells; one copied cell
          // goes in at the caret as text (`table-clipboard.ts`).
          const paste = tablePaste(editor.snapshot().selection, { text: data });
          if (paste && 'cells' in paste) editor.exec({ type: 'pasteCells', texts: paste.cells });
          else editor.input('text', paste ? paste.text : data);
        })
        .catch(() => undefined);
      return;
    }
    input?.focus({ preventScroll: true });
    input?.ownerDocument.execCommand(slot === 'edit.cut' ? 'cut' : 'copy');
    return;
  }
  if (slot === 'edit.paste' && editor.lastCopied)
    editor.exec({ type: 'paste', payload: editor.lastCopied });
}

/**
 * Copy the selected shapes through the surface's copy event (never its cut event, whose delete
 * does not check what the copy carried). False when no event fired.
 */
function copyThroughSurface(editor: PptxEditorInstance): boolean {
  const root = surfaceRootOf(editor);
  // The surface's own text input takes focus for the command: Chromium sends the copy event to
  // the node that holds the selection, so a focused frame alone would send it to the page (or
  // copy some other selected text). The surface then takes its normal focus back.
  const input = surfaceInputOf(editor);
  if (!root || !input) return false;
  input.focus({ preventScroll: true });
  let fired = false;
  const seen = (): void => void (fired = true);
  root.addEventListener('copy', seen, true);
  try {
    root.ownerDocument.execCommand('copy');
  } catch {
    // An old browser throws; the fallback below runs.
  } finally {
    root.removeEventListener('copy', seen, true);
    focusSurfaceOf(editor);
  }
  return fired;
}

/** The text input of the editor surface inside `scope`, for text copy and cut. @internal */
export function surfaceInputIn(scope: Element | null | undefined): HTMLElement | null {
  return scope?.querySelector<HTMLElement>('.pptx-surface__input') ?? null;
}

/** The picture types the chrome offers for Insert > Picture. @internal */
export const IMAGE_FILE_TYPES = 'image/png,image/jpeg,image/gif,image/svg+xml';

/**
 * The picture types a file picker offers: Change Picture takes raster images only (an SVG needs
 * a raster copy beside it, which the editor cannot draw). @internal
 */
export function imageFileTypes(target?: 'replaceImage'): string {
  return target === 'replaceImage' ? 'image/png,image/jpeg,image/gif' : IMAGE_FILE_TYPES;
}

/**
 * Insert a picture file the user chose, at its own pixel size (4:3 when the size cannot be
 * read, as for SVG); with `replaceImage`, replace the selected picture's image with it. @internal
 */
export async function insertImageFile(
  editor: PptxEditorInstance,
  file: File,
  target?: 'replaceImage'
): Promise<void> {
  // Change Picture changes the picture selected now: a selection that moves while the file
  // loads (another picture, another slide) refuses it, so the wrong picture never changes.
  const before = selectionKey(editor);
  const bytes = new Uint8Array(await file.arrayBuffer());
  let width: number | undefined;
  let height: number | undefined;
  try {
    // The stored pixels: a JPEG's EXIF orientation does not count.
    const bitmap = await createImageBitmap(file, { imageOrientation: 'none' });
    width = bitmap.width;
    height = bitmap.height;
    bitmap.close();
  } catch {
    // SVG and unknown sizes fall back to 4:3.
  }
  // Change Picture reads the size from the image when the browser cannot decode it; Insert
  // Picture falls back to 4:3.
  if (target === 'replaceImage') {
    if (selectionKey(editor) !== before) editor.reject('disabled.selectionChanged');
    else editor.exec({ type: target, bytes, mime: file.type, width, height });
  } else
    editor.exec({
      type: 'insertImage',
      bytes,
      mime: file.type,
      width: width ?? 640,
      height: height ?? 480,
    });
}

/** The slide and the selected elements, to tell whether the selection moved. */
function selectionKey(editor: PptxEditorInstance): string {
  const snap = editor.snapshot();
  const sel = snap.selection;
  return JSON.stringify([snap.currentSlide, sel.kind, sel.kind === 'shapes' ? sel.nodeIds : []]);
}

/**
 * Ask the user for a picture file and insert it (Insert > Picture from a menu), or change the
 * selected picture to it (Change Picture). @internal
 */
export function pickImageInto(
  editor: PptxEditorInstance,
  doc: Document = document,
  target?: 'replaceImage'
): void {
  const input = doc.createElement('input');
  input.type = 'file';
  input.accept = imageFileTypes(target);
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (file) void insertImageFile(editor, file, target);
  });
  input.click();
}
