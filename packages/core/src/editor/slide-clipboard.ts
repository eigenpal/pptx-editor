// Slides on the clipboard: a copy in the slide list writes the presentation as a file with the
// indexes of the copied slides, so a paste in this editor, another editor on the page, or
// another tab puts in the same slides (`pasteSlides`). The file is the saved presentation, as
// Save writes it; a paste reads it as untrusted bytes with the limits of any opened file.

import { writePptx } from '../store/index.ts';
import { fromBase64, toBase64 } from '../automation/commands/base64.ts';
import { repairTablesForSave } from '../automation/commands/table-repair.ts';
import type { DocumentSession } from './document-session.ts';

/** The clipboard type for slides from this editor. */
export const SLIDES_MIME = 'application/x-pptx-editor-slides+json';

/** The largest presentation a slide copy carries, in bytes. */
const MAX_FILE = 50 * 1024 * 1024;
const MAX_TEXT = Math.ceil(MAX_FILE / 3) * 4 + 64 * 1024;
const MAX_SLIDES = 10_000;

/** The clipboard text for slides `slides` of the session's presentation, or null when too large. */
export function slidesClipboard(
  session: DocumentSession,
  slides: readonly number[]
): string | null {
  // One save per revision: copying slides again before an edit reuses it.
  const pkg = session.store.package;
  let file = saved.get(pkg);
  if (file === undefined) {
    const bytes = writePptx(repairTablesForSave(pkg));
    file = bytes.length > MAX_FILE ? null : toBase64(bytes);
    saved.set(pkg, file);
  }
  return file === null ? null : JSON.stringify({ version: 1, slides, file });
}

/** The saved file of each package a slide copy saved, as base64 (null: too large). */
const saved = new WeakMap<object, string | null>();

/** Read slides from clipboard text: the file and the slide indexes, or null. */
export function readSlidesClipboard(
  raw: string
): { readonly file: Uint8Array; readonly slides: readonly number[] } | null {
  if (!raw || raw.length > MAX_TEXT) return null;
  let v: unknown;
  try {
    v = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof v !== 'object' || v === null) return null;
  const { version, slides, file } = v as Record<string, unknown>;
  if (version !== 1 || typeof file !== 'string' || !Array.isArray(slides)) return null;
  if (slides.length === 0 || slides.length > MAX_SLIDES) return null;
  if (!slides.every((i) => Number.isInteger(i) && (i as number) >= 0)) return null;
  const bytes = fromBase64(file);
  return bytes && bytes.length ? { file: bytes, slides: slides as number[] } : null;
}
