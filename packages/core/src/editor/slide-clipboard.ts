// Slides on the clipboard: a copy in the slide list writes a presentation with the copied
// slides, and their indexes in it, so a paste in this editor, another editor on the page, or
// another tab puts in the same slides (`pasteSlides`). The file is the presentation as Save
// writes it, less the slides that were not copied (with the parts only they used), so a copy
// of one slide of a long deck writes and reads one slide. A paste reads it as untrusted bytes
// with the limits of any opened file.

import { writePptx, type PptxPackage } from '../store/index.ts';
import { fromBase64, toBase64 } from '../automation/commands/base64.ts';
import { OpBatch } from '../automation/commands/op-batch.ts';
import { deleteSlidesAt } from '../automation/commands/slide-commands.ts';
import { repairTablesForSave } from '../automation/commands/table-repair.ts';
import type { DocumentSession } from './document-session.ts';

/** The clipboard type for slides from this editor. */
export const SLIDES_MIME = 'application/x-pptx-editor-slides+json';

/** The largest presentation a slide copy carries, in bytes. */
const MAX_FILE = 50 * 1024 * 1024;
const MAX_TEXT = Math.ceil(MAX_FILE / 3) * 4 + 64 * 1024;
const MAX_SLIDES = 10_000;

/** `pkg` with only the slides `keep` (distinct indexes): the others go with their parts. */
function withSlides(pkg: PptxPackage, keep: readonly number[], count: number): PptxPackage {
  const kept = new Set(keep);
  const drop: number[] = [];
  for (let i = count - 1; i >= 0; i -= 1) if (!kept.has(i)) drop.push(i);
  if (drop.length === 0) return pkg;
  const batch = new OpBatch(pkg);
  deleteSlidesAt(batch, drop);
  return batch.pkg;
}

/**
 * The clipboard text for slides `slides` of the session's presentation, or null when it is too
 * large or does not save.
 */
export function slidesClipboard(
  session: DocumentSession,
  slides: readonly number[]
): string | null {
  const count = session.info.slides.length;
  const copied = [...new Set(slides)].filter((i) => i >= 0 && i < count).sort((a, b) => a - b);
  if (copied.length === 0) return null;
  // One save per revision and set of slides: copying them again before an edit reuses it.
  const pkg = session.store.package;
  const key = copied.join(',');
  const files = saved.get(pkg) ?? new Map<string, string | null>();
  saved.set(pkg, files);
  let file = files.get(key);
  if (file === undefined) {
    // A deck with a slide that did not parse has no file to copy, as it has none to save.
    if (session.damage) return null;
    let bytes: Uint8Array;
    try {
      bytes = writePptx(repairTablesForSave(withSlides(pkg, copied, count)));
    } catch {
      return null;
    }
    file = bytes.length > MAX_FILE ? null : toBase64(bytes);
    files.set(key, file);
  }
  // The copied slides keep their order in the file, which holds only them.
  const indexes = slides.map((i) => copied.indexOf(i)).filter((i) => i >= 0);
  return file === null ? null : JSON.stringify({ version: 1, slides: indexes, file });
}

/** The saved files of each package that a slide copy saved, by slides, as base64. */
const saved = new WeakMap<object, Map<string, string | null>>();

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
