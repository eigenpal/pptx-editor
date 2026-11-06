// The slide rail's state: which slides are selected, and the summaries the rail paints.
//
// Selection in the slide pane. A click selects one slide; Ctrl/Cmd+click adds or
// removes one slide; Shift+click selects the range from the anchor (the last slide clicked
// without Shift) to the clicked slide. Slides are held by part name, which never changes for a
// slide and is never reused in a session, so a selection survives moves, inserts, and undo.

import type { EditorSnapshot, SlideSelectMode } from '../contracts/index.ts';
import { changedPartNames, presentationInfo, type PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { renderKey } from './slide-dependencies.ts';

type SlideRef = { readonly partName: string };

export class SlideSelection {
  #ids = new Set<string>();
  #anchor: string | null = null;
  #version = 0;
  #last: {
    slides: readonly SlideRef[];
    current: number;
    version: number;
    out: readonly number[];
  } | null = null;

  /**
   * A number that changes whenever the selected part names change. With the slide list and
   * the current slide, it decides `indices`, so a cache keys on it in place of the indices.
   */
  get version(): number {
    return this.#version;
  }

  /**
   * Selected slide indices, ascending. The current slide is always one of them. The same
   * array until the slide list, the current slide, or the selection changes: do not mutate it.
   */
  indices(slides: readonly SlideRef[], current: number): number[] {
    const last = this.#last;
    if (last?.slides === slides && last.current === current && last.version === this.#version)
      return last.out as number[];
    const out: number[] = [];
    slides.forEach((s, i) => {
      if (i === current || this.#ids.has(s.partName)) out.push(i);
    });
    this.#last = { slides, current, version: this.#version, out };
    return out;
  }

  #select(ids: Set<string>, anchor: string | null): void {
    this.#ids = ids;
    this.#anchor = anchor;
    this.#version += 1;
  }

  /** Select only the slide with this part name, and make it the anchor. */
  reset(partName: string | undefined): void {
    this.#select(new Set(partName ? [partName] : []), partName ?? null);
  }

  /** Select these slides; the first becomes the anchor. */
  set(slides: readonly SlideRef[], indices: readonly number[]): void {
    const names = indices.map((i) => slides[i]?.partName).filter((n): n is string => !!n);
    this.#select(new Set(names), names[0] ?? null);
  }

  /** Apply a click on slide `index`. Returns the new current slide. */
  click(
    slides: readonly SlideRef[],
    current: number,
    index: number,
    mode: SlideSelectMode
  ): number {
    const name = slides[index]?.partName;
    if (!name) return current;
    if (mode === 'replace') {
      this.reset(name);
      return index;
    }
    const selected = this.indices(slides, current);
    if (mode === 'range') {
      const anchorAt = slides.findIndex((s) => s.partName === this.#anchor);
      const from = anchorAt >= 0 ? anchorAt : current;
      const [lo, hi] = from <= index ? [from, index] : [index, from];
      const ids = new Set(slides.slice(lo, hi + 1).map((s) => s.partName));
      this.#select(ids, anchorAt < 0 ? slides[from].partName : this.#anchor);
      return index;
    }
    // Toggle: the last selected slide stays selected.
    if (selected.includes(index) && selected.length > 1) {
      const rest = selected.filter((i) => i !== index);
      this.#select(new Set(rest.map((i) => slides[i].partName)), this.#anchor);
      if (index !== current) return current;
      return rest.reduce((best, i) => (Math.abs(i - index) < Math.abs(best - index) ? i : best));
    }
    this.#select(new Set([...selected.map((i) => slides[i].partName), name]), name);
    return index;
  }
}

/**
 * Slide summaries for the rail, recomputed only when the package or the face set changed. A
 * summary keeps its object while its slide renders the same (`renderKey`), so the rail repaints
 * only the thumbnails of slides an edit changed.
 *
 * After an edit that keeps the slide list, only the slides that read a changed part get a new
 * key: each summary remembers the parts its key read (`slideDependencies`). The result is the
 * one a full pass gives.
 */
export function createSlideSummaries(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined
) => EditorSnapshot['slides'] {
  let cachedSlides: EditorSnapshot['slides'] = [];
  let slidesOf: { pkg: PptxPackage | undefined; fonts: number } = { pkg: undefined, fonts: -1 };
  /** The parts each slide's key read, by slide index, for `slidesOf.pkg`. */
  let readParts: (readonly string[])[] = [];
  return (session, pkg) => {
    const fonts = session?.fontGeneration ?? -1;
    if (slidesOf.pkg === pkg && slidesOf.fonts === fonts) return cachedSlides;
    const before = slidesOf;
    slidesOf = { pkg, fonts };
    const prev = cachedSlides;
    if (!session || !pkg) {
      readParts = [];
      return (cachedSlides = prev.length ? [] : prev);
    }
    const slides = session.info.slides;
    // Slides whose parts did not change keep their key: the slide list and the fonts are the
    // same, and every part the key read is the same object.
    const changed =
      before.pkg && before.fonts === fonts && presentationInfo(before.pkg).slides === slides
        ? changedPartNames(before.pkg, pkg, 64)
        : null;
    const dirty = changed ? new Set(changed) : null;
    const parts: (readonly string[])[] = [];
    const next = slides.map((s, index) => {
      const old = prev[index];
      const read = readParts[index];
      if (dirty && old && read && !read.some((name) => dirty.has(name))) {
        parts.push(read);
        return old;
      }
      const names: string[] = [];
      const key = renderKey(pkg, index, fonts, names) ?? '';
      parts.push(names);
      // Part names are never reused in a session, so the part name is the slide's identity.
      return old &&
        old.key === key &&
        old.hidden === s.hidden &&
        old.id === s.partName &&
        old.layoutId === s.layoutPart
        ? old
        : {
            index,
            id: s.partName,
            hidden: s.hidden,
            key,
            ...(s.layoutPart ? { layoutId: s.layoutPart } : {}),
          };
    });
    readParts = parts;
    const same = next.length === prev.length && next.every((s, i) => s === prev[i]);
    if (!same) cachedSlides = next;
    return cachedSlides;
  };
}
