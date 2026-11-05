// The slide rail's state: which slides are selected, and the summaries the rail paints.
//
// Selection in the slide pane. A click selects one slide; Ctrl/Cmd+click adds or
// removes one slide; Shift+click selects the range from the anchor (the last slide clicked
// without Shift) to the clicked slide. Slides are held by part name, which never changes for a
// slide and is never reused in a session, so a selection survives moves, inserts, and undo.

import type { EditorSnapshot, SlideSelectMode } from '../contracts/index.ts';
import type { PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { renderKey } from './slide-dependencies.ts';

type SlideRef = { readonly partName: string };

export class SlideSelection {
  #ids = new Set<string>();
  #anchor: string | null = null;

  /** Selected slide indices, ascending. The current slide is always one of them. */
  indices(slides: readonly SlideRef[], current: number): number[] {
    const out: number[] = [];
    slides.forEach((s, i) => {
      if (i === current || this.#ids.has(s.partName)) out.push(i);
    });
    return out;
  }

  /** Select only the slide with this part name, and make it the anchor. */
  reset(partName: string | undefined): void {
    this.#ids = new Set(partName ? [partName] : []);
    this.#anchor = partName ?? null;
  }

  /** Select these slides; the first becomes the anchor. */
  set(slides: readonly SlideRef[], indices: readonly number[]): void {
    const names = indices.map((i) => slides[i]?.partName).filter((n): n is string => !!n);
    this.#ids = new Set(names);
    this.#anchor = names[0] ?? null;
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
      this.#ids = new Set(slides.slice(lo, hi + 1).map((s) => s.partName));
      if (anchorAt < 0) this.#anchor = slides[from].partName;
      return index;
    }
    // Toggle: the last selected slide stays selected.
    if (selected.includes(index) && selected.length > 1) {
      const rest = selected.filter((i) => i !== index);
      this.#ids = new Set(rest.map((i) => slides[i].partName));
      if (index !== current) return current;
      return rest.reduce((best, i) => (Math.abs(i - index) < Math.abs(best - index) ? i : best));
    }
    this.#ids = new Set([...selected.map((i) => slides[i].partName), name]);
    this.#anchor = name;
    return index;
  }
}

/**
 * Slide summaries for the rail, recomputed only when the package or the face set changed. A
 * summary keeps its object while its slide renders the same (`renderKey`), so the rail repaints
 * only the thumbnails of slides an edit changed.
 */
export function createSlideSummaries(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined
) => EditorSnapshot['slides'] {
  let cachedSlides: EditorSnapshot['slides'] = [];
  let slidesOf: { pkg: PptxPackage | undefined; fonts: number } = { pkg: undefined, fonts: -1 };
  return (session, pkg) => {
    const fonts = session?.fontGeneration ?? -1;
    if (slidesOf.pkg === pkg && slidesOf.fonts === fonts) return cachedSlides;
    slidesOf = { pkg, fonts };
    const prev = cachedSlides;
    const next =
      session && pkg
        ? session.info.slides.map((s, index) => {
            const key = renderKey(pkg, index, fonts) ?? '';
            const old = prev[index];
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
          })
        : [];
    const same = next.length === prev.length && next.every((s, i) => s === prev[i]);
    if (!same) cachedSlides = next;
    return cachedSlides;
  };
}
