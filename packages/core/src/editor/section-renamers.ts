// Which slide rails and slide sorters of each editor rename a section in place. Rename Section
// (`section.rename`) and Add Section (`section.add`, which asks for the new section's name)
// open the name field there. The toolbar, the menu bar, a context menu, and
// `useEditorCommand` all read this one registry, so each reports the same enabled state and
// runs the same way. The adapters register their rails and sorters.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { warnSectionRenameUnheard } from './dev-warnings.ts';

const renamers = new WeakMap<object, Set<(sectionId: string) => void>>();
const listeners = new Set<() => void>();
const changed = (): void => {
  for (const listen of listeners) listen();
};

/** Count `take` as a rail or sorter of `editor` that renames in place. Returns the undo. @internal */
export function registerSectionRenamer(
  editor: object,
  take: (sectionId: string) => void
): () => void {
  let set = renamers.get(editor);
  if (!set) renamers.set(editor, (set = new Set()));
  set.add(take);
  changed();
  return () => {
    set.delete(take);
    changed();
  };
}

/** Listen for rails and sorters that register or leave. Returns the undo. @internal */
export function subscribeSectionRenamers(listen: () => void): () => void {
  listeners.add(listen);
  return () => void listeners.delete(listen);
}

/** Whether a mounted rail or sorter of `editor` renames a section in place. @internal */
export function canRenameSectionInPlace(editor: object | null): boolean {
  return !!editor && !!renamers.get(editor)?.size;
}

/**
 * Open the name field of section `sectionId` in the rails and sorters of `editor`. When none is
 * mounted, nothing opens and a development warning says so. @internal
 */
export function requestSectionRenameIn(editor: object | null, sectionId: string): void {
  if (!editor) return;
  const set = renamers.get(editor);
  if (!set?.size) return warnSectionRenameUnheard();
  for (const take of set) take(sectionId);
}

/** The section that holds the current slide, or undefined. */
function currentSection(editor: PptxEditorInstance) {
  const s = editor.snapshot();
  return (s.sections ?? []).find(
    (x) => s.currentSlide >= x.firstSlide && s.currentSlide < x.firstSlide + x.slideCount
  );
}

/**
 * Run `section.rename` or `section.add` as chrome does: Rename opens the name field of the
 * current slide's section; Add adds a section at the current slide and, where a rail or sorter
 * can show the field, opens it for the new section. Returns whether it ran. @internal
 */
export function runSectionSlot(
  editor: PptxEditorInstance,
  slot: 'section.rename' | 'section.add'
): boolean {
  if (slot === 'section.rename') {
    const own = currentSection(editor);
    if (!own || !canRenameSectionInPlace(editor)) return false;
    requestSectionRenameIn(editor, own.id);
    return true;
  }
  const start = editor.snapshot();
  const before = new Set((start.sections ?? []).map((x) => x.id));
  if (!editor.exec({ type: 'addSection' }).ok) return false;
  // In a deck without sections, Add Section also adds "Default Section" for the slides before
  // the current slide. The name field opens on the section that starts at the current slide.
  const fresh = (editor.snapshot().sections ?? []).filter((x) => !before.has(x.id));
  const added = fresh.find((x) => x.firstSlide === start.currentSlide) ?? fresh.at(-1);
  if (added && canRenameSectionInPlace(editor)) requestSectionRenameIn(editor, added.id);
  return true;
}
