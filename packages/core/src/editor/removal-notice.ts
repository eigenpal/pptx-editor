// The last Delete of shapes or slides (`snapshot.lastRemoval`), so chrome can offer to undo it.
// It holds while the Delete is the newest edit: the next applied edit, an undo or redo, and a
// load clear it. Cut deletes through the same command, but the content moved to the clipboard,
// so it records none.

import type { EditorCommand, EditorRemoval, EditorSelection } from '../contracts/index.ts';

/** What a Delete command removes; null for any other command. */
function removalOf(
  command: EditorCommand,
  selection: EditorSelection
): Omit<EditorRemoval, 'serial'> | null {
  if (command.type === 'deleteSlide') return { kind: 'slides', count: 1 };
  if (command.type === 'deleteSlides')
    return { kind: 'slides', count: new Set(command.indices).size };
  if (command.type === 'delete' && selection.kind === 'shapes')
    return { kind: 'shapes', count: selection.nodeIds.length };
  return null;
}

const notices = new WeakMap<object, RemovalNotice>();

/**
 * Run `run` as the delete of a Cut on `editor`: chrome that copies first and then deletes (the
 * Cut menu item) reports no removal, as the surface's cut event does. @internal
 */
export function asCut<T>(editor: object, run: () => T): T {
  const notice = notices.get(editor);
  return notice ? notice.cut(run) : run();
}

/** The removal the editor reports. @internal */
export class RemovalNotice {
  #last: EditorRemoval | null = null;
  #serial = 0;
  #cutting = false;

  /** Record that this notice is `editor`'s, for `asCut`; returns `editor`. */
  attach<T extends object>(editor: T): T {
    notices.set(editor, this);
    return editor;
  }

  get current(): EditorRemoval | null {
    return this.#last;
  }

  clear(): void {
    this.#last = null;
  }

  /** What `command` would remove from `selection`, read before it runs. */
  pending(
    command: EditorCommand,
    selection: EditorSelection
  ): Omit<EditorRemoval, 'serial'> | null {
    return this.#cutting ? null : removalOf(command, selection);
  }

  /** Report a removal that applied. */
  record(removed: Omit<EditorRemoval, 'serial'> | null): void {
    if (removed) this.#last = { ...removed, serial: (this.#serial += 1) };
  }

  /** Run a Cut's delete: it reports no removal. */
  cut<T>(run: () => T): T {
    this.#cutting = true;
    try {
      return run();
    } finally {
      this.#cutting = false;
    }
  }
}
