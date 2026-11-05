// A format armed at a caret for the text typed next, as DOCX Editor's `pendingFormats`: Bold or
// a font size at a word edge changes nothing until the user types. It is session state, never a
// store edit, so it draws nothing, makes no undo step, never reaches a peer or the file, and
// ends when the caret moves or the user undoes. Typing commits the armed ops and the text as
// one transaction. As in DOCX Editor's caret format history, the first undo removes the text
// and arms the format again at the caret (`rearm`); the next undo ends the arm. Enter,
// Shift+Enter, Backspace, and Delete carry the arm to the new caret, as DOCX Editor's
// `rearmPending` does; in an empty paragraph the paragraph mark takes the format with the edit.
//
// The arm keeps its commands, each made absolute when it arms (Bold becomes "bold on" or "bold
// off", a size step becomes the size), and replays them over the store package when that
// package changes under the caret (a peer's edit) or the caret moves with an edit.

import type { SlideLayout } from '../layout/index.ts';
import type {
  EditorCommand,
  EditorSelection,
  SelectionFormat,
  TableCellAddress,
  TextPoint,
} from '../contracts/index.ts';
import type { PptxOp, PptxPackage } from '../store/index.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { comparePoints } from '../automation/commands/text-model.ts';

/** Where a format is armed: a collapsed text caret on a slide part. */
export interface ArmedCaret {
  readonly part: string;
  readonly nodeId: string;
  readonly at: TextPoint;
  /** The table cell of the caret; `nodeId` is then the table's graphic frame. */
  readonly cell?: TableCellAddress;
}

interface Armed extends ArmedCaret {
  readonly commands: readonly EditorCommand[];
  /** The store package the ops apply to. */
  readonly base: PptxPackage;
  readonly ops: readonly PptxOp[];
  /** `base` after the ops: the preview that format reads and typing use. */
  readonly pkg: PptxPackage;
}

/** What the arm needs from the editor. */
export interface CaretArmHost {
  /** The store package now, or undefined with no document. */
  readonly package: () => PptxPackage | undefined;
  /** The collapsed text caret now, or null for any other selection. */
  readonly caret: () => ArmedCaret | null;
  /** Compile `command` at `caret` over `base`, with the toggle state read from `layout`. */
  readonly compileOn: (
    command: EditorCommand,
    base: PptxPackage,
    layout: SlideLayout | undefined,
    caret: ArmedCaret
  ) => OpBatch;
  /** Layout of the current slide over a package (cached by the session for the store's). */
  readonly layoutOf: (pkg: PptxPackage) => SlideLayout | undefined;
  /** The format read at `caret` in `layout`. */
  readonly formatAt: (layout: SlideLayout | undefined, caret: ArmedCaret) => SelectionFormat | null;
}

/** The caret of `selection` on `part` when it is a collapsed text caret, else null. @internal */
export function armedCaretOf(
  selection: EditorSelection,
  part: string | null | undefined
): ArmedCaret | null {
  return selection.kind === 'text' && part && comparePoints(selection.anchor, selection.focus) === 0
    ? {
        part,
        nodeId: selection.nodeId,
        at: selection.focus,
        ...(selection.cell ? { cell: selection.cell } : {}),
      }
    : null;
}

/** The collapsed text selection of a caret. @internal */
export const caretSelection = (c: ArmedCaret): EditorSelection => ({
  kind: 'text',
  nodeId: c.nodeId,
  anchor: c.at,
  focus: c.at,
  ...(c.cell ? { cell: c.cell } : {}),
});

/**
 * A relative command as the absolute one it amounts to at the armed caret, read from the
 * preview after it ran: a replay then cannot flip it (Bold armed off stays off on a bold word).
 */
function absoluteOf(command: EditorCommand, f: SelectionFormat | null): EditorCommand {
  const font = f?.font;
  if (!f || !font) return command;
  const set = (fields: Omit<Extract<EditorCommand, { type: 'setFontFormat' }>, 'type'>) =>
    ({ type: 'setFontFormat', ...fields }) as const;
  switch (command.type) {
    case 'toggleBold':
      return set({ bold: f.bold });
    case 'toggleItalic':
      return set({ italic: f.italic });
    case 'toggleUnderline':
      return font.underline ? set({ underline: font.underline }) : command;
    case 'toggleStrike':
      return font.strike ? set({ strike: font.strike }) : command;
    case 'toggleSuperscript':
    case 'toggleSubscript':
      return font.baseline !== null ? set({ baseline: font.baseline }) : command;
    case 'stepFontSize':
      return font.fontSizePt !== null ? set({ fontSizePt: font.fontSizePt }) : command;
    default:
      return command;
  }
}

const sameCaret = (a: ArmedCaret, b: ArmedCaret | null): boolean =>
  b !== null &&
  a.part === b.part &&
  a.nodeId === b.nodeId &&
  a.cell?.row === b.cell?.row &&
  a.cell?.column === b.cell?.column &&
  comparePoints(a.at, b.at) === 0;

/** The arm of one editor. @internal */
export class CaretArm {
  #armed: Armed | null = null;
  #preview: { pkg: PptxPackage; part: string; layout: SlideLayout | undefined } | null = null;
  /** While commands replay, a read of the arm (the toggle state) sees no arm. */
  #replaying = false;
  /** The armed commands typing took, by the history mark (`selectionBefore`) of its entry. */
  #typed = new WeakMap<object, readonly EditorCommand[]>();
  /** Bumped on every change, for caches keyed by the editor state. */
  version = 0;

  constructor(private readonly host: CaretArmHost) {}

  /** End the arm (undo, a commit, or a new document). */
  clear(): void {
    this.#preview = null;
    if (!this.#armed) return;
    this.#armed = null;
    this.version += 1;
  }

  /** End the arm when the caret has left it. Call before every publish of the state. */
  sync(): void {
    if (this.#armed && !sameCaret(this.#armed, this.host.caret())) this.clear();
  }

  /**
   * Run a selection update for a peer's edit (`followRemote`): an arm at the caret moves with
   * the caret, and replays over the new package when it is next read.
   */
  follow(update: () => void): void {
    const armed = this.#armed;
    const atCaret = armed !== null && sameCaret(armed, this.host.caret());
    update();
    const caret = this.host.caret();
    if (!armed || !atCaret || !caret) return;
    this.#armed = { ...armed, ...caret };
    this.version += 1;
  }

  /** The arm at the caret, replayed over the store package when that changed. */
  current(): Armed | null {
    const armed = this.#armed;
    if (!armed || this.#replaying) return null;
    const caret = this.host.caret();
    if (!caret || !sameCaret(armed, caret)) return (this.clear(), null);
    if (armed.base === this.host.package()) return armed;
    const replayed = this.#replay(caret, armed.commands);
    this.#armed = replayed;
    this.version += 1;
    return replayed;
  }

  /** The preview layout while a format is armed, for format reads; undefined otherwise. */
  layout(): SlideLayout | undefined {
    const armed = this.current();
    return armed ? this.#layoutOf(armed.pkg, armed.part) : undefined;
  }

  /**
   * Enter, Shift+Enter, Backspace, or Delete at an armed caret: the arm moves to `caret`, the
   * new caret in `base` (the edited package). In an empty paragraph (a new one after Enter, or
   * one a deletion emptied) its commands format the paragraph mark at once: the returned ops go
   * into the edit's transaction and the arm ends. Elsewhere, also inside a word, the arm stays
   * session state and nothing is returned.
   */
  carry(base: PptxPackage, caret: ArmedCaret): readonly PptxOp[] {
    const armed = this.current();
    if (!armed) return [];
    let pkg = base;
    const ops: PptxOp[] = [];
    let arms = true;
    this.#replaying = true;
    try {
      for (const command of armed.commands) {
        const batch = this.host.compileOn(command, pkg, this.#layoutOf(pkg, caret.part), caret);
        arms = batch.armedCaret;
        ops.push(...batch.ops);
        pkg = batch.pkg;
      }
    } catch (e) {
      if (!(e instanceof CommandRefusal)) throw e;
      this.clear();
      return [];
    } finally {
      this.#replaying = false;
    }
    if (!arms) {
      this.clear();
      return ops;
    }
    // Replayed over the committed package when it is next read.
    this.#armed = { ...armed, ...caret };
    this.version += 1;
    return [];
  }

  /** Add a command that armed (its batch had `armedCaret`). False when it cannot arm here. */
  add(command: EditorCommand): boolean {
    const caret = this.host.caret();
    if (!caret) return false;
    const before = this.current()?.commands ?? [];
    const armed = this.#replay(caret, [...before, command]);
    if (!armed) return false;
    const read = this.host.formatAt(this.#layoutOf(armed.pkg, caret.part), caret);
    this.#armed = { ...armed, commands: [...before, absoluteOf(command, read)] };
    this.version += 1;
    return true;
  }

  /**
   * Typing took the arm in the history entry whose mark (its `selectionBefore`, one object per
   * entry) is `mark`: remember the commands, for an undo of that entry.
   */
  typed(mark: object, commands: readonly EditorCommand[]): void {
    this.#typed.set(mark, commands);
  }

  /**
   * An undo removed the entry with `mark`. When that entry was typing that took an arm, arm its
   * commands again at the caret. False when nothing arms (no such entry, or the caret is not at
   * a word edge any more).
   */
  rearm(mark: object | undefined): boolean {
    const commands = mark ? this.#typed.get(mark) : undefined;
    const caret = this.host.caret();
    if (!caret || !commands) return false;
    const armed = this.#replay(caret, commands);
    if (!armed) return false;
    this.#armed = armed;
    this.version += 1;
    return true;
  }

  /** The layout over `pkg` on the slide `part`: one preview is kept, by package and slide. */
  #layoutOf(pkg: PptxPackage, part: string): SlideLayout | undefined {
    if (this.#preview?.pkg !== pkg || this.#preview.part !== part)
      this.#preview = { pkg, part, layout: this.host.layoutOf(pkg) };
    return this.#preview.layout;
  }

  /** Compile the commands one after another over the store package, each over the last. */
  #replay(caret: ArmedCaret, commands: readonly EditorCommand[]): Armed | null {
    this.#replaying = true;
    try {
      return this.#compileAll(caret, commands);
    } finally {
      this.#replaying = false;
    }
  }

  #compileAll(caret: ArmedCaret, commands: readonly EditorCommand[]): Armed | null {
    const base = this.host.package();
    if (!base) return null;
    let pkg = base;
    const ops: PptxOp[] = [];
    for (const command of commands) {
      let batch: OpBatch;
      try {
        batch = this.host.compileOn(command, pkg, this.#layoutOf(pkg, caret.part), caret);
      } catch (e) {
        if (e instanceof CommandRefusal) return null;
        throw e;
      }
      // The caret is inside a word now (a peer typed there): the arm ends.
      if (!batch.armedCaret) return null;
      ops.push(...batch.ops);
      pkg = batch.pkg;
    }
    return { ...caret, commands, base, ops, pkg };
  }
}
