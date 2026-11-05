// The DOM-typed editor contract: `attach` and `renderThumbnail` take an `HTMLElement`. Only the
// editor lane and the adapters import it, through `index.ts`. The DOM-free contracts are in
// `neutral.ts`, which the automation lane imports.

import type { EditorExecOptions, HistoryGroup } from './history-group.ts';
import type {
  CanResult,
  ClipboardPayload,
  CommandResult,
  EditorCommand,
  EditorMode,
  EditorSnapshot,
  SlideSelectMode,
} from './neutral.ts';

/** The editor contract adapters render. @public */
export interface Editor {
  snapshot(): EditorSnapshot;
  subscribe(listener: () => void): () => void;
  /** Dry run: whether `exec` would apply. A closed or foreign `historyGroup` refuses. */
  can(command: EditorCommand, options?: EditorExecOptions): CanResult;
  /**
   * Run a command. Pass `options.historyGroup` to make several commands one undo step.
   * See {@link HistoryGroupOptions}.
   */
  exec(command: EditorCommand, options?: EditorExecOptions): CommandResult;
  /**
   * Open a history group: every command run with it (`exec(command, { historyGroup })`) until
   * `end()` is one undo step. Opening one creates no undo entry. Undo, redo, and a command with
   * another handle or none split the step; a selection change does not. `load` and `destroy`
   * close every handle. End it in a `finally` block, so a refused or throwing command does not
   * leave it open.
   *
   * @example
   * ```ts
   * const group = editor.beginHistoryGroup();
   * try {
   *   editor.exec({ type: 'nudge', dx: 91440, dy: 0 }, { historyGroup: group });
   *   editor.exec({ type: 'rotateObjects', degrees: 90 }, { historyGroup: group });
   * } finally {
   *   group.end();
   * }
   * ```
   */
  beginHistoryGroup(): HistoryGroup;
  load(bytes: Uint8Array): Promise<CommandResult>;
  /**
   * The presentation as `.pptx` bytes, or `null` when no presentation is open. The bytes own an
   * `ArrayBuffer`, so `new Blob([bytes])` takes them as they are.
   */
  save(): Uint8Array<ArrayBuffer> | null;
  setZoom(zoom: number | 'fit'): void;
  setMode(mode: EditorMode): void;
  goToSlide(index: number): void;
  /**
   * Click slide `index` in the rail: `replace` selects only it (as `goToSlide`), `toggle`
   * (Ctrl/Cmd+click) adds or removes it, and `range` (Shift+click) selects from the anchor to
   * it. The clicked slide becomes current, unless a toggle removed it.
   */
  selectSlide(index: number, mode?: SlideSelectMode): void;
  attach(element: HTMLElement): void;
  detach(): void;
  /** The selected shapes as a clipboard payload, or null when nothing is copyable. */
  copy(): ClipboardPayload | null;
  /** Select every shape in the current scope (the slide or the entered group). */
  selectAll(): void;
  /**
   * Give the keyboard back to the slide, as DOCX Editor's `focus()`: the text caret for a text
   * selection, else the slide surface. Host pickers call it after a pick.
   */
  focus(): void;
  /** Paint a slide thumbnail into a container (the rail). */
  renderThumbnail(index: number, container: HTMLElement): void;
  destroy(): void;
}
