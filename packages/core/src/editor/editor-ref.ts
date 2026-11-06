// The members of the adapters' imperative `PptxEditorRef`, after DOCX Editor's
// `createDocxEditorRefApi` (docx-editor/packages/vue/src/components/DocxEditor/useDocxEditorRefApi.ts)
// and `useDocxEditorRefApi` (docx-editor/packages/react/src/components/DocxEditor/hooks/).
// React's `useImperativeHandle` and Vue's `expose` both build the handle here, so the two refs
// cannot drift.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type {
  CommandResult,
  EditorCommand,
  EditorExecOptions,
  EditorSnapshot,
  LoadOptions,
} from '../contracts/index.ts';
import { EMPTY_EDITOR_SNAPSHOT } from './adapter-support.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';

/**
 * The imperative handle both adapters expose, member for member. Every member forwards to the
 * live instance and is safe to call before it mounts or after it unmounts. @internal
 */
export interface EditorRefApi {
  load(bytes: Uint8Array, options?: LoadOptions): Promise<CommandResult>;
  save(): Uint8Array<ArrayBuffer> | null;
  getEditor(): PptxEditorInstance | null;
  focus(): void;
  exec(command: EditorCommand, options?: EditorExecOptions): CommandResult;
  snapshot(): EditorSnapshot;
}

/** The refusal a ref command gets while no editor is mounted. */
const NO_EDITOR: CommandResult = Object.freeze({ ok: false, reason: 'disabled.noEditor' });

/**
 * Build the ref members over `getEditor`, which reads the live instance at each call. Before
 * mount, `load` and `exec` refuse with `disabled.noEditor`, `save` and `getEditor` return null,
 * `focus` does nothing, and `snapshot` returns the empty snapshot the hooks read. @internal
 */
export function createEditorRefApi(getEditor: () => PptxEditorInstance | null): EditorRefApi {
  return {
    load: (bytes, options) => getEditor()?.load(bytes, options) ?? Promise.resolve(NO_EDITOR),
    save: () => getEditor()?.save() ?? null,
    getEditor: () => getEditor(),
    focus: () => {
      getEditor()?.focus();
    },
    exec: (command, options) => getEditor()?.exec(command, options) ?? NO_EDITOR,
    snapshot: () => getEditor()?.snapshot() ?? EMPTY_EDITOR_SNAPSHOT,
  };
}
