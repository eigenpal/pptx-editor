// The core types that the adapter's props and hooks use come from the adapter package, as DOCX
// Editor's adapters re-export `Editor`, `EditorCommand`, and `EditorSnapshot`: a host types its
// handlers with no import from `@pptx-editor.dev/core`. This file imports from the adapter
// index only, so `bun run typecheck` fails when a re-export goes missing.

import { describe, expect, test } from 'bun:test';
import {
  PptxEditor,
  type ChromeSlotId,
  type CommandResult,
  type EditorCommand,
  type EditorMode,
  type EditorModule,
  type EditorSnapshot,
  type FontPickerRow,
  type PptxEditorInstance,
} from '../index.ts';

describe('adapter type re-exports', () => {
  test('onReady and onChange take handlers typed with adapter imports only', () => {
    const seen: PptxEditorInstance[] = [];
    const mode: EditorMode = 'view';
    const modules: readonly EditorModule[] = [];
    const element = (
      <PptxEditor
        mode={mode}
        modules={modules}
        onReady={(e: PptxEditorInstance) => seen.push(e)}
        onChange={(e: PptxEditorInstance) => {
          const snapshot: EditorSnapshot = e.snapshot();
          void snapshot;
        }}
      />
    );
    const slot: ChromeSlotId = 'text.bold';
    const command: EditorCommand = { type: 'undo' };
    const run = (e: PptxEditorInstance): CommandResult => e.exec(command);
    const row: Pick<FontPickerRow, 'family' | 'value'> = { family: 'Arial', value: 'Arial' };
    expect(element.type).toBe(PptxEditor);
    expect([slot, typeof run, row.value]).toEqual(['text.bold', 'function', 'Arial']);
  });
});
