// The Vue twin of packages/react/src/__tests__/adapter-types.test.tsx: the core types that the
// adapter's props and composables use come from the adapter package, as DOCX Editor's Vue
// adapter re-exports `EditorCommand` and `EditorSnapshot`. This file imports from the adapter
// index only, so `bun run typecheck` fails when a re-export goes missing.

import { describe, expect, test } from 'bun:test';
import { h } from 'vue';
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
} from '../src/index.ts';

describe('adapter type re-exports (Vue)', () => {
  test('ready and change take handlers typed with adapter imports only', () => {
    const seen: PptxEditorInstance[] = [];
    const mode: EditorMode = 'view';
    const modules: readonly EditorModule[] = [];
    const vnode = h(PptxEditor, {
      mode,
      modules,
      onReady: (e: PptxEditorInstance) => seen.push(e),
      onChange: (e: PptxEditorInstance) => {
        const snapshot: EditorSnapshot = e.snapshot();
        void snapshot;
      },
    });
    const slot: ChromeSlotId = 'text.bold';
    const command: EditorCommand = { type: 'undo' };
    const run = (e: PptxEditorInstance): CommandResult => e.exec(command);
    const row: Pick<FontPickerRow, 'family' | 'value'> = { family: 'Arial', value: 'Arial' };
    expect(vnode.type).toBe(PptxEditor);
    expect([slot, typeof run, row.value]).toEqual(['text.bold', 'function', 'Arial']);
  });
});
