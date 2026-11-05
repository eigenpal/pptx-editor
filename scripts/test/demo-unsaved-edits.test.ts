// The demos' unsaved-edits guard (`examples/shared/demo-files.ts`): an edit makes the deck
// dirty, an undo back to the opened deck makes it clean again, and an export marks it saved.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../fixtures/builder.ts';
import {
  initialTheme,
  rememberTheme,
  trackUnsavedEdits,
} from '../../examples/shared/demo-files.ts';

async function open() {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({ slides: [{ shapes: '', notes: 'one' }, { shapes: '' }] });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

describe('trackUnsavedEdits', () => {
  test('undo back to the opened deck is clean', async () => {
    const editor = await open();
    const unsaved = trackUnsavedEdits(editor);
    expect(unsaved.dirty()).toBe(false);
    expect(editor.exec({ type: 'setNotes', index: 0, text: 'two' }).ok).toBe(true);
    expect(unsaved.dirty()).toBe(true);
    editor.exec({ type: 'undo' });
    expect(unsaved.dirty()).toBe(false);
    editor.exec({ type: 'redo' });
    expect(unsaved.dirty()).toBe(true);
    unsaved.dispose();
  });

  test('an undo after a structural edit (new ids) is clean too', async () => {
    const editor = await open();
    const unsaved = trackUnsavedEdits(editor);
    expect(editor.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    expect(unsaved.dirty()).toBe(true);
    editor.exec({ type: 'undo' });
    expect(unsaved.dirty()).toBe(false);
    unsaved.dispose();
  });

  test('an export marks the edits saved; undo past it is dirty', async () => {
    const editor = await open();
    const unsaved = trackUnsavedEdits(editor);
    editor.exec({ type: 'setNotes', index: 0, text: 'two' });
    unsaved.markSaved();
    expect(unsaved.dirty()).toBe(false);
    editor.exec({ type: 'undo' });
    expect(unsaved.dirty()).toBe(true);
    unsaved.dispose();
  });
});

describe('initialTheme', () => {
  const g = globalThis as { matchMedia?: unknown; localStorage?: unknown };

  test('the URL, then the theme picked before, then the system color scheme', () => {
    const saved = { matchMedia: g.matchMedia, localStorage: g.localStorage };
    const store = new Map<string, string>();
    let dark = true;
    g.matchMedia = (q: string) => ({ matches: dark && q === '(prefers-color-scheme: dark)' });
    g.localStorage = {
      getItem: (k: string) => store.get(k) ?? null,
      setItem: (k: string, v: string) => void store.set(k, v),
    };
    try {
      // Nothing stored: the system decides.
      expect(initialTheme('')).toBe('dark');
      dark = false;
      expect(initialTheme('')).toBe('light');
      expect(initialTheme('?theme=dark')).toBe('dark');
      rememberTheme('dark');
      expect(initialTheme('')).toBe('dark');
      expect(initialTheme('?theme=light')).toBe('light');
      // Storage that throws (a private window) falls back to the system.
      g.localStorage = {
        getItem: () => {
          throw new Error('off');
        },
      };
      expect(initialTheme('')).toBe('light');
    } finally {
      g.matchMedia = saved.matchMedia;
      g.localStorage = saved.localStorage;
    }
  });
});
