// `onEditorEvent` derives typed events from the editor's one subscription: each event fires when
// its part of the snapshot moves, and only then.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, onEditorEvent, type PptxEditorEvents } from '../index.ts';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';

const deck = () => buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] });

function record(editor: ReturnType<typeof createPptxEditor>) {
  const log: string[] = [];
  const names: (keyof PptxEditorEvents)[] = ['change', 'selectionChange', 'rejection'];
  const offs = names.map((name) => onEditorEvent(editor, name, () => void log.push(name)));
  offs.push(onEditorEvent(editor, 'openFailure', (f) => void log.push(`openFailure:${f.reason}`)));
  return { log, off: () => offs.forEach((off) => off()) };
}

describe('onEditorEvent', () => {
  test('each event fires for its own part of the state', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true });
    const { log, off } = record(editor);
    await editor.load(deck());
    expect(log).toContain('change');
    log.length = 0;

    editor.selectSlide(1);
    expect(log).toEqual(['selectionChange']);
    log.length = 0;

    editor.setZoom(2);
    expect(log).toEqual([]);

    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(log).toContain('change');
    log.length = 0;

    expect(editor.exec({ type: 'redo' }).ok).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(false);
    expect(log).toEqual(['rejection', 'rejection']);
    log.length = 0;

    await editor.load(new Uint8Array([1, 2, 3]));
    expect(log).toContain('openFailure:damaged');

    off();
    log.length = 0;
    editor.selectSlide(0);
    expect(log).toEqual([]);
    editor.destroy();
  });

  test('the handler receives the rejection and the snapshot', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true });
    await editor.load(deck());
    let seen: unknown[] = [];
    onEditorEvent(editor, 'rejection', (...args) => void (seen = args));
    editor.exec({ type: 'undo' });
    expect((seen[0] as { reason: string }).reason).toBe('disabled.nothingToUndo');
    expect(seen[1]).toBe(editor.snapshot());
    editor.destroy();
  });
});
