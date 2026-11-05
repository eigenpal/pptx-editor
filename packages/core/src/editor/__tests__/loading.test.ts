import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../pptx-editor.ts';
import { loadingSheetSize, showsLoading } from '../loading.ts';

// The loading screen's state, which both adapters read: shown with no deck and while a deck
// opens, cleared when the deck is ready or the file is refused.

const deck = () => buildPptx({ slides: [{ shapes: '' }] });

describe('loading state', () => {
  test('load publishes `loading` before it parses, and `ready` after', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect(showsLoading(editor.snapshot())).toBe(true);
    const seen: string[] = [];
    editor.subscribe(() => seen.push(editor.snapshot().status));
    const opening = editor.load(deck());
    // Synchronous: the adapters render the screen before the parse blocks the thread.
    expect(editor.snapshot().status).toBe('loading');
    expect((await opening).ok).toBe(true);
    expect(seen[0]).toBe('loading');
    expect(showsLoading(editor.snapshot())).toBe(false);
  });

  test('a refused file clears the loading state', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(new Uint8Array([1, 2, 3]));
    expect(showsLoading(editor.snapshot())).toBe(false);
  });

  test('a later load supersedes one that is still waiting for its frame', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const first = editor.load(deck());
    const second = editor.load(deck());
    expect(await first).toEqual({ ok: false, reason: 'errors.superseded' });
    expect((await second).ok).toBe(true);
  });

  test('the sheet has the slide size at the zoom, and 16:9 before a deck', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const before = loadingSheetSize(editor.snapshot());
    expect(before.width / before.height).toBeCloseTo(16 / 9, 3);
    await editor.load(buildPptx({ slides: [{ shapes: '' }], width: 9_144_000, height: 6_858_000 }));
    editor.setZoom(0.5);
    expect(loadingSheetSize(editor.snapshot())).toEqual({ width: 480, height: 360 });
  });
});
