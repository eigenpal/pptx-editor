import { describe, expect, test } from 'bun:test';
import type { RegisteredFace } from '../../layout/index.ts';
import { registerFontFaces } from '../font-faces.ts';

const face = { id: 'pptx-face-1', bytes: new Uint8Array([0, 1, 0, 0]) } as RegisteredFace;

describe('registerFontFaces', () => {
  test('registers nothing, and does not throw, in a DOM without the Font Loading API', async () => {
    // The engine tests run with no DOM, so `FontFace` is undefined here, as in a test DOM.
    expect((globalThis as { FontFace?: unknown }).FontFace).toBeUndefined();
    const added: unknown[] = [];
    const doc = { fonts: { add: (f: unknown) => added.push(f) } } as unknown as Document;
    const handle = registerFontFaces(doc, [face]);
    await handle.ready;
    handle.dispose();
    expect(added).toEqual([]);
  });

  test('adds each face to the document font set when the API exists', async () => {
    const created: { family: string; loaded: boolean }[] = [];
    class FakeFontFace {
      loaded = false;
      constructor(readonly family: string) {
        created.push(this);
      }
      load(): Promise<this> {
        this.loaded = true;
        return Promise.resolve(this);
      }
    }
    const fonts = new Set<unknown>();
    const doc = {
      fonts: { add: (f: unknown) => fonts.add(f), delete: (f: unknown) => fonts.delete(f) },
    };
    const host = globalThis as { FontFace?: unknown };
    host.FontFace = FakeFontFace;
    try {
      const handle = registerFontFaces(doc as unknown as Document, [face]);
      await handle.ready;
      expect(created.map((f) => [f.family, f.loaded])).toEqual([['pptx-face-1', true]]);
      expect(fonts.size).toBe(1);
      handle.dispose();
      expect(fonts.size).toBe(0);
    } finally {
      delete host.FontFace;
    }
  });
});
