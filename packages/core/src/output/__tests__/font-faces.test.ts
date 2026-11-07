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

  test("takes the constructor from the document's own window", async () => {
    const created: string[] = [];
    class FrameFontFace {
      constructor(readonly family: string) {
        created.push(family);
      }
      load(): Promise<this> {
        return Promise.resolve(this);
      }
    }
    const fonts = new Set<unknown>();
    const doc = {
      defaultView: { FontFace: FrameFontFace },
      fonts: { add: (f: unknown) => fonts.add(f), delete: (f: unknown) => fonts.delete(f) },
    };
    const handle = registerFontFaces(doc as unknown as Document, [face]);
    await handle.ready;
    expect(created).toEqual(['pptx-face-1']);
    expect([...fonts][0]).toBeInstanceOf(FrameFontFace);
    handle.dispose();
  });

  test('names each face as the `family` option says', async () => {
    const created: string[] = [];
    class FrameFontFace {
      constructor(readonly family: string) {
        created.push(family);
      }
      load(): Promise<this> {
        return Promise.resolve(this);
      }
    }
    const doc = {
      defaultView: { FontFace: FrameFontFace },
      fonts: { add: () => undefined, delete: () => undefined },
    };
    const other = { ...face, id: 'pptx-face-2' } as RegisteredFace;
    const handle = registerFontFaces(doc as unknown as Document, [face, other], {
      family: (f) => `host-${f.id}`,
    });
    await handle.ready;
    expect(created).toEqual(['host-pptx-face-1', 'host-pptx-face-2']);
  });

  test('ready settles after the wait when a face never loads', async () => {
    class StuckFontFace {
      load(): Promise<this> {
        return new Promise(() => {});
      }
    }
    const doc = { fonts: { add: () => undefined, delete: () => undefined } };
    const host = globalThis as { FontFace?: unknown };
    host.FontFace = StuckFontFace;
    try {
      const t0 = performance.now();
      const handle = registerFontFaces(doc as unknown as Document, [face], { waitMs: 40 });
      await handle.ready;
      expect(performance.now() - t0).toBeGreaterThanOrEqual(35);
      handle.dispose();
    } finally {
      delete host.FontFace;
    }
  });
});

test('registration keeps late faces until disposal and never adds them again', async () => {
  let complete!: () => void;
  class SlowFontFace {
    load(): Promise<void> {
      return new Promise((resolve) => {
        complete = resolve;
      });
    }
  }
  const fonts = new Set<unknown>();
  const doc = {
    defaultView: { FontFace: SlowFontFace },
    fonts: { add: (f: unknown) => fonts.add(f), delete: (f: unknown) => fonts.delete(f) },
  };
  const handle = registerFontFaces(doc as unknown as Document, [face]);
  let ready = false;
  void handle.ready.then(() => {
    ready = true;
  });
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(ready).toBe(false);
  expect(fonts.size).toBe(1);
  handle.dispose();
  await handle.ready;
  complete();
  await Promise.resolve();
  expect(fonts.size).toBe(0);
});

test('dispose releases a pending browser registration wait', async () => {
  class SlowFontFace {
    load(): Promise<void> {
      return new Promise(() => {});
    }
  }
  const fonts = new Set<unknown>();
  const doc = {
    defaultView: { FontFace: SlowFontFace },
    fonts: { add: (f: unknown) => fonts.add(f), delete: (f: unknown) => fonts.delete(f) },
  };
  const handle = registerFontFaces(doc as unknown as Document, [face]);
  handle.dispose();
  await handle.ready;
  expect(fonts.size).toBe(0);
}, 500);

test('a constructor failure removes faces already registered', () => {
  let count = 0;
  class BrokenFontFace {
    constructor() {
      if (++count === 2) throw new Error('font failure');
    }
    load(): Promise<void> {
      return Promise.resolve();
    }
  }
  const fonts = new Set<unknown>();
  const doc = {
    defaultView: { FontFace: BrokenFontFace },
    fonts: { add: (f: unknown) => fonts.add(f), delete: (f: unknown) => fonts.delete(f) },
  };
  expect(() => registerFontFaces(doc as unknown as Document, [face, face])).toThrow('font failure');
  expect(fonts.size).toBe(0);
});
