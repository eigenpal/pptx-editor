// Holds the packaged substitute loads of chosen families until a test releases them, and
// records the families each load asked for. Import it before the editor: it mocks
// `@docx-editor.dev/fonts`, so each test file that uses it needs its own process, as
// `bun run test` gives it. The session caches each family for the page, so a family that one
// test loaded is not held in a later test of the same file.

import { mock } from 'bun:test';
import * as realFonts from '@docx-editor.dev/fonts';

const original = realFonts.loadDefaultFonts;
const copy = { ...realFonts };

/** The families the substitute loads asked for, in order. */
export const requested: string[] = [];
const held = new Map<string, { promise: Promise<void>; resolve: () => void }>();

/** Hold the loads of `family` until `release(family)`. */
export function hold(family: string): void {
  let resolve!: () => void;
  const promise = new Promise<void>((r) => (resolve = r));
  held.set(family, { promise, resolve });
}

/** Let the loads of `family` finish. */
export function release(family: string): void {
  held.get(family)?.resolve();
}

/** Release every held family and forget what was asked for. */
export function resetGate(): void {
  for (const family of held.keys()) release(family);
  held.clear();
  requested.length = 0;
}

mock.module('@docx-editor.dev/fonts', () => ({
  ...copy,
  loadDefaultFonts: (o: Parameters<typeof realFonts.loadDefaultFonts>[0] = {}) => {
    const families = [...(o.families ?? [])];
    requested.push(...families);
    const gates = families.map((f) => held.get(f)?.promise ?? Promise.resolve());
    return Promise.all(gates).then(() => original(o));
  },
}));

/** Wait until `done` answers true, or fail after `ms`. */
export async function waitFor(done: () => boolean, ms = 4000): Promise<void> {
  const end = Date.now() + ms;
  while (!done()) {
    if (Date.now() > end) throw new Error('timed out');
    await new Promise((r) => setTimeout(r, 5));
  }
}

const IN = 914_400;

/**
 * A deck whose theme font is Calibri, with one slide for each entry of `slides`: a text box that
 * names that family, or none (`undefined`) for theme-font text.
 */
export async function deckNaming(slides: readonly (string | undefined)[]): Promise<Uint8Array> {
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  return buildPptx({
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    slides: slides.map((typeface, i) => ({
      shapes: shape({
        id: 2,
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: IN,
        text: para(run(`Slide ${i + 1}`, '', typeface ? `<a:latin typeface="${typeface}"/>` : '')),
      }),
    })),
  });
}
