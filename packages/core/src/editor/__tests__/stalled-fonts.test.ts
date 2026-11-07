// A font origin that never answers must not hold the open: the deck paints with the faces it
// has after `FONT_WAIT_MS`, and the host's faces install when they arrive.

import { expect, test } from 'bun:test';
import { createPptxEditor, defineFontResolver, type FontConfigurationFragment } from '../index.ts';
import { FONT_WAIT_MS } from '../host-fonts.ts';
import { deferred, face, settle, missingFontsDeck } from './font-test-helpers.ts';

test(
  'a stalled font origin opens the deck with stand-in faces, then installs the late ones',
  async () => {
    const slow = deferred<FontConfigurationFragment | undefined>();
    const editor = createPptxEditor({ fonts: defineFontResolver(() => slow.promise) });
    const start = performance.now();
    expect((await editor.load(missingFontsDeck())).ok).toBe(true);
    const waited = performance.now() - start;
    expect(waited).toBeGreaterThanOrEqual(FONT_WAIT_MS - 50);
    expect(waited).toBeLessThan(FONT_WAIT_MS + 1000);
    expect(editor.snapshot().status).toBe('ready');
    // The fonts are still on the way: no substitution notice yet.
    expect(editor.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: true });
    expect(editor.snapshot().substitutedFonts).toEqual([]);
    const generation = editor.session!.fontGeneration;
    slow.resolve({ sources: [face('Aptos')] });
    await settle(10);
    expect(editor.session!.fontGeneration).toBeGreaterThan(generation);
    expect(editor.fontMeasurement().resolving).toBe(false);
    expect(editor.snapshot().substitutedFonts.map((f) => f.family)).not.toContain('Aptos');
    editor.destroy();
  },
  FONT_WAIT_MS + 5000
);

test(
  'a late answer for a replaced deck never installs',
  async () => {
    const slow = deferred<FontConfigurationFragment | undefined>();
    let calls = 0;
    const editor = createPptxEditor({
      fonts: defineFontResolver(() => (++calls === 1 ? slow.promise : undefined)),
    });
    await editor.load(missingFontsDeck());
    await editor.load(missingFontsDeck());
    slow.resolve({ sources: [face('Aptos')] });
    await settle(10);
    expect(editor.snapshot().substitutedFonts.map((f) => f.family)).toContain('Aptos');
    editor.destroy();
  },
  FONT_WAIT_MS + 5000
);

test('fonts that never arrive stop holding the notice after the late wait', async () => {
  const { createHostFonts } = await import('../host-fonts.ts');
  const { openPresentation } = await import('../document-session.ts');
  const opened = await openPresentation(missingFontsDeck());
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  let changes = 0;
  const fonts = createHostFonts(
    { fonts: defineFontResolver(() => new Promise<FontConfigurationFragment>(() => {})) },
    {
      session: () => session,
      installed: async () => {},
      changed: () => void (changes += 1),
      waitMs: 10,
      lateWaitMs: 50,
    }
  );
  await fonts.open(session, () => false);
  expect(fonts.resolving()).toBe(true);
  await new Promise((r) => setTimeout(r, 80));
  expect(fonts.resolving()).toBe(false);
  expect(changes).toBe(1);
  fonts.dispose();
  session.dispose();
});

test('an open that does not mount gives the shown deck its font state back', async () => {
  const { createHostFonts } = await import('../host-fonts.ts');
  const { openPresentation } = await import('../document-session.ts');
  const open = async () => {
    const opened = await openPresentation(missingFontsDeck());
    if (!opened.ok) throw new Error('open failed');
    return opened.session;
  };
  const a = await open();
  const b = await open();
  let calls = 0;
  const fonts = createHostFonts(
    {
      fonts: defineFontResolver((request) => {
        calls += 1;
        if (calls === 1) return { sources: [face('Aptos')] };
        // The open that fails: its answer never arrives.
        if (calls === 2) return new Promise<FontConfigurationFragment>(() => {});
        return request.families.includes('Calibri') ? { sources: [face('Calibri')] } : undefined;
      }),
    },
    { session: () => a, installed: async () => {}, changed: () => {}, waitMs: 10 }
  );
  const families = () => new Set(a.faces.map((f) => f.family));
  await fonts.open(a, () => false);
  expect(families().has('Aptos')).toBe(true);
  await fonts.open(b, () => false);
  fonts.abandon(b);
  await fonts.request(['Calibri']);
  expect(families().has('Calibri')).toBe(true);
  expect(families().has('Aptos')).toBe(true);
  fonts.dispose();
  a.dispose();
  b.dispose();
});

test('an open that does not mount keeps the shown deck waiting for its shaping', async () => {
  const { createHostFonts } = await import('../host-fonts.ts');
  const { openPresentation } = await import('../document-session.ts');
  const open = async () => {
    const opened = await openPresentation(missingFontsDeck());
    if (!opened.ok) throw new Error('open failed');
    return opened.session;
  };
  const a = await open();
  const b = await open();
  const fonts = createHostFonts(
    {},
    { session: () => a, installed: async () => {}, changed: () => {}, waitMs: 10 }
  );
  await fonts.open(a, () => false);
  // The shown deck's text engine is still on the way.
  const shaping = deferred<boolean>();
  a.shaping = shaping.promise;
  fonts.awaitShaping(a);
  expect(fonts.resolving()).toBe(true);
  await fonts.open(b, () => false);
  fonts.abandon(b);
  expect(fonts.resolving()).toBe(true);
  shaping.resolve(true);
  await settle(5);
  expect(fonts.resolving()).toBe(false);
  fonts.dispose();
  a.dispose();
  b.dispose();
});
