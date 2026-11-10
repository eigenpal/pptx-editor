/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { defineFontResolver, type FontConfigurationFragment } from '@pptx-editor.dev/core/layout';
import { fontConfiguration } from '../../../../scripts/fixtures/font-configuration.ts';
import { buildPptx, shape, para, run } from '../../../../scripts/fixtures/builder.ts';
import { exportFonts } from '../font-provisioning.ts';
import { openDocumentForExport } from '../index.ts';
import { sessionState } from '../open-session.ts';

const bytes = new Uint8Array(
  readFileSync(new URL('../../../../fixtures/fonts/NotoSansArabic-Regular.ttf', import.meta.url))
);
const config = fontConfiguration([
  { family: 'Noto Sans Arabic', bytes, bold: false, italic: false, origin: 'fixture' },
]);
const deck = buildPptx({
  slides: [
    {
      shapes: shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 5000000,
        cy: 1000000,
        text: para(run('مرحبا', '', '<a:cs typeface="Alias"/>'), '<a:pPr rtl="1"/>'),
      }),
    },
  ],
});

test('ordered configurations and resolvers receive first-wins face coverage', async () => {
  let called = false;
  const resolver = defineFontResolver((request) => {
    called = true;
    expect(request.families).toContain('Amiri');
    expect(request.resolvedFaces?.some((face) => face.family === 'Noto Sans Arabic')).toBe(true);
    return config;
  });
  const fonts = await exportFonts(deck, [config, resolver]);
  expect(called).toBe(true);
  expect(fonts.faces).toHaveLength(1);
});

test('configuration promises and deferred faces finish before session open', async () => {
  const opened = await openDocumentForExport(deck, {
    fonts: Promise.resolve({ sources: [], deferred: Promise.resolve(config) }),
  });
  expect(opened.ok).toBe(true);
  if (!opened.ok) return;
  try {
    expect(opened.session.fontResolution.faces).toContainEqual({
      family: 'Noto Sans Arabic',
      origin: 'host',
    });
  } finally {
    opened.session.dispose();
  }
});

test('empty deferred results retain immediate faces', async () => {
  const fonts = await exportFonts(deck, { ...config, deferred: Promise.resolve(undefined) });
  expect(fonts.faces).toHaveLength(1);
});

test('substitution configurations reach the shared session measurer', async () => {
  const target = config.sources![0]!.request;
  const opened = await openDocumentForExport(deck, {
    substitutes: false,
    fonts: {
      ...config,
      substitutions: [{ from: { ...target, family: 'Alias' }, to: target }],
    },
  });
  if (!opened.ok) throw new Error('fixture');
  try {
    const aliases = sessionState(opened.session)!.document.hostFaces.substitutions;
    expect(aliases).toHaveLength(1);
    expect(aliases[0]!.target).toBe('Noto Sans Arabic');
    expect(aliases[0]!.face.family).toBe('Alias');
  } finally {
    opened.session.dispose();
  }
});

test('invalid hashes and resource limits fail admission', async () => {
  const face = config.sources![0]!;
  await expect(
    exportFonts(deck, { sources: [{ ...face, hash: 'sha256:' + '0'.repeat(64) }] })
  ).rejects.toThrow('admission');
  await expect(exportFonts(deck, { ...config, maxFontBytes: 1 })).rejects.toThrow();
});

test('origin and deferred failures cannot silently remove fonts', async () => {
  await expect(
    exportFonts(
      deck,
      defineFontResolver(() => {
        throw new Error('unavailable');
      })
    )
  ).rejects.toThrow('resolution');
  await expect(
    exportFonts(deck, { ...config, deferred: Promise.reject(new Error('unavailable')) })
  ).rejects.toThrow('resolution');
});

for (const deferred of [false, true]) {
  test(`cancellation stops an uncooperative ${deferred ? 'deferred source' : 'resolver'}`, async () => {
    const controller = new AbortController();
    let began!: () => void;
    const started = new Promise<void>((resolve) => {
      began = resolve;
    });
    const pending = new Promise<FontConfigurationFragment>(() => {});
    const fonts = defineFontResolver(() => {
      began();
      return deferred ? { ...config, deferred: pending } : pending;
    });
    const opening = openDocumentForExport(deck, { fonts, signal: controller.signal });
    await started;
    controller.abort();
    expect(await opening).toEqual({ ok: false, reason: 'aborted' });
  });
}
