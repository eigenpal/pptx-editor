import { afterAll, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { FONT_ASSET_ROOT } from '@docx-editor.dev/fonts';
import { createPackagedFileFetch } from '../packaged-file-fetch.ts';
import { openPresentation } from '../document-session.ts';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';

const base = mkdtempSync(join(tmpdir(), 'pptx-packaged-fetch-'));
const root = join(base, 'assets');
mkdirSync(root);
writeFileSync(join(root, 'face.ttf'), new Uint8Array([1, 2, 3, 4]));
writeFileSync(join(root, 'big.ttf'), new Uint8Array(64));
writeFileSync(join(base, 'secret.txt'), 'outside');
symlinkSync(join(base, 'secret.txt'), join(root, 'link.ttf'));
afterAll(() => rmSync(base, { recursive: true, force: true }));

const rootUrl = pathToFileURL(`${root}/`);
const read = createPackagedFileFetch({ trustedRoot: rootUrl, maxBytes: 16 })!;
const at = (name: string): string => new URL(name, rootUrl).href;

describe('createPackagedFileFetch', () => {
  test('reads a file inside the trusted folder', async () => {
    const response = await read(at('face.ttf'));
    expect(response.status).toBe(200);
    expect([...new Uint8Array(await response.arrayBuffer())]).toEqual([1, 2, 3, 4]);
  });

  test('refuses files outside the folder, through `..` or a symbolic link', async () => {
    expect((await read(`${rootUrl.href}..%2Fsecret.txt`)).status).toBe(404);
    expect((await read(pathToFileURL(join(base, 'secret.txt')).href)).status).toBe(404);
    expect((await read(at('link.ttf'))).status).toBe(404);
    expect((await read(at('missing.ttf'))).status).toBe(404);
  });

  test('refuses a file over maxBytes before it reads it', async () => {
    expect((await read(at('big.ttf'))).status).toBe(413);
  });

  test('leaves a non-file folder to the global fetch', () => {
    expect(
      createPackagedFileFetch({ trustedRoot: new URL('https://example.test/'), maxBytes: 1 })
    ).toBeUndefined();
  });

  test('is undefined off Node, where the global fetch serves the faces', () => {
    const runtime = globalThis.process as { getBuiltinModule?: unknown };
    const saved = runtime.getBuiltinModule;
    runtime.getBuiltinModule = undefined;
    try {
      expect(createPackagedFileFetch({ trustedRoot: rootUrl, maxBytes: 1 })).toBeUndefined();
    } finally {
      runtime.getBuiltinModule = saved;
    }
  });
});

describe('packaged substitutes', () => {
  test('load from disk with a fetch that refuses file: URLs, as Node does', async () => {
    expect(FONT_ASSET_ROOT.protocol).toBe('file:');
    const saved = globalThis.fetch;
    // Node's fetch rejects every file: URL ("fetch failed"). Bun's reads them, which hid the
    // fault from every test that ran under Bun.
    globalThis.fetch = (async (input: string | URL | Request) => {
      const href = typeof input === 'string' ? input : 'href' in input ? input.href : input.url;
      if (href.startsWith('file:')) throw new TypeError('fetch failed');
      return saved(input);
    }) as typeof fetch;
    try {
      const opened = await openPresentation(
        buildPptx({ slides: [{ layout: 'blank', shapes: '' }] })
      );
      if (!opened.ok) throw new Error('deck');
      const report = opened.session.fontReport;
      expect(report.substituteFailures).toEqual([]);
      expect(report.faces.some((f) => f.origin === 'substitute')).toBe(true);
      opened.session.dispose();
    } finally {
      globalThis.fetch = saved;
    }
  });
});
