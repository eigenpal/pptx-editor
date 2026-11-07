// A failed open always ends in a refusal a person can read and report, with DOCX Editor's
// codes: an open that throws is `failed` and never leaves the loading screen up, a limit has a
// code of its own, and the report a person copies holds codes and versions, never the file's
// content.

import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, openFailureReport } from '../index.ts';

/** A slide whose shape tree nests `depth` groups. */
function deepDeck(depth: number): Uint8Array {
  let inner = shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100, text: para(run('Private text')) });
  for (let i = 0; i < depth; i += 1)
    inner =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${100 + i}" name="G"/><p:cNvGrpSpPr/><p:nvPr/>` +
      `</p:nvGrpSpPr><p:grpSpPr/>${inner}</p:grpSp>`;
  return buildPptx({ slides: [{ shapes: inner }] });
}

/** A deck whose slide part declares a DTD. */
function dtdDeck(): Uint8Array {
  const files = unzipSync(buildPptx({ slides: [{ shapes: '' }] }));
  const slide = strFromU8(files['ppt/slides/slide1.xml']!);
  files['ppt/slides/slide1.xml'] = strToU8(slide.replace('<p:sld ', '<!DOCTYPE p:sld []><p:sld '));
  return zipSync(files);
}

test('a DTD is blocked, as DOCX Editor blocks it', async () => {
  const editor = createPptxEditor();
  await editor.load(dtdDeck());
  expect(editor.snapshot().openFailure).toEqual({ code: 'xml-dtd-forbidden', reason: 'blocked' });
  editor.destroy();
});

test('an open that throws ends in a refusal, not the loading screen', async () => {
  const editor = createPptxEditor();
  const errors: unknown[] = [];
  const log = console.error;
  console.error = (...args: unknown[]) => void errors.push(args);
  const bytes = {
    get length(): number {
      throw new TypeError('boom');
    },
  } as unknown as Uint8Array;
  const result = await editor.load(bytes).finally(() => (console.error = log));
  expect(result.ok).toBe(false);
  const snapshot = editor.snapshot();
  expect(snapshot.status).toBe('error');
  expect(snapshot.openFailure).toEqual({ code: 'failed', reason: 'unknown' });
  expect(errors).toHaveLength(1);
  editor.destroy();
});

test('an XML limit has a code of its own and reads as too large', async () => {
  const editor = createPptxEditor();
  await editor.load(deepDeck(400));
  expect(editor.snapshot().openFailure).toEqual({ code: 'xml-depth-limit', reason: 'too-large' });
  editor.destroy();
});

test('the report holds the code, the version, and the browser, never the file content', async () => {
  const editor = createPptxEditor();
  await editor.load(deepDeck(400));
  const report = openFailureReport(editor.snapshot().openFailure!, {
    userAgent: 'TestBrowser/1.0',
    time: new Date('2026-11-07T10:00:00Z'),
  });
  expect(report).toBe(
    [
      'PPTX Editor could not open a file.',
      'Code: xml-depth-limit',
      'Reason: too-large',
      'Editor version: source',
      'Browser: TestBrowser/1.0',
      'Time: 2026-11-07T10:00:00.000Z',
    ].join('\n')
  );
  expect(report).not.toContain('Private text');
  editor.destroy();
});
