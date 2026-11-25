/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The text highlight through the automation `setFont` operation on two editors on two linked
// Y.Docs, as `AGENTS.md` asks: "Verify collaboration convergence when changing editing
// behavior". `ShapeFont.highlightColor` in
// `@pptx-editor.dev/editor-api` sends exactly these operations (its `text-highlight.test.ts`
// checks the sequence). The run highlight is `a:highlight` (CT_Color, ECMA-376 Part 1,
// 21.1.2.3.4) in `a:rPr` (CT_TextCharacterProperties, ECMA-376 Part 1, 21.1.2.3.9).

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  writePptx,
  xmlPartOf,
  serializeXmlPart,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text: para(
            run('Hello', 'b="1"', '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>') +
              run(' world')
          ),
        }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

async function replicas() {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  docs[0].clientID = 101;
  docs[1].clientID = 202;
  const sync = link(docs[0], docs[1]);
  const editors: PptxEditorInstance[] = [];
  for (const [i, doc] of docs.entries()) {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: `p${i}`, name: `p${i}` },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(deck())).ok).toBe(true);
    editors.push(editor);
  }
  return { a: editors[0]!, b: editors[1]!, sync };
}

/** Send one `setFont` highlight write, as `ShapeFont.highlightColor` does. */
function highlight(
  editor: PptxEditorInstance,
  highlightColor: string | null,
  span?: { start: number; length: number }
): void {
  const operations: AutomationOperation[] = [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
    { op: 'setFont', shape: { result: 1 }, ...(span ? { span } : {}), font: { highlightColor } },
  ];
  const response = createBrowserAutomationHost(editor).execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);

/** Both replicas match, live and after save and reopen. */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = (e: PptxEditorInstance) => {
    const reopened = readPptx(writePptx(e.session!.store.package));
    if (!reopened.ok) throw new Error('reopen failed');
    return canonicalPackageFingerprint(reopened.pkg);
  };
  expect(saved(a)).toBe(saved(b));
}

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);

describe('text highlight through automation under collaboration', () => {
  test('a set on one peer reaches the other', async () => {
    const { a, b } = await replicas();
    highlight(a, '#00FF00', { start: 6, length: 5 });
    converged(a, b);
    expect(slideXml(b)).toContain('<a:highlight><a:srgbClr val="00FF00"/></a:highlight>');
  });

  test('a removal on one peer and a set on another at once converge', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    highlight(a, null);
    highlight(b, '#00B0F0', { start: 6, length: 5 });
    sync.flush();
    converged(a, b);
    const xml = slideXml(a);
    expect(xml).not.toContain('FFFF00');
    // The removal keeps the other run properties, and the set on "world" stays.
    expect(xml).toContain('b="1"');
    expect(xml).toContain('<a:highlight><a:srgbClr val="00B0F0"/></a:highlight>');
  });

  for (const [label, first, second] of [
    ['two sets', '#00FF00', '#00B0F0'],
    ['a set and a removal', '#00FF00', null],
  ] as const)
    test(`${label} on the same run at once converge to at most one a:highlight`, async () => {
      const { a, b, sync } = await replicas();
      sync.pause();
      highlight(a, first, { start: 0, length: 5 });
      highlight(b, second, { start: 0, length: 5 });
      sync.flush();
      converged(a, b);
      const hello = /<a:r><a:rPr[^>]*>(.*?)<\/a:rPr><a:t>Hello<\/a:t>/.exec(slideXml(a))?.[1];
      expect(hello).toBeDefined();
      expect((hello!.match(/<a:highlight>/g) ?? []).length).toBeLessThanOrEqual(1);
      expect(hello).not.toContain('FFFF00');
    });

  test('a removal leaves no empty a:highlight on either peer', async () => {
    const { a, b } = await replicas();
    highlight(b, null);
    converged(a, b);
    for (const e of [a, b]) expect(slideXml(e)).not.toContain('highlight');
  });
});
