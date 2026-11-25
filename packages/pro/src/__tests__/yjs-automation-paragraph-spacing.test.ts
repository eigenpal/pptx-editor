/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Paragraph spacing and indents through the automation `setParagraph` operation on two editors
// on two linked Y.Docs, as `AGENTS.md` asks: "Verify collaboration convergence when changing
// editing behavior". The `ParagraphFormat` spacing members in `@pptx-editor.dev/editor-api` send
// exactly these operations (its `paragraph-spacing.test.ts` checks the sequence). They write
// `a:lnSpc`, `a:spcBef`, and `a:spcAft` (CT_TextSpacing, ECMA-376 Part 1, 21.1.2.2.5,
// 21.1.2.2.10, 21.1.2.2.9) and `marL` and `indent` in `a:pPr` (CT_TextParagraphProperties,
// ECMA-376 Part 1, 21.1.2.2.7).

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type {
  AutomationOperation,
  AutomationParagraphWrite,
} from '@pptx-editor.dev/core/automation';
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
          text:
            para(run('Hello'), '<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>') +
            para(run('World')),
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

/** Send `setParagraph` writes, as the `ParagraphFormat` members do: one operation each. */
function format(
  editor: PptxEditorInstance,
  writes: readonly AutomationParagraphWrite[],
  span?: { start: number; length: number }
): void {
  const operations: AutomationOperation[] = [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
    ...writes.map((paragraph): AutomationOperation => ({
      op: 'setParagraph',
      shape: { result: 1 },
      ...(span ? { span } : {}),
      paragraph,
    })),
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

/** The `a:pPr` of the first paragraph. */
const firstProps = (e: PptxEditorInstance): string =>
  /<a:p>(<a:pPr\b[^>]*?(?:\/>|>.*?<\/a:pPr>))/.exec(slideXml(e))?.[1] ?? '';

const HELLO = { start: 0, length: 5 };

describe('paragraph spacing through automation under collaboration', () => {
  test('a write on one peer reaches the other', async () => {
    const { a, b } = await replicas();
    format(a, [{ lineSpacing: 18 }, { spaceBefore: 6 }, { leftIndent: 36 }], HELLO);
    converged(a, b);
    expect(firstProps(b)).toBe(
      '<a:pPr marL="457200"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>'
    );
  });

  test('different fields of one paragraph at once converge and keep both writes', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    format(a, [{ spaceAfter: 4 }, { firstLineIndent: -9 }], HELLO);
    format(b, [{ lineSpacing: 24 }, { lineUnitBefore: 0.5 }], HELLO);
    sync.flush();
    converged(a, b);
    const props = firstProps(a);
    expect(props).toContain('indent="-114300"');
    expect(props).toContain('<a:lnSpc><a:spcPct val="200000"/></a:lnSpc>');
    expect(props).toContain('<a:spcBef><a:spcPct val="50000"/></a:spcBef>');
    expect(props).toContain('<a:spcAft><a:spcPts val="400"/></a:spcAft>');
  });

  for (const [label, first, second] of [
    ['two line spacings', { lineSpacing: 15 }, { lineSpacing: 30 }],
    ['a line spacing and a rule', { lineSpacing: 15 }, { lineSpacingRule: 'exactly' }],
    ['a space in points and a space in lines', { spaceBefore: 6 }, { lineUnitBefore: 1 }],
  ] as const)
    test(`${label} on the same paragraph at once converge to one spacing element`, async () => {
      const { a, b, sync } = await replicas();
      sync.pause();
      format(a, [first], HELLO);
      format(b, [second], HELLO);
      sync.flush();
      converged(a, b);
      const props = firstProps(a);
      for (const local of ['lnSpc', 'spcBef'])
        expect((props.match(new RegExp(`<a:${local}>`, 'g')) ?? []).length).toBeLessThanOrEqual(1);
      for (const local of ['lnSpc', 'spcBef']) {
        const holder = new RegExp(`<a:${local}>(.*?)</a:${local}>`).exec(props)?.[1];
        if (holder) expect((holder.match(/<a:spc(Pct|Pts) /g) ?? []).length).toBe(1);
      }
    });

  test('a write while a peer types in the same paragraph converges', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    format(a, [{ lineSpacingRule: 'exactly' }, { spaceAfter: 12 }]);
    const response = createBrowserAutomationHost(b).execute({
      operations: [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
        { op: 'setText', shape: { result: 1 }, span: { start: 5, length: 0 }, text: ' there' },
      ],
    });
    expect(response.ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(slideXml(a)).toContain('there');
    expect(firstProps(a)).toContain('<a:spcAft><a:spcPts val="1200"/></a:spcAft>');
  });
});
