/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Schema singletons under concurrent edits: two peers that both replace one shape's fill or
// line, or both add its first text properties, end with one element, not two.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  find,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { META_KEY, NODES_KEY } from '../collaboration/yjs/schema.ts';

const IN = 914_400;
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
            text: para(run('Hello')),
          }) + shape({ id: 3, x: 4 * IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('ED7D31') }),
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

async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

async function pair() {
  const da = new Y.Doc();
  const db = new Y.Doc();
  const l = link(da, db);
  const bytes = deck();
  const a = await replica(da, 'a', bytes);
  const b = await replica(db, 'b', bytes);
  return { a, b, l };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const count = (text: string, re: RegExp) => text.match(re)?.length ?? 0;
/** The first shape's `p:spPr`, as text. */
const firstSpPr = (e: PptxEditorInstance) => /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(xml(e))![0];

function selectFirst(e: PptxEditorInstance) {
  const id = e.session!.layout(0)!.items[0].element.nodeId;
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
}

const FILLS = /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b/g;

describe('schema singletons under concurrent edits', () => {
  test('two fill changes on one shape leave one fill', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setShapeFill', color: '00FF00' }).ok).toBe(true);
    expect(b.exec({ type: 'setShapeFill', color: 'FF0000' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(firstSpPr(a), FILLS)).toBe(1);
    expect(readPptx(a.save()!).ok).toBe(true);
  });

  test('a fill change against no fill leaves one fill', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setShapeFill', color: null }).ok).toBe(true);
    expect(b.exec({ type: 'setShapeGradient', gradient: GRADIENT }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(firstSpPr(a), FILLS)).toBe(1);
  });

  test('two first line colors leave one a:ln with one fill', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setLineColor', color: '00FF00' }).ok).toBe(true);
    expect(b.exec({ type: 'setLineColor', color: 'FF0000' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const spPr = firstSpPr(a);
    expect(count(spPr, /<a:ln\b/g)).toBe(1);
    const ln = /<a:ln\b[\s\S]*?<\/a:ln>/.exec(spPr)![0];
    expect(count(ln, FILLS)).toBe(1);
  });

  test('a change after the merge still reaches the other replica', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    a.exec({ type: 'setShapeFill', color: '00FF00' });
    b.exec({ type: 'setShapeFill', color: 'FF0000' });
    l.flush();
    expect(a.exec({ type: 'setShapeFill', color: '0000FF' }).ok).toBe(true);
    expect(firstSpPr(b)).toContain('0000FF');
    expect(fingerprint(a)).toBe(fingerprint(b));
    // Undo goes back to the merged fill on both replicas, and redo comes back.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(firstSpPr(b)).not.toContain('0000FF');
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expect(firstSpPr(b)).toContain('0000FF');
    expect(count(firstSpPr(b), FILLS)).toBe(1);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });

  test('a fill change while a peer deletes the shape converges on the delete', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setShapeFill', color: '00FF00' }).ok).toBe(true);
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<p:sp>/g)).toBe(1);
    expect(readPptx(a.save()!).ok).toBe(true);
  });

  test('two slide backgrounds leave one p:bg', async () => {
    const { a, b, l } = await pair();
    l.pause();
    expect(
      a.exec({ type: 'setSlideBackground', background: { kind: 'solid', color: '00FF00' } }).ok
    ).toBe(true);
    expect(
      b.exec({ type: 'setSlideBackground', background: { kind: 'solid', color: 'FF0000' } }).ok
    ).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<p:bg>/g)).toBe(1);
    expect(count(xml(a), /<p:bgPr>/g)).toBe(1);
  });

  test('two new shapes from two peers stay two shapes', async () => {
    const { a, b, l } = await pair();
    l.pause();
    expect(a.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(b.exec({ type: 'insertShape', preset: 'ellipse' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<p:sp>/g)).toBe(4);
  });
});

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

describe('other single-instance children', () => {
  test('two first paragraph properties leave one a:pPr', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
    expect(b.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<a:pPr\b/g)).toBe(1);
  });

  test('two first outline widths leave one a:ln', async () => {
    const { a, b, l } = await pair();
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setLineWidth', points: 2 }).ok).toBe(true);
    expect(b.exec({ type: 'setLineWidth', points: 6 }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(firstSpPr(a), /<a:ln\b/g)).toBe(1);
  });

  test('two first crops leave one a:srcRect, and two first a:extLst one', async () => {
    const { a, b, l } = await pair();
    expect(
      a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
    ).toBe(true);
    l.pause();
    // After `a:blip`, where a crop goes (CT_BlipFillProperties).
    insertAt(a, 'p:blipFill', 1, 'srcRect');
    insertAt(b, 'p:blipFill', 1, 'srcRect');
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<a:srcRect\b/g)).toBe(1);
    l.pause();
    insertAt(a, 'p:cNvPr', 0, 'extLst');
    insertAt(b, 'p:cNvPr', 0, 'extLst');
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(a), /<a:extLst\b/g)).toBe(1);
    expect(readPptx(a.save()!).ok).toBe(true);
  });
});

/** A pair whose Yjs client ids are fixed, so a concurrent insert order is known. */
async function pairWithIds(first: number, second: number, bytes: Uint8Array) {
  const da = new Y.Doc();
  const db = new Y.Doc();
  da.clientID = first;
  db.clientID = second;
  const l = link(da, db);
  const a = await replica(da, 'a', bytes);
  const b = await replica(db, 'b', bytes);
  return { a, b, l, da };
}

/** Insert an empty `local` element at `index` in the first `name` element, as one store write. */
function insertAt(e: PptxEditorInstance, name: string, index: number, local: string, rad = '') {
  const part = e.session!.info.slides[0].partName;
  const parent = find(xmlPartOf(e.session!.store.package, part)!.root, name) as Element;
  const attrs = rad ? [{ ns: '', local: 'rad', value: rad }] : [];
  const r = e.session!.store.transact([
    {
      op: 'insertNode',
      part,
      parent: parent.id,
      index,
      spec: { ns: NS_A, local, prefix: 'a', attrs },
    },
  ]);
  expect(r.ok).toBe(true);
}

/** Remove the first `name` element, as one store write. */
function removeFirst(e: PptxEditorInstance, name: string) {
  const part = e.session!.info.slides[0].partName;
  const node = find(xmlPartOf(e.session!.store.package, part)!.root, name) as Element;
  expect(e.session!.store.transact([{ op: 'removeNode', part, node: node.id }]).ok).toBe(true);
}

const SHADOW =
  '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:prstClr val="black"/></a:outerShdw></a:effectLst>';

describe('schema order after concurrent singletons', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`two new effects stay in CT_EffectList order (client ids ${first}, ${second})`, async () => {
      const bytes = buildPptx({
        slides: [
          {
            layout: 'title',
            shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, effects: SHADOW }),
          },
        ],
      });
      const { a, b, l } = await pairWithIds(first, second, bytes);
      l.pause();
      insertAt(a, 'a:effectLst', 1, 'reflection');
      insertAt(b, 'a:effectLst', 1, 'softEdge', '12700');
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const list = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/.exec(xml(a))![1]!;
      const names = [...list.matchAll(/<a:(\w+)/g)].map((m) => m[1]).filter((n) => n !== 'prstClr');
      expect(names).toEqual(['outerShdw', 'reflection', 'softEdge']);
    });

    test(`paragraph properties stay before a new line break (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      l.pause();
      expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      insertAt(b, 'a:p', 0, 'br');
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const p = /<a:p>([\s\S]*?)<\/a:p>/.exec(xml(a))![1]!;
      expect(p.startsWith('<a:pPr')).toBe(true);
      expect(p).toContain('<a:br');
    });

    test(`an outline and an effect list stay in CT_ShapeProperties order (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(b);
      l.pause();
      insertAt(a, 'p:spPr', 3, 'effectLst');
      expect(b.exec({ type: 'setLineColor', color: 'FF0000' }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const spPr = firstSpPr(a);
      expect(spPr.indexOf('<a:ln')).toBeGreaterThan(0);
      expect(spPr.indexOf('<a:ln')).toBeLessThan(spPr.indexOf('<a:effectLst'));
    });
  }
});

const GRADIENT = {
  kind: 'linear',
  angle: 90,
  rotWithShape: true,
  stops: [
    { pos: 0, color: '#FFFFFF' },
    { pos: 1, color: '#000000' },
  ],
} as const;

describe('hidden singletons', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`an edit by the peer whose outline is hidden shows everywhere (client ids ${first}, ${second})`, async () => {
      const { a, b, l, da } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineColor', color: '00FF00' }).ok).toBe(true);
      expect(b.exec({ type: 'setLineColor', color: 'FF0000' }).ok).toBe(true);
      l.flush();
      const loser = firstSpPr(a).includes('00FF00') ? b : a;
      expect(loser.exec({ type: 'setLineWidth', points: 6 }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(a)).toContain('w="76200"');
      const dc = new Y.Doc();
      Y.applyUpdate(dc, Y.encodeStateAsUpdate(da));
      const c = await replica(dc, 'c', deck());
      expect(fingerprint(c)).toBe(fingerprint(a));
    });
  }

  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`the hidden outline keeps its own color for the winner's undo (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineColor', color: '00FF00' }).ok).toBe(true);
      expect(b.exec({ type: 'setLineColor', color: 'FF0000' }).ok).toBe(true);
      l.flush();
      const winner = firstSpPr(a).includes('00FF00') ? a : b;
      const loser = winner === a ? b : a;
      // The loser edits the outline it sees, which is the winner's copy.
      expect(loser.exec({ type: 'setLineWidth', points: 6 }).ok).toBe(true);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      // The loser's own outline shows with its own color and its width: the outline merges, so
      // the loser's width went to its own copy.
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(a)).toContain(winner === a ? 'FF0000' : '00FF00');
      expect(firstSpPr(a)).toContain('w="76200"');
    });
  }
});

describe('undo after a merged first singleton', () => {
  test("the winner's undo shows the other peer's outline, kept as a companion", async () => {
    const { a, b, l } = await pairWithIds(2, 1, deck());
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setLineColor', color: '00FF00' }).ok).toBe(true);
    expect(b.exec({ type: 'setLineColor', color: 'FF0000' }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const winner = firstSpPr(a).includes('00FF00') ? a : b;
    expect(winner.exec({ type: 'undo' }).ok).toBe(true);
    // The losing copy stayed in the room, so it shows again (as in DOCX Editor).
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(firstSpPr(a), /<a:ln\b/g)).toBe(1);
    expect(firstSpPr(a)).toContain(winner === a ? 'FF0000' : '00FF00');
    expect(readPptx(a.save()!).ok).toBe(true);
  });
});

describe('property containers merge', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`an alignment and bullets from two peers both stay (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleBullets' }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const text = xml(a);
      expect(count(text, /<a:pPr\b/g)).toBe(1);
      const pPr = /<a:pPr\b[\s\S]*?<\/a:pPr>/.exec(text)![0];
      expect(pPr).toContain('algn="ctr"');
      expect(pPr).toMatch(/<a:bu(Char|AutoNum)\b/);
      expect(readPptx(a.save()!).ok).toBe(true);
      // A later edit saves the merged element, and both replicas keep it.
      expect(a.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(count(xml(b), /<a:pPr\b/g)).toBe(1);
      expect(xml(b)).toMatch(/<a:bu(Char|AutoNum)\b/);
    });
  }

  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`a later edit elsewhere in the paragraph keeps the merge (client ids ${first}, ${second})`, async () => {
      const { a, b, l, da } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleBullets' }).ok).toBe(true);
      l.flush();
      // An edit of the run, not of the paragraph properties.
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      const dc = new Y.Doc();
      Y.applyUpdate(dc, Y.encodeStateAsUpdate(da));
      const c = await replica(dc, 'c', deck());
      expect(fingerprint(c)).toBe(fingerprint(a));
      expect(xml(c)).toContain('algn="ctr"');
      expect(xml(c)).toMatch(/<a:bu(Char|AutoNum)\b/);
    });
  }

  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`each peer's undo reverts only that peer's properties (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'toggleBullets' }).ok).toBe(true);
      expect(b.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
      l.flush();
      // Later edits by A: one of the run, one of the merged paragraph properties.
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(a.exec({ type: 'setLineSpacing', multiple: 1.5 }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      const pPr = /<a:pPr\b[\s\S]*?<\/a:pPr>/.exec(xml(a))![0];
      expect(pPr).not.toContain('algn="r"');
      expect(pPr).toMatch(/<a:bu(Char|AutoNum)\b/);
      expect(pPr).toContain('<a:lnSpc');
    });

    test(`an edit of a property a peer wrote goes to that peer's record (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'toggleBullets' }).ok).toBe(true);
      expect(b.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
      l.flush();
      expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(xml(b)).toContain('algn="ctr"');
      // A's undo of that edit gives B's alignment back.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).toContain('algn="r"');
      expect(xml(a)).toMatch(/<a:bu(Char|AutoNum)\b/);
    });
  }

  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`an edit after the merge outlives the peer's undo (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'toggleBullets' }).ok).toBe(true);
      expect(b.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
      l.flush();
      // A changes the alignment B wrote; then B undoes its own older step.
      expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).toContain('algn="ctr"');
      expect(xml(a)).toMatch(/<a:bu(Char|AutoNum)\b/);
    });
  }

  test("a source file's own duplicates stay as loaded on every replica", async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'title',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            text: para(run('Hi'), '<a:pPr algn="r"/><a:pPr marL="100"/>'),
          }),
        },
      ],
    });
    const { a, b } = await pairWithIds(1, 2, bytes);
    expect(fingerprint(b)).toBe(fingerprint(a));
    expect(count(xml(b), /<a:pPr\b/g)).toBe(2);
  });

  test("a source file's own duplicate content types stay as loaded on every replica", async () => {
    const files = unzipSync(deck());
    const types = strFromU8(files['[Content_Types].xml']!);
    const png = '<Default Extension="png" ContentType="image/png"/>';
    files['[Content_Types].xml'] = strToU8(
      types.replace('<Default Extension="xml"', `${png}${png}<Default Extension="xml"`)
    );
    const { a, b } = await pairWithIds(1, 2, zipSync(files));
    expect(fingerprint(b)).toBe(fingerprint(a));
    const loaded = serializeXmlPart(xmlPartOf(b.session!.store.package, '/[Content_Types].xml')!);
    expect(count(loaded, /<Default [^>]*Extension="png"/g)).toBe(2);
  });

  test("one peer's undo keeps the other peer's properties", async () => {
    const { a, b, l } = await pairWithIds(2, 1, deck());
    selectFirst(a);
    selectFirst(b);
    l.pause();
    expect(a.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleBullets' }).ok).toBe(true);
    l.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(xml(a)).not.toContain('algn="ctr"');
    expect(xml(a)).toMatch(/<a:bu(Char|AutoNum)\b/);
  });
});

describe('records an undo removes', () => {
  test('after the winner undoes a replaced fill, every replica and a new one agree', async () => {
    const da = new Y.Doc();
    const db = new Y.Doc();
    da.clientID = 2;
    db.clientID = 1;
    const l = link(da, db);
    const bytes = deck();
    const a = await replica(da, 'a', bytes);
    const b = await replica(db, 'b', bytes);
    selectFirst(a);
    selectFirst(b);
    l.pause();
    a.exec({ type: 'setShapeFill', color: '00FF00' });
    b.exec({ type: 'setShapeFill', color: 'FF0000' });
    l.flush();
    const winner = firstSpPr(a).includes('00FF00') ? a : b;
    expect(winner.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    // The other peer's fill shows again.
    expect(count(firstSpPr(a), FILLS)).toBe(1);
    expect(firstSpPr(a)).toContain(winner === a ? 'FF0000' : '00FF00');
    // A replica that joins now reads the same deck from the room.
    const dc = new Y.Doc();
    Y.applyUpdate(dc, Y.encodeStateAsUpdate(da));
    const c = await replica(dc, 'c', bytes);
    expect(fingerprint(c)).toBe(fingerprint(a));
    // Later concurrent edits of the shape keep a valid part.
    l.pause();
    expect(a.exec({ type: 'setLineWidth', points: 3 }).ok).toBe(true);
    expect(b.exec({ type: 'setLineWidth', points: 5 }).ok).toBe(true);
    l.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(firstSpPr(a).replace(/<a:ln\b[\s\S]*?<\/a:ln>/g, ''), FILLS)).toBe(1);
    expect(readPptx(a.save()!).ok).toBe(true);
  });

  test('a second load on one session does not reuse its earlier keys', async () => {
    const da = new Y.Doc();
    const db = new Y.Doc();
    da.clientID = 7;
    db.clientID = 8;
    link(da, db);
    const bytes = deck();
    const a = await replica(da, 'a', bytes);
    const b = await replica(db, 'b', bytes);
    // The second load attaches a new mirror to the same room and client id.
    expect((await a.load(bytes)).ok).toBe(true);
    expect(a.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(b), /<p:sp>/g)).toBe(3);
  });
});

describe('minted keys', () => {
  test('a peer key with a huge number does not stop new keys', async () => {
    const da = new Y.Doc();
    const db = new Y.Doc();
    da.clientID = 7;
    db.clientID = 8;
    link(da, db);
    const bytes = deck();
    const a = await replica(da, 'a', bytes);
    const b = await replica(db, 'b', bytes);
    // A hostile peer writes a key in a's name with Number.MAX_SAFE_INTEGER as its number.
    db.getMap(NODES_KEY).set(
      `${(7).toString(36)}.${Number.MAX_SAFE_INTEGER.toString(36)}`,
      new Y.Map()
    );
    expect((await a.load(bytes)).ok).toBe(true);
    expect(a.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(a.exec({ type: 'insertShape', preset: 'ellipse' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(count(xml(b), /<p:sp>/g)).toBe(4);
    expect(readPptx(b.save()!).ok).toBe(true);
  });
});

describe('room format', () => {
  test('a room of the earlier format (schema 1) refuses to attach', async () => {
    const doc = new Y.Doc();
    doc.getMap(META_KEY).set('schema', 1);
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: 'a', name: 'a' },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    await editor.load(deck());
    expect(session.statusSnapshot()).toMatchObject({
      status: 'error',
      reason: { code: 'schema-version-mismatch' },
    });
  });
});

describe('writes after concurrent singletons', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`an undo after a later edit does not bring back children a peer deleted (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      l.pause();
      insertAt(a, 'p:spPr', 3, 'ln');
      insertAt(b, 'p:spPr', 3, 'effectLst');
      l.flush();
      // A's next edit of the shape properties must not move B's child in the room.
      expect(a.exec({ type: 'setShapeFill', color: '00FF00' }).ok).toBe(true);
      removeFirst(b, 'a:effectLst');
      removeFirst(b, 'a:ln');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(a)).not.toContain('<a:effectLst');
      expect(firstSpPr(a)).not.toContain('<a:ln');
      expect(firstSpPr(a)).toContain('4472C4');
    });

    test(`a loser's edit keeps its shadowed property for the winner's undo (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineSpacing', multiple: 1.5 }).ok).toBe(true);
      expect(b.exec({ type: 'setLineSpacing', multiple: 2 }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const winner = xml(a).includes('150000') ? a : b;
      const loser = winner === a ? b : a;
      // An edit of the merged paragraph properties by the peer whose spacing is hidden.
      expect(loser.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      const pPr = /<a:pPr\b[\s\S]*?(<\/a:pPr>|\/>)/.exec(xml(a))![0];
      expect(pPr).toContain('algn="ctr"');
      expect(pPr).toContain(winner === a ? '200000' : '150000');
    });

    test(`a line width from one peer merges with a line color from the other, and the first copy's fill shows (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineColor', color: '00FF00' }).ok).toBe(true);
      expect(b.exec({ type: 'setLineWidth', points: 6 }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const spPr = firstSpPr(a);
      expect(count(spPr, /<a:ln\b/g)).toBe(1);
      const ln = /<a:ln\b[\s\S]*?<\/a:ln>/.exec(spPr)![0];
      // The width is an attribute, so it merges. The fill is a choice: the first copy's shows.
      // A first width on a shape with no outline also writes a default color (tx1), so with
      // client ids (2, 1) that color shows instead of the peer's green.
      expect(ln).toContain('w="76200"');
      expect(count(ln, FILLS)).toBe(1);
      expect(ln).toContain(first === 1 ? 'val="00FF00"' : 'val="tx1"');
      expect(readPptx(a.save()!).ok).toBe(true);
      // Each author's undo reverts only that author's value.
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(a)).not.toContain('w="76200"');
      expect(firstSpPr(a)).toContain('00FF00');
    });
  }
});

describe('a loser replaces a shadowed slot', () => {
  const lnSpc = (e: PptxEditorInstance) => /<a:lnSpc>[\s\S]*?<\/a:lnSpc>/.exec(xml(e))?.[0] ?? '';
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`the loser's new line spacing shows on every replica (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineSpacing', multiple: 1.5 }).ok).toBe(true);
      expect(b.exec({ type: 'setLineSpacing', multiple: 2 }).ok).toBe(true);
      l.flush();
      const winner = xml(a).includes('150000') ? a : b;
      const loser = winner === a ? b : a;
      expect(loser.exec({ type: 'setLineSpacing', multiple: 3 }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(lnSpc(a)).toContain('300000');
      // A later remote edit rebuilds the part: the loser's value stays.
      selectFirst(winner);
      expect(winner.exec({ type: 'nudge', dx: IN / 4, dy: 0 }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(lnSpc(loser)).toContain('300000');
      expect(readPptx(a.save()!).ok).toBe(true);
    });

    test(`the loser's new line color shows on every replica (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pairWithIds(first, second, deck());
      selectFirst(a);
      selectFirst(b);
      l.pause();
      expect(a.exec({ type: 'setLineWidth', points: 2 }).ok).toBe(true);
      expect(b.exec({ type: 'setLineWidth', points: 5 }).ok).toBe(true);
      l.flush();
      const winner = firstSpPr(a).includes('w="25400"') ? a : b;
      const loser = winner === a ? b : a;
      expect(loser.exec({ type: 'setLineColor', color: 'FF00FF' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(a)).toContain('FF00FF');
      selectFirst(winner);
      expect(winner.exec({ type: 'nudge', dx: IN / 4, dy: 0 }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(firstSpPr(loser)).toContain('FF00FF');
    });
  }

  test('with three peers, a shadowed copy of another peer does not hide the new value', async () => {
    for (const ids of [
      [1, 2, 3],
      [3, 2, 1],
      [2, 3, 1],
    ]) {
      const docs = ids.map((id) => {
        const d = new Y.Doc();
        d.clientID = id;
        return d;
      });
      const links = [link(docs[0]!, docs[1]!), link(docs[1]!, docs[2]!), link(docs[0]!, docs[2]!)];
      const bytes = deck();
      const eds: PptxEditorInstance[] = [];
      for (const [i, d] of docs.entries()) eds.push(await replica(d, `p${i}`, bytes));
      for (const e of eds) selectFirst(e);
      for (const k of links) k.pause();
      [1.5, 2, 2.5].forEach((m, i) =>
        expect(eds[i]!.exec({ type: 'setLineSpacing', multiple: m }).ok).toBe(true)
      );
      for (const k of links) k.flush();
      const shown = lnSpc(eds[0]!);
      for (const e of eds) {
        expect(fingerprint(e)).toBe(fingerprint(eds[0]!));
        // Every peer but the winner sets a new spacing; the last one shows everywhere.
        const value = ['150000', '200000', '250000'].find((v) => shown.includes(v));
        if (value === ['150000', '200000', '250000'][eds.indexOf(e)]) continue;
        expect(e.exec({ type: 'setLineSpacing', multiple: 3 + eds.indexOf(e) }).ok).toBe(true);
        const want = String((3 + eds.indexOf(e)) * 100000);
        for (const other of eds) {
          expect(fingerprint(other)).toBe(fingerprint(e));
          expect(lnSpc(other)).toContain(want);
        }
      }
    }
  });
});
