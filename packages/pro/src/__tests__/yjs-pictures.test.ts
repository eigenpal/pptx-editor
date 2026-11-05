import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { PNG } from 'pngjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

// Picture commands through two editors on linked Y.Docs (AGENTS.md "Collaboration
// verification"): concurrent edits converge, survive save and reopen, undo and redo act on
// this replica's step only, and a delete wins over a concurrent edit of the same picture.
//
// Not covered here: the same command on the same picture on both replicas writes a singleton
// child twice (two `a:srcRect`), an older mirror defect.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function deck(): Uint8Array {
  const img = new PNG({ width: 8, height: 6 });
  img.data.fill(180);
  const pic = (id: number, x: number) =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes: pic(2, IN) + pic(3, 6 * IN),
      },
    ],
    media: [
      { name: 'image1.png', bytes: new Uint8Array(PNG.sync.write(img)), contentType: 'image/png' },
    ],
  });
}

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

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
function select(e: PptxEditorInstance, index: number): void {
  const id = e.session!.layout(0)!.items[index]!.element.nodeId;
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
}
const image = (e: PptxEditorInstance, index: number) => {
  select(e, index);
  return { picture: e.snapshot().image!, alt: e.snapshot().geometry?.altText };
};

/** Save, reopen in a fresh editor, and return it with the reopened package's fingerprint. */
async function reopen(e: PptxEditorInstance) {
  const saved = readPptx(e.save()!);
  if (!saved.ok) throw new Error('save does not reopen');
  const fresh = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await fresh.load(e.save()!)).ok).toBe(true);
  return { editor: fresh, fingerprint: canonicalPackageFingerprint(saved.pkg) };
}

describe('picture commands with collaboration', () => {
  test('concurrent edits on two pictures and on one picture converge and survive save', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const sync = link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    sync.pause();
    select(a, 0);
    expect(
      a.exec({ type: 'cropImage', crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 0 } })
        .ok
    ).toBe(true);
    expect(a.exec({ type: 'setAltText', description: 'A chart' }).ok).toBe(true);
    select(b, 1);
    expect(b.exec({ type: 'setImageAdjustments', adjustments: { recolor: 'grayscale' } }).ok).toBe(
      true
    );
    expect(b.exec({ type: 'cropToShape', preset: 'ellipse' }).ok).toBe(true);
    // The same picture, other properties: transparency on B, alt text on A above.
    select(b, 0);
    expect(b.exec({ type: 'setImageAdjustments', adjustments: { transparency: 0.5 } }).ok).toBe(
      true
    );
    sync.flush();

    expect(fingerprint(a)).toBe(fingerprint(b));
    for (const e of [a, b]) {
      const first = image(e, 0);
      expect(first.picture.crop.left).toBe(25);
      expect(first.picture.transparency).toBe(0.5);
      expect(first.alt?.description).toBe('A chart');
      const second = image(e, 1);
      expect([second.picture.recolor, second.picture.shape]).toEqual(['grayscale', 'ellipse']);
    }
    // Both saves reopen to the same package, with the merged edits.
    const [ra, rb] = [await reopen(a), await reopen(b)];
    expect(ra.fingerprint).toBe(rb.fingerprint);
    const first = image(ra.editor, 0);
    expect([first.picture.crop.left, first.picture.transparency]).toEqual([25, 0.5]);
    expect(first.alt?.description).toBe('A chart');
    expect(image(ra.editor, 1).picture.shape).toBe('ellipse');
  });

  test('a delete wins over a concurrent edit of the same picture', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const sync = link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    sync.pause();
    select(a, 0);
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    select(b, 0);
    expect(
      b.exec({ type: 'cropImage', crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 0 } })
        .ok
    ).toBe(true);
    expect(b.exec({ type: 'setImageStyle', style: 'simpleFrameBlack' }).ok).toBe(true);
    sync.flush();

    expect(fingerprint(a)).toBe(fingerprint(b));
    for (const e of [a, b]) expect(e.session!.layout(0)!.items.length).toBe(1);
    const [ra, rb] = [await reopen(a), await reopen(b)];
    expect(ra.fingerprint).toBe(rb.fingerprint);
    expect(ra.editor.session!.layout(0)!.items.length).toBe(1);
  });

  test('undo reverts only this replica’s picture edit', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const sync = link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    sync.pause();
    select(a, 0);
    a.exec({ type: 'setAltText', description: 'From A' });
    select(b, 1);
    b.exec({ type: 'setImageStyle', style: 'softEdgeOval' });
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);

    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(image(b, 0).alt?.description).toBe('');
    expect(image(b, 1).picture.style).toBe('softEdgeOval');

    // Redo brings A's step back on both replicas, and B's step stays.
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(image(b, 0).alt?.description).toBe('From A');
    expect(image(b, 1).picture.style).toBe('softEdgeOval');
    const [ra, rb] = [await reopen(a), await reopen(b)];
    expect(ra.fingerprint).toBe(rb.fingerprint);
  });
});
