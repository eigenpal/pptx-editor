// The automation `addShape` operation with a `picture` shape, which `ShapeCollection.addPicture`
// of the object model sends, through two editors on two linked Y.Docs (`AGENTS.md`
// "Collaboration verification"). The write adds a `p:pic` (ECMA-376 Part 1, §19.3.1.37) over a
// new Image part (§15.2.14); both replicas must converge.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

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

/** One slide with a text box. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank' as const,
        shapes: shape({
          id: 2,
          name: 'Title',
          x: 914_400,
          y: 914_400,
          cx: 914_400,
          cy: 914_400,
          text: para(run('Report')),
        }),
      },
    ],
  });

/** The PNG signature and an `IHDR` chunk of `width` × `height` pixels, in base64. */
function png(width: number, height: number): string {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(b.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  b.set([8, 6, 0, 0, 0], 24);
  return Buffer.from(b).toString('base64');
}

async function replica(doc: Y.Doc, actorId: string): Promise<PptxEditorInstance> {
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
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

function execute(host: AutomationHost, operations: AutomationOperation[]): void {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

/**
 * Add a picture to the first slide with the operations `addPicture` sends: the lookups, then the
 * write. `packages/editor-api/src/__tests__/add-picture.test.ts` checks that it sends these.
 */
function addPicture(host: AutomationHost, base64: string): void {
  execute(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index: 0 } },
    { op: 'addShape', container: { result: 1 }, shape: { kind: 'picture', base64 }, box: {} },
  ]);
}

/** Move the text box of the first slide: a concurrent edit of the same slide. */
function moveTitle(host: AutomationHost): void {
  execute(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 1 }, key: { id: '2' } },
    { op: 'setShape', shape: { result: 2 }, props: { left: 400, top: 300 } },
  ]);
}

const SLIDE = '/ppt/slides/slide1.xml';
const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

/** The saved package of `e`, read back. */
function saved(e: PptxEditorInstance): PptxPackage {
  const read = readPptx(e.save()!);
  if (!read.ok) throw new Error('reopen');
  return read.pkg;
}

/** The media part each picture of the first slide embeds. */
function embedded(pkg: PptxPackage): (string | undefined)[] {
  const tree = child(child(xmlPartOf(pkg, SLIDE)!.root, 'p:cSld'), 'p:spTree');
  const rels = relationshipsOf(pkg, SLIDE);
  return elements(tree)
    .filter((e) => e.local === 'pic')
    .map((pic) => {
      const id = attr(child(child(pic, 'p:blipFill'), 'a:blip')!, 'r:embed');
      return rels.find((r) => r.id === id)?.target;
    });
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('ShapeCollection.addPicture on two peers', () => {
  test('a picture added on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    addPicture(createBrowserAutomationHost(b), png(40, 20));
    expect(fingerprint(a)).toBe(fingerprint(b));
    const [target] = embedded(saved(a));
    expect(target).toMatch(/^\/ppt\/media\/image\d+\.png$/);
    expect(saved(a).parts.get(target!)?.kind).toBe('binary');
    expect(canonicalPackageFingerprint(saved(a))).toBe(canonicalPackageFingerprint(saved(b)));
  });

  test('a picture added while a peer edits the same slide keeps both edits', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    addPicture(createBrowserAutomationHost(a), png(40, 20));
    moveTitle(createBrowserAutomationHost(b));
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(embedded(saved(a))).toHaveLength(1);
    const tree = child(child(xmlPartOf(saved(b), SLIDE)!.root, 'p:cSld'), 'p:spTree');
    const title = elements(tree).find((e) => e.local === 'sp')!;
    const off = child(child(child(title, 'p:spPr'), 'a:xfrm'), 'a:off')!;
    expect(attr(off, 'x')).toBe(String(400 * 12_700));
  });

  test('pictures added on both peers at once keep two media parts', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    addPicture(createBrowserAutomationHost(a), png(40, 20));
    addPicture(createBrowserAutomationHost(b), png(30, 30));
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    const targets = embedded(saved(a));
    expect(targets).toHaveLength(2);
    expect(new Set(targets).size).toBe(2);
    for (const t of targets) expect(saved(a).parts.get(t!)?.kind).toBe('binary');
  });

  test('undo of the picture on one peer removes it on both', async () => {
    const { a, b } = await pair();
    addPicture(createBrowserAutomationHost(b), png(40, 20));
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(embedded(saved(a))).toHaveLength(0);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
