import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint } from '@pptx-editor.dev/core/store';
import { buildPptx, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

// A pointer drag of a shape on one of two editors on linked Y.Docs: the peer gets nothing while
// the pointer moves, one update set on release, and both replicas converge. One undo on the
// dragging peer reverts the whole drag on both.

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'A',
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
          }) +
          shape({
            id: 3,
            name: 'B',
            x: 6 * IN,
            y: 4 * IN,
            cx: IN,
            cy: IN,
            fill: solid('ED7D31'),
          }),
      },
    ],
  });

// A DOM for this file only: the test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());
const frame = (): Promise<void> => new Promise((r) => setTimeout(r, 40));

async function replica(ydoc: Y.Doc, actorId: string) {
  const session = createYjsCollaborationSession({
    doc: ydoc,
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

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const xfrmOf = (e: PptxEditorInstance) => e.session!.layout(0)!.items[0]!.element.xfrm;

/** Two linked editors; A is attached to a surface that pointer events drive. */
async function room() {
  const g = globalThis as { ResizeObserver?: unknown };
  const had = 'ResizeObserver' in g;
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sent = { count: 0 };
  docA.on('update', (u: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    sent.count += 1;
    Y.applyUpdate(docB, u, 'link');
  });
  docB.on('update', (u: Uint8Array, origin: unknown) => {
    if (origin !== 'link') Y.applyUpdate(docA, u, 'link');
  });
  const a = await replica(docA, 'a');
  const b = await replica(docB, 'b');
  expect(fingerprint(a)).toBe(fingerprint(b));
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  a.attach(host);
  a.setZoom(1);
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({ left: 0, top: 0, right: 1280, bottom: 720, width: 1280, height: 720 }) as DOMRect;
  const width = a.session!.layout(0)!.slide.width;
  const pointer = (type: string, target: EventTarget, x: number, y: number) =>
    target.dispatchEvent(
      new win.MouseEvent(type, {
        clientX: (x / width) * 1280,
        clientY: (y / width) * 1280,
        button: 0,
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  const done = () => {
    a.destroy();
    b.destroy();
    host.remove();
    if (!had) delete g.ResizeObserver;
  };
  /** Press on shape A and move it by 1.5 in and 0.75 in, in six steps. */
  const dragA = () => {
    pointer('pointerdown', stage, 2 * IN, 1.5 * IN);
    sent.count = 0;
    for (let i = 1; i <= 6; i += 1)
      pointer('pointermove', doc, 2 * IN + i * 0.25 * IN, 1.5 * IN + i * 0.125 * IN);
  };
  const release = () => pointer('pointerup', doc, 3.5 * IN, 2.25 * IN);
  return { a, b, host, sent, dragA, release, done };
}

const preview = (host: HTMLElement) => host.querySelector('[data-kind="drag-preview"]');

describe('a shape drag in a collaboration room', () => {
  test('sends one update set on release, converges, and undoes in one step', async () => {
    const { a, b, sent, dragA, release, done } = await room();
    dragA();
    expect(sent.count).toBe(0);
    expect(xfrmOf(b)).toMatchObject({ x: IN, y: IN });
    release();
    expect(sent.count).toBeGreaterThan(0);
    expect(xfrmOf(a)).toMatchObject({ x: 2.5 * IN, y: 1.75 * IN });
    expect(xfrmOf(b)).toMatchObject({ x: 2.5 * IN, y: 1.75 * IN });
    expect(fingerprint(a)).toBe(fingerprint(b));

    a.exec({ type: 'undo' });
    expect(xfrmOf(a)).toMatchObject({ x: IN, y: IN });
    expect(xfrmOf(b)).toMatchObject({ x: IN, y: IN });
    expect(fingerprint(a)).toBe(fingerprint(b));
    done();
  });

  test('a peer edit to another shape repaints the slide and keeps the drag', async () => {
    const { a, b, host, dragA, release, done } = await room();
    dragA();
    const slide = host.querySelector('.pptx-surface__slide');
    // The peer fills shape B green.
    b.setSelection({ kind: 'shapes', nodeIds: [b.session!.layout(0)!.items[1]!.element.nodeId] });
    expect(b.exec({ type: 'setShapeFill', color: '#70AD47' }).ok).toBe(true);
    expect(host.querySelector('.pptx-surface__slide')).not.toBe(slide);
    expect(preview(host)).not.toBeNull();
    release();
    // Both edits stay: the move and the peer's fill.
    expect(xfrmOf(a)).toMatchObject({ x: 2.5 * IN, y: 1.75 * IN });
    expect(xfrmOf(b)).toMatchObject({ x: 2.5 * IN, y: 1.75 * IN });
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(host.querySelector('.pptx-surface__slide')?.innerHTML.toLowerCase()).toContain('70ad47');
    done();
  });

  test('a peer edit to the dragged shape cancels the drag', async () => {
    const { a, b, host, sent, dragA, release, done } = await room();
    dragA();
    // The peer fills the dragged shape green.
    b.setSelection({ kind: 'shapes', nodeIds: [b.session!.layout(0)!.items[0]!.element.nodeId] });
    expect(b.exec({ type: 'setShapeFill', color: '#70AD47' }).ok).toBe(true);
    expect(preview(host)).toBeNull();
    // The region is cleared, then says the cancel on the next frame.
    await frame();
    expect(host.querySelector('[aria-live]')?.textContent).toBe('canvas.dragCanceled');
    sent.count = 0;
    release();
    // The release writes nothing: the shape keeps its place and the peer's fill.
    expect(sent.count).toBe(0);
    expect(xfrmOf(a)).toMatchObject({ x: IN, y: IN });
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(host.querySelector('.pptx-surface__slide')?.innerHTML.toLowerCase()).toContain('70ad47');
    done();
  });

  test('a peer delete during the drag leaves no preview and writes nothing on release', async () => {
    const { a, b, host, dragA, release, done } = await room();
    dragA();
    b.setSelection({ kind: 'shapes', nodeIds: [b.session!.layout(0)!.items[0]!.element.nodeId] });
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    expect(a.session!.layout(0)!.items).toHaveLength(1);
    expect(preview(host)).toBeNull();
    release();
    expect(preview(host)).toBeNull();
    expect(a.snapshot().lastRejection ?? null).toBeNull();
    expect(a.session!.layout(0)!.items).toHaveLength(1);
    expect(fingerprint(a)).toBe(fingerprint(b));
    done();
  });
});
