// Slide and shape pictures through the automation host: the check in the batch, the capability,
// and what the host asks its port to draw. The drawing itself is the editor's slide picture
// paint, which needs a browser; here a port answers fixed bytes.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx, shapeTreeOf, presentationInfo } from '../../store/index.ts';
import { CommandRefusal } from '../commands/op-batch.ts';
import { findShape } from '../targets.ts';
import type { AutomationImageTarget } from '../plan-images.ts';
import {
  createAutomationHost,
  createServerAutomationHost,
  storeAutomationPort,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;
const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      { layout: 'blank', shapes: shape({ id: 2, name: 'Box', x: IN, y: IN, cx: IN, cy: IN }) },
      { layout: 'title', shapes: shape({ id: 4, name: 'Other', x: IN, y: IN, cx: IN, cy: IN }) },
    ],
  });
}

interface Drawing {
  host: AutomationHost;
  store: PptxDocumentStore;
  calls: AutomationImageTarget[];
}

/** A host that draws: its port records each target and answers `draw`. */
function drawing(draw: () => Promise<Uint8Array> = async () => PNG): Drawing {
  const read = readPptx(deck());
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  const calls: AutomationImageTarget[] = [];
  const host = createAutomationHost({
    port: {
      ...storeAutomationPort(store),
      renderImage: (target) => {
        calls.push(target);
        return draw();
      },
    },
    capabilities: {
      document: true,
      save: true,
      events: true,
      selection: false,
      autofit: false,
      images: true,
    },
  });
  return { host, store, calls };
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' }));
}

function refusal(host: AutomationHost, operations: AutomationOperation[]) {
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  return response.results.find((r) => r.status === 'error')!;
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

const slideAt = (index: number): AutomationOperation => ({ op: 'getSlide', key: { index } });

describe('slide and shape pictures', () => {
  test('a slide picture is drawn after the batch, at the size asked', async () => {
    const { host, calls } = drawing();
    const [, prepared] = ok(host, [
      slideAt(1),
      { op: 'prepareImage', target: { result: 0 }, width: 640 },
    ]);
    expect(calls).toEqual([]);
    const answer = await host.renderImage!({ target: handle(prepared), width: 640 });
    expect(answer).toEqual({ ok: true, base64: btoa(String.fromCharCode(...PNG)) });
    expect(calls).toEqual([{ slideIndex: 1, width: 640 }]);
  });

  test('a shape picture names the shape by its tree node', async () => {
    const { host, store, calls } = drawing();
    const [, , prepared] = ok(host, [
      slideAt(0),
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
      { op: 'prepareImage', target: { result: 1 }, height: 50 },
    ]);
    await host.renderImage!({ target: handle(prepared), height: 50 });
    const part = presentationInfo(store.package).slides[0]!.partName;
    const node = findShape(shapeTreeOf(store.package, part), '2')!.el.id;
    expect(calls).toEqual([{ slideIndex: 0, shapeNodeId: node, height: 50 }]);
  });

  test('a shape of a layout has no picture', () => {
    const { host } = drawing();
    const [layout] = ok(host, [{ op: 'getSlideLayout', slide: handle(ok(host, [slideAt(1)])[0]) }]);
    const shapes = ok(host, [{ op: 'listShapes', container: handle(layout) }])[0];
    if (shapes?.kind !== 'handles') throw new Error('not a listing');
    expect(shapes.handles.length).toBeGreaterThan(0);
    const error = refusal(host, [{ op: 'prepareImage', target: shapes.handles[0]! }]);
    expect(error).toMatchObject({
      status: 'error',
      error: { code: 'unsupported-content', detail: 'image-owner' },
    });
  });

  test('a size that is not a positive number refuses in the batch', async () => {
    const { host, calls } = drawing();
    const slide = handle(ok(host, [slideAt(0)])[0]);
    for (const width of [0, -1, Number.NaN, Number.POSITIVE_INFINITY])
      expect(refusal(host, [{ op: 'prepareImage', target: slide, width }])).toMatchObject({
        error: { code: 'invalid-argument', detail: 'image-size' },
      });
    const answer = await host.renderImage!({ target: slide, height: -5 });
    expect(answer).toMatchObject({ ok: false, error: { detail: 'image-size' } });
    expect(calls).toEqual([]);
  });

  test('a null object refuses, as a write on it does', () => {
    const { host } = drawing();
    expect(
      refusal(host, [
        { op: 'getSlide', key: { id: '999' }, orNull: true },
        { op: 'prepareImage', target: { result: 0 } },
      ])
    ).toMatchObject({ error: { code: 'invalid-handle', detail: 'null-object' } });
  });

  test('a slide removed before the drawing is gone', async () => {
    const { host, calls } = drawing();
    const slide = handle(ok(host, [slideAt(1)])[0]);
    ok(host, [{ op: 'prepareImage', target: slide }]);
    ok(host, [{ op: 'deleteSlide', slide }]);
    const answer = await host.renderImage!({ target: slide });
    expect(answer).toMatchObject({ ok: false, error: { code: 'invalid-handle', detail: 'gone' } });
    expect(calls).toEqual([]);
  });

  test('a forged handle is refused', async () => {
    const { host } = drawing();
    const answer = await host.renderImage!({
      target: { kind: 'slide', ref: 'forged' } as unknown as AutomationHandle,
    });
    expect(answer).toMatchObject({ ok: false, error: { code: 'invalid-handle' } });
  });

  test('a renderer refusal is a typed failure with no engine text', async () => {
    const { host } = drawing(() => Promise.reject(new CommandRefusal('errors.exportFailed', 'x')));
    const slide = handle(ok(host, [slideAt(0)])[0]);
    const answer = await host.renderImage!({ target: slide });
    expect(answer).toEqual({
      ok: false,
      error: {
        code: 'transaction-refused',
        message: 'the picture could not be drawn',
        detail: 'image-failed',
      },
    });
  });

  test('a disposed host draws nothing', async () => {
    const { host, calls } = drawing();
    const slide = handle(ok(host, [slideAt(0)])[0]);
    host.dispose();
    expect(await host.renderImage!({ target: slide })).toMatchObject({
      ok: false,
      error: { code: 'disposed' },
    });
    expect(calls).toEqual([]);
  });

  test('the server host draws no pictures and says so', async () => {
    const opened = createServerAutomationHost(deck());
    if (!opened.ok) throw new Error('fixture');
    const host = opened.host;
    expect(host.capabilities.images).toBe(false);
    const slide = handle(ok(host, [slideAt(0)])[0]);
    expect(refusal(host, [{ op: 'prepareImage', target: slide }])).toMatchObject({
      error: { code: 'unsupported-capability', detail: 'images' },
    });
    expect(await host.renderImage!({ target: slide })).toMatchObject({
      ok: false,
      error: { code: 'unsupported-capability', detail: 'images' },
    });
  });

  test('a server host given a renderImage option claims images and draws through it', async () => {
    const calls: AutomationImageTarget[] = [];
    const opened = createServerAutomationHost(deck(), {
      renderImage: async (target) => {
        calls.push(target);
        return PNG;
      },
    });
    if (!opened.ok) throw new Error('fixture');
    const host = opened.host;
    expect(host.capabilities.images).toBe(true);
    const slide = handle(ok(host, [slideAt(1)])[0]);
    ok(host, [{ op: 'prepareImage', target: slide, width: 64 }]);
    expect(await host.renderImage!({ target: slide, width: 64 })).toEqual({
      ok: true,
      base64: btoa(String.fromCharCode(...PNG)),
    });
    expect(calls).toEqual([{ slideIndex: 1, width: 64 }]);
  });
});
