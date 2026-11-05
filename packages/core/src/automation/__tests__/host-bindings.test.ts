// Bindings: named links from an automation client to shapes, kept by the host in memory and
// never written to the package.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx } from '../../store/index.ts';
import {
  createStoreAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';
import { MAX_BINDINGS } from '../plan-bindings.ts';

const IN = 914_400;

function setup(): {
  host: AutomationHost;
  store: PptxDocumentStore;
  presentation: AutomationHandle;
  box: AutomationHandle;
  other: AutomationHandle;
} {
  const read = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            shape({ id: 2, name: 'Box', x: IN, y: IN, cx: IN, cy: IN, text: para(run('Hi')) }) +
            shape({ id: 3, name: 'Other', x: 3 * IN, y: IN, cx: IN, cy: IN }),
        },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  const host = createStoreAutomationHost(store);
  const [presentation, , box, other] = ok(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 1 }, key: { id: '2' } },
    { op: 'getShape', container: { result: 1 }, key: { id: '3' } },
  ]).map(handleOf);
  return { host, store, presentation: presentation!, box: box!, other: other! };
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function handleOf(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

function ids(host: AutomationHost, owner: AutomationHandle): string[] {
  const [value] = ok(host, [{ op: 'listBindings', owner }]);
  if (value?.kind !== 'bindings') throw new Error('not bindings');
  return value.bindings.map((b) => `${b.id}:${b.type}`);
}

function refusal(host: AutomationHost, operations: AutomationOperation[]): unknown {
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  return response.results.find((r) => r.status === 'error');
}

describe('automation host: bindings', () => {
  test('adds, lists, finds the shape of, and deletes a binding, with no package change', () => {
    const { host, store, presentation, box, other } = setup();
    const revision = store.revision;
    const [, found] = ok(host, [
      { op: 'addBinding', owner: presentation, shape: box, id: 'title' },
      { op: 'getBindingShape', owner: presentation, id: 'title' },
    ]);
    expect(handleOf(found)).toEqual(box);
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'note' }]);
    expect(ids(host, presentation)).toEqual(['title:shape', 'note:shape']);
    // A binding with the same id moves to the new shape and keeps its place.
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'title' }]);
    expect(
      handleOf(ok(host, [{ op: 'getBindingShape', owner: presentation, id: 'title' }])[0])
    ).toEqual(other);
    expect(ids(host, presentation)).toEqual(['title:shape', 'note:shape']);
    ok(host, [{ op: 'deleteBinding', owner: presentation, id: 'note' }]);
    expect(ids(host, presentation)).toEqual(['title:shape']);
    // Bindings are not file content: the store never moved.
    expect(store.revision).toBe(revision);
  });

  test('a binding whose shape is deleted stays listed and refuses its shape', () => {
    const { host, presentation, box } = setup();
    ok(host, [
      { op: 'addBinding', owner: presentation, shape: box, id: 'title' },
      { op: 'deleteShape', shape: box },
    ]);
    expect(ids(host, presentation)).toEqual(['title:shape']);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'title' }])
    ).toMatchObject({
      error: { code: 'item-not-found', detail: 'binding-shape' },
    });
    ok(host, [{ op: 'deleteBinding', owner: presentation, id: 'title' }]);
    expect(ids(host, presentation)).toEqual([]);
  });

  test('a new shape that takes the id of a deleted bound shape does not bind', () => {
    const { host, presentation, other } = setup();
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'note' }]);
    ok(host, [{ op: 'deleteShape', shape: other }]);
    // The next shape takes the highest id plus one: 3 again.
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const [added] = ok(host, [
      { op: 'addShape', container: handleOf(slide), shape: { kind: 'textBox', text: 'New' } },
    ]);
    expect(handleOf(added)).toEqual(other);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'note' }])
    ).toMatchObject({ error: { code: 'item-not-found', detail: 'binding-shape' } });
  });

  test('a delete and an add in one batch do not rebind', () => {
    const { host, presentation, other } = setup();
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'note' }]);
    const response = host.execute({
      operations: [
        { op: 'deleteShape', shape: other },
        { op: 'getSlide', key: { index: 0 } },
        { op: 'addShape', container: { result: 1 }, shape: { kind: 'textBox', text: 'New' } },
        { op: 'getBindingShape', owner: presentation, id: 'note' },
      ],
    });
    expect(response.results[3]).toMatchObject({ error: { code: 'item-not-found' } });
  });

  test('an undo that restores the bound shape makes the binding work again', () => {
    const { host, store, presentation, other } = setup();
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'note' }]);
    ok(host, [{ op: 'deleteShape', shape: other }]);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'note' }])
    ).toBeTruthy();
    expect(store.undo().ok).toBe(true);
    const [found] = ok(host, [{ op: 'getBindingShape', owner: presentation, id: 'note' }]);
    expect(handleOf(found)).toEqual(other);
    // A redo removes it again, and a new shape in its place still does not bind.
    expect(store.redo().ok).toBe(true);
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    ok(host, [
      { op: 'addShape', container: handleOf(slide), shape: { kind: 'textBox', text: 'N' } },
    ]);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'note' }])
    ).toBeTruthy();
  });

  test('a delete by another writer to the store orphans the binding', () => {
    const { host, store, presentation, other } = setup();
    ok(host, [{ op: 'addBinding', owner: presentation, shape: other, id: 'note' }]);
    // A second host stands for an edit in the editor: it commits to the same store.
    const editor = createStoreAutomationHost(store);
    const [slide, shape] = ok(editor, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { id: '3' } },
    ]).map(handleOf);
    ok(editor, [{ op: 'deleteShape', shape: shape! }]);
    ok(editor, [{ op: 'addShape', container: slide!, shape: { kind: 'textBox', text: 'New' } }]);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'note' }])
    ).toMatchObject({ error: { code: 'item-not-found', detail: 'binding-shape' } });
  });

  test('a new last slide that takes the id of a deleted one does not rebind its shapes', () => {
    const { host, presentation } = setup();
    const [slide, box] = ok(host, [
      { op: 'addSlide' },
      { op: 'addShape', container: { result: 0 }, shape: { kind: 'textBox', text: 'Bound' } },
    ]).map(handleOf);
    ok(host, [{ op: 'addBinding', owner: presentation, shape: box!, id: 'last' }]);
    ok(host, [{ op: 'deleteSlide', slide: slide! }]);
    const [again] = ok(host, [
      { op: 'addSlide' },
      { op: 'addShape', container: { result: 0 }, shape: { kind: 'textBox', text: 'Fresh' } },
    ]).map(handleOf);
    expect(again).toEqual(slide);
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'last' }])
    ).toMatchObject({ error: { code: 'item-not-found', detail: 'binding-shape' } });
  });

  test('a refused batch changes no binding', () => {
    const { host, presentation, box } = setup();
    refusal(host, [
      { op: 'addBinding', owner: presentation, shape: box, id: 'title' },
      { op: 'setText', shape: box, text: 'x', span: { start: 99, length: 1 } },
    ]);
    expect(ids(host, presentation)).toEqual([]);
  });

  test('refuses an unknown id, a bad id, and a selection on a host with none', () => {
    const { host, presentation, box } = setup();
    expect(refusal(host, [{ op: 'deleteBinding', owner: presentation, id: 'none' }])).toMatchObject(
      {
        error: { code: 'item-not-found' },
      }
    );
    expect(
      refusal(host, [{ op: 'getBindingShape', owner: presentation, id: 'none' }])
    ).toMatchObject({
      error: { code: 'item-not-found' },
    });
    expect(
      refusal(host, [{ op: 'addBinding', owner: presentation, shape: box, id: '' }])
    ).toMatchObject({ error: { code: 'invalid-argument' } });
    expect(refusal(host, [{ op: 'addBinding', owner: box, shape: box, id: 'x' }])).toMatchObject({
      error: { code: 'invalid-handle' },
    });
    expect(
      refusal(host, [{ op: 'addBindingFromSelection', owner: presentation, id: 'x' }])
    ).toMatchObject({ error: { code: 'unsupported-capability', detail: 'selection' } });
  });

  test('caps the number of bindings', () => {
    const { host, presentation, box } = setup();
    const many: AutomationOperation[] = [];
    for (let i = 0; i < MAX_BINDINGS; i += 1)
      many.push({ op: 'addBinding', owner: presentation, shape: box, id: `b${i}` });
    ok(host, many);
    expect(
      refusal(host, [{ op: 'addBinding', owner: presentation, shape: box, id: 'one-more' }])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'binding-limit' } });
    // Moving a binding that exists is still allowed.
    ok(host, [{ op: 'addBinding', owner: presentation, shape: box, id: 'b0' }]);
  });

  test('a disposed host forgets its bindings', () => {
    const { host, presentation, box } = setup();
    ok(host, [{ op: 'addBinding', owner: presentation, shape: box, id: 'title' }]);
    host.dispose();
    expect(host.execute({ operations: [{ op: 'listBindings', owner: presentation }] }).ok).toBe(
      false
    );
  });
});
