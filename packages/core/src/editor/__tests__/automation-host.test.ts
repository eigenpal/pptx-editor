import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import type {
  AutomationHandle,
  AutomationHost,
  AutomationOperation,
  AutomationValue,
} from '../../automation/index.ts';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '../index.ts';
import { isCollaborating, withCollaborationFlag, type EditorModuleRegistry } from '../modules.ts';

const IN = 914_400;

describe('the collaboration flag of an editor', () => {
  test('follows a collaboration module, so automation merges keep their copies in a room', () => {
    const alone = withCollaborationFlag({}, { review: null, collaboration: null });
    const room = withCollaborationFlag({}, {
      review: null,
      collaboration: {},
    } as unknown as EditorModuleRegistry);
    expect([isCollaborating(alone), isCollaborating(room)]).toEqual([false, true]);
    expect(isCollaborating(createPptxEditor({ fixedMeasurer: true }))).toBe(false);
  });
});

const deck = (title = 'Hello'): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            fill: solid('4472C4'),
            text: para(run(title)) + para(run('World')),
          }) + shape({ id: 3, name: 'Other', x: 5 * IN, y: IN, cx: IN, cy: IN }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

async function open(bytes = deck()): Promise<{ editor: PptxEditorInstance; host: AutomationHost }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return { editor, host: createBrowserAutomationHost(editor) };
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

const box = (host: AutomationHost): AutomationHandle => {
  const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
  const [shape] = ok(host, [{ op: 'getShape', container: handle(slide), key: { id: '2' } }]);
  return handle(shape);
};

describe('the browser automation host', () => {
  test('claims selection and no save', async () => {
    const { host } = await open();
    expect(host.capabilities).toMatchObject({ selection: true, save: false, autofit: true });
    // Pictures need a DOM and a canvas: claimed only where a document exists.
    expect(host.capabilities.images).toBe(typeof document !== 'undefined');
  });

  test('without a DOM, a picture request refuses with a typed reason', async () => {
    const { host } = await open();
    if (typeof document !== 'undefined') return;
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const response = host.execute({ operations: [{ op: 'prepareImage', target: handle(slide) }] });
    expect(response.results[0]).toMatchObject({
      status: 'error',
      error: { code: 'unsupported-capability', detail: 'images' },
    });
  });

  test('a batch is one undo step in the editor history', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    ok(host, [
      { op: 'setText', shape, text: 'Changed' },
      { op: 'setFill', shape, fill: { kind: 'solid', color: '#FF0000' } },
    ]);
    expect(ok(host, [{ op: 'readText', shape }])).toEqual([{ kind: 'text', text: 'Changed' }]);
    expect(editor.snapshot().canUndo).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ok(host, [{ op: 'readText', shape }])).toEqual([{ kind: 'text', text: 'Hello\rWorld' }]);
    expect(ok(host, [{ op: 'readFill', shape }])[0]).toMatchObject({
      fill: { foregroundColor: '#4472C4' },
    });
  });

  test('batches of one history group are one undo step; another group starts a new one', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    const text = (): AutomationValue[] => ok(host, [{ op: 'readText', shape }]);
    const run = (value: string, historyGroup?: string): void => {
      const response = host.execute({
        operations: [{ op: 'setText', shape, text: value }],
        ...(historyGroup === undefined ? {} : { historyGroup }),
      });
      expect(response.ok).toBe(true);
    };
    run('One', 'a');
    run('Two', 'a');
    run('Three', 'b');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(text()).toEqual([{ kind: 'text', text: 'Two' }]);
    // Both batches of group `a` go back in one step.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(text()).toEqual([{ kind: 'text', text: 'Hello\rWorld' }]);
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('an editor edit between two batches of one group closes the group', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    expect(
      host.execute({ operations: [{ op: 'setText', shape, text: 'One' }], historyGroup: 'a' }).ok
    ).toBe(true);
    editor.selectSlide(1);
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(
      host.execute({ operations: [{ op: 'setText', shape, text: 'Two' }], historyGroup: 'a' }).ok
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(ok(host, [{ op: 'readText', shape }])).toEqual([{ kind: 'text', text: 'One' }]);
  });

  test('undo of a batch restores the slide and selection from before it', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    ok(host, [{ op: 'selectShapes', shapes: [shape] }]);
    const selected = editor.snapshot().selection;
    expect(selected.kind).toBe('shapes');
    const [slide2] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    // The batch writes, then moves the reader to the second slide.
    ok(host, [
      { op: 'setText', shape, text: 'Changed' },
      { op: 'selectSlides', slides: [handle(slide2)] },
    ]);
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(0);
    expect(editor.snapshot().selection).toEqual(selected);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(1);
  });

  test('editor edits move the revision and notify subscribers', async () => {
    const { editor, host } = await open();
    const seen: number[] = [];
    host.subscribe((e) => seen.push(e.revision));
    const before = host.revision();
    editor.selectSlide(0);
    expect(editor.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
    expect(host.revision()).toBeGreaterThan(before);
    expect(seen.length).toBeGreaterThan(0);
    const stale = host.execute({
      operations: [{ op: 'listSlides' }],
      expectedRevision: before,
    });
    expect(stale.ok).toBe(false);
  });

  test('reads and moves the reader selection', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    const [slide2] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    ok(host, [{ op: 'selectShapes', shapes: [shape] }]);
    const [shapes] = ok(host, [{ op: 'readSelection' }]);
    expect(shapes).toMatchObject({ selection: { shapes: [shape], text: null } });

    ok(host, [{ op: 'selectText', shape, span: { start: 6, length: 3 } }]);
    expect(editor.snapshot().selection).toMatchObject({
      kind: 'text',
      anchor: { paragraph: 1, offset: 0 },
      focus: { paragraph: 1, offset: 3 },
    });
    const [text] = ok(host, [{ op: 'readSelection' }]);
    expect(text).toMatchObject({ selection: { text: { shape, start: 6, length: 3 } } });

    ok(host, [{ op: 'selectSlides', slides: [handle(slide2)] }]);
    expect(editor.snapshot().currentSlide).toBe(1);
    const [slides] = ok(host, [{ op: 'readSelection' }]);
    expect(slides).toMatchObject({ selection: { slides: [handle(slide2)] } });
  });

  test('a selection read sees the selection writes before it in the same batch', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    const [slide2] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    const [, shapes, , text, , slides] = ok(host, [
      { op: 'selectShapes', shapes: [shape] },
      { op: 'readSelection' },
      { op: 'selectText', shape, span: { start: 6, length: 3 } },
      { op: 'readSelection' },
      { op: 'selectSlides', slides: [handle(slide2)] },
      { op: 'readSelection' },
    ]);
    expect(shapes).toMatchObject({ selection: { shapes: [shape], text: null } });
    expect(text).toMatchObject({ selection: { shapes: [shape], text: { shape, start: 6 } } });
    expect(slides).toMatchObject({
      selection: { slides: [handle(slide2)], shapes: [], text: null },
    });
    // The last write wins after the commit, as the last read answered.
    expect(editor.snapshot().currentSlide).toBe(1);
    const [after] = ok(host, [{ op: 'readSelection' }]);
    expect(after).toMatchObject({ selection: { slides: [handle(slide2)] } });
  });

  test('refuses writes while the editor is open for viewing', async () => {
    const { editor, host } = await open();
    const shape = box(host);
    editor.setMode('view');
    const r = host.execute({ operations: [{ op: 'setText', shape, text: 'No' }] });
    expect(r.ok).toBe(false);
    expect(r.results[0]).toMatchObject({
      status: 'error',
      error: { code: 'transaction-refused', detail: 'disabled.viewing' },
    });
    expect(ok(host, [{ op: 'readText', shape }])).toEqual([{ kind: 'text', text: 'Hello\rWorld' }]);
  });

  test('a new document keeps the revision monotonic', async () => {
    const { editor, host } = await open();
    const before = host.revision();
    expect((await editor.load(deck('Again'))).ok).toBe(true);
    expect(host.revision()).toBeGreaterThan(before);
    expect(ok(host, [{ op: 'readText', shape: box(host) }])).toEqual([
      { kind: 'text', text: 'Again\rWorld' },
    ]);
  });

  test('a handle of the previous document does not address the new one', async () => {
    const { editor, host } = await open();
    const old = box(host);
    expect((await editor.load(deck('Again'))).ok).toBe(true);
    const r = host.execute({ operations: [{ op: 'setText', shape: old, text: 'Lost' }] });
    expect(r.ok).toBe(false);
    expect(r.results[0]).toMatchObject({ status: 'error', error: { code: 'invalid-handle' } });
    expect(ok(host, [{ op: 'readText', shape: box(host) }])).toEqual([
      { kind: 'text', text: 'Again\rWorld' },
    ]);
  });

  test('a repeated slide in a slide selection selects it once', async () => {
    const { editor, host } = await open();
    const [first, second] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getSlide', key: { index: 1 } },
    ]);
    ok(host, [{ op: 'selectSlides', slides: [handle(first), handle(second), handle(second)] }]);
    expect([...editor.snapshot().selectedSlides].sort()).toEqual([0, 1]);
  });
});

describe('the browser automation host: bindings', () => {
  const presentation = (host: AutomationHost): AutomationHandle =>
    handle(ok(host, [{ op: 'getPresentation' }])[0]);

  test('binds the one selected shape, and refuses any other selection', async () => {
    const { host } = await open();
    const owner = presentation(host);
    const empty = host.execute({
      operations: [{ op: 'addBindingFromSelection', owner, id: 'picked' }],
    });
    expect(empty.results[0]).toMatchObject({
      status: 'error',
      error: { code: 'invalid-argument', detail: 'binding-selection' },
    });
    const shape = box(host);
    ok(host, [{ op: 'selectShapes', shapes: [shape] }]);
    const [, found] = ok(host, [
      { op: 'addBindingFromSelection', owner, id: 'picked' },
      { op: 'getBindingShape', owner, id: 'picked' },
    ]);
    expect(handle(found)).toEqual(shape);
  });

  test('a binding adds no undo step, and works again when an undo restores its shape', async () => {
    const { editor, host } = await open();
    const owner = presentation(host);
    const shape = box(host);
    const before = editor.snapshot().canUndo;
    ok(host, [{ op: 'addBinding', owner, shape, id: 'title' }]);
    expect(editor.snapshot().canUndo).toBe(before);
    ok(host, [{ op: 'deleteShape', shape }]);
    const gone = host.execute({ operations: [{ op: 'getBindingShape', owner, id: 'title' }] });
    expect(gone.results[0]).toMatchObject({ error: { code: 'item-not-found' } });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(handle(ok(host, [{ op: 'getBindingShape', owner, id: 'title' }])[0])).toEqual(shape);
  });

  test('a new document drops the bindings of the previous one', async () => {
    const { editor, host } = await open();
    ok(host, [{ op: 'addBinding', owner: presentation(host), shape: box(host), id: 'title' }]);
    expect((await editor.load(deck('Again'))).ok).toBe(true);
    expect(ok(host, [{ op: 'listBindings', owner: presentation(host) }])).toEqual([
      { kind: 'bindings', bindings: [] },
    ]);
  });
});
