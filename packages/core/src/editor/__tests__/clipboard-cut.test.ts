// A cut deletes only what its copy carried. A chart whose parts cannot travel on the clipboard
// (here, a chart part with an external relationship) is never lost through a cut: the keyboard
// cut event, a cut from the slide list's focus, and the menu's Cut refuse the whole cut, and the
// clipboard and the last copy stay as they were.

import { afterAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { GlobalWindow } from 'happy-dom';
import { runClipboardSlot } from '../chrome-actions.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const CHART_RELS = 'ppt/charts/_rels/chart1.xml.rels';
const PACKAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

/** The charts fixture, with a chart part that links an external workbook. */
function deck(): Uint8Array {
  const root = join(import.meta.dir, '../../../../../fixtures/charts.pptx');
  const files = unzipSync(new Uint8Array(readFileSync(root)));
  files[CHART_RELS] = new TextEncoder().encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdBook" Type="${PACKAGE}" Target="file:///book.xlsx" TargetMode="External"/>` +
      '</Relationships>'
  );
  return zipSync(files);
}

async function open(bytes = deck()): Promise<{ editor: PptxEditorInstance; host: HTMLElement }> {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  editor.attach(host);
  return { editor, host };
}

const names = (e: PptxEditorInstance): string[] =>
  (e.snapshot().shapes ?? []).map((s) => s.name).sort();
const idOf = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;

/** A clipboard event's data, as a browser hands it to the handler. */
function clipboardData() {
  const data = new Map<string, string>();
  return {
    data,
    transfer: {
      setData: (type: string, value: string) => void data.set(type, value),
      getData: (type: string) => data.get(type) ?? '',
      items: [],
      files: [],
      types: [],
    } as unknown as DataTransfer,
  };
}

/** Fire a native `cut` at the slide surface, as Ctrl/Cmd+X does. True when it was handled. */
function nativeCut(host: HTMLElement): { handled: boolean; data: Map<string, string> } {
  const { data, transfer } = clipboardData();
  const event = new win.Event('cut', { bubbles: true, cancelable: true }) as unknown as Event;
  Object.defineProperty(event, 'clipboardData', { value: transfer });
  const input = host.querySelector('.pptx-surface__input') as HTMLElement;
  input.dispatchEvent(event);
  return { handled: event.defaultPrevented, data };
}

describe('a keyboard cut', () => {
  test('of a shape and a chart that cannot travel is refused, and both stay', async () => {
    const { editor, host } = await open();
    editor.setSelection({
      kind: 'shapes',
      nodeIds: [idOf(editor, 'Shape 2'), idOf(editor, 'Chart 1')],
    });
    const before = editor.lastCopied;
    const revision = editor.session!.store.revision;

    const cut = nativeCut(host);

    expect(cut.handled).toBe(true);
    expect(cut.data.size).toBe(0);
    expect(editor.lastCopied).toBe(before);
    expect(names(editor)).toEqual(['Chart 1', 'Shape 2']);
    expect(editor.session!.store.revision).toBe(revision);
    expect(editor.snapshot().lastRejection?.reason).toBe('contextMenu.cannotCopyFrame');
    editor.destroy();
  });

  test('of a group that holds such a chart is refused', async () => {
    const { editor, host } = await open();
    editor.setSelection({
      kind: 'shapes',
      nodeIds: [idOf(editor, 'Shape 2'), idOf(editor, 'Chart 1')],
    });
    expect(editor.exec({ type: 'group' }).ok).toBe(true);
    const group = editor.snapshot().selection;
    expect(group.kind === 'shapes' && group.nodeIds.length).toBe(1);

    expect(nativeCut(host).data.size).toBe(0);
    expect(editor.snapshot().shapes).toHaveLength(1);
    expect(editor.snapshot().lastRejection?.reason).toBe('contextMenu.cannotCopyFrame');
    editor.destroy();
  });

  test('of shapes that travel copies them, deletes them, and undoes', async () => {
    const { editor, host } = await open();
    editor.setSelection({ kind: 'shapes', nodeIds: [idOf(editor, 'Shape 2')] });

    const cut = nativeCut(host);

    expect(cut.handled).toBe(true);
    expect(cut.data.has('application/x-pptx-editor-shapes+json') || cut.data.size > 0).toBe(true);
    expect(editor.lastCopied?.shapes).toHaveLength(1);
    expect(names(editor)).toEqual(['Chart 1']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(names(editor)).toEqual(['Chart 1', 'Shape 2']);
    editor.destroy();
  });
});

describe('the menu Cut', () => {
  test('of a shape and a chart that cannot travel deletes nothing and says why', async () => {
    const { editor, host } = await open();
    editor.setSelection({
      kind: 'shapes',
      nodeIds: [idOf(editor, 'Shape 2'), idOf(editor, 'Chart 1')],
    });
    runClipboardSlot(editor, 'edit.cut', host.querySelector('.pptx-surface__input'));
    expect(names(editor)).toEqual(['Chart 1', 'Shape 2']);
    expect(editor.snapshot().lastRejection?.reason).toBe('contextMenu.cannotCopyFrame');
    editor.destroy();
  });
});
