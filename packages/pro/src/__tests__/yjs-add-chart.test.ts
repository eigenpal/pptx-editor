// The automation `addShape` operation with a chart kind, which `ShapeCollection.addChart` of the
// object model sends, through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration
// verification"). The write adds a chart part (`c:chartSpace`, `CT_ChartSpace`, ECMA-376 Part 1,
// §21.2.2.29), its embedded workbook (`c:externalData`, `CT_ExternalData`, §21.2.2.63), and a
// graphic frame (`p:graphicFrame`, `CT_GraphicalObjectFrame`, §19.3.1.21); both replicas must
// converge, also when both peers add a chart, or one edits the same slide.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type {
  AutomationHandle,
  AutomationHost,
  AutomationOperation,
} from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
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

const IN = 914_400;

/** One slide with a title box. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'Title',
          x: IN,
          y: IN / 4,
          cx: 6 * IN,
          cy: IN,
          text: para(run('Report')),
        }),
      },
    ],
  });

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

function exec(host: AutomationHost, operations: AutomationOperation[]) {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results;
}

/**
 * Add a chart with the operations `addChart` sends: the lookups, then the add.
 * `packages/editor-api/src/__tests__/add-chart.test.ts` checks that it sends these.
 */
function addChart(editor: PptxEditorInstance, name: string, left: number): void {
  exec(createBrowserAutomationHost(editor), [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index: 0 } },
    {
      op: 'addShape',
      container: { result: 1 },
      shape: {
        kind: 'chart',
        chartType: 'columnClustered',
        data: { categories: ['Q1', 'Q2'], series: [{ name, values: [1, 2] }] },
      },
      box: { left, top: 100, width: 300, height: 200 },
    },
  ]);
}

/** The chart names of the first slide, by series name, in z-order. */
function chartSeries(editor: PptxEditorInstance): string[] {
  const host = createBrowserAutomationHost(editor);
  const [, listed] = exec(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'listShapes', container: { result: 0 } },
  ]);
  const handles =
    listed!.status === 'ok' && listed!.value.kind === 'handles' ? listed!.value.handles : [];
  const names: string[] = [];
  for (const handle of handles as AutomationHandle[]) {
    const [read] = exec(host, [{ op: 'readChart', shape: handle, orNull: true }]);
    if (read!.status === 'ok' && read!.value.kind === 'chart')
      names.push(read!.value.chart.series.map((s) => s.name).join(','));
  }
  return names;
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('ShapeCollection.addChart on two peers', () => {
  test('a chart added on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    addChart(b, 'Revenue', 60);
    expect(chartSeries(a)).toEqual(['Revenue']);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const reopened = [a, b].map((e) => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen');
      return canonicalPackageFingerprint(r.pkg);
    });
    expect(reopened[0]).toBe(reopened[1]!);
  });

  test('charts added on both peers at once keep their own parts', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    addChart(a, 'From A', 60);
    addChart(b, 'From B', 400);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect([...chartSeries(a)].sort()).toEqual(['From A', 'From B']);
  });

  test('a chart added while a peer edits the same slide keeps both changes', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    addChart(a, 'Revenue', 60);
    const host = createBrowserAutomationHost(b);
    exec(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
      { op: 'setText', shape: { result: 1 }, text: 'Quarterly report' },
    ]);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(chartSeries(b)).toEqual(['Revenue']);
  });

  test('undo of the add on one peer removes the chart on both', async () => {
    const { a, b } = await pair();
    addChart(b, 'Revenue', 60);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(chartSeries(a)).toEqual([]);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
