// The automation `setChartData` operation, which `Chart.setData` of the object model sends,
// through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The
// write is the chart part (`c:chartSpace`, ECMA-376 Part 1, §21.2.2.29) and its embedded
// workbook (`c:externalData`, §21.2.2.63), each written again whole; both replicas must
// converge, also while a peer edits the same slide.

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
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { chartDataDeck } from '../../../../scripts/fixtures/chart-data.ts';
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
  expect((await editor.load(buildPptx(chartDataDeck()))).ok).toBe(true);
  return editor;
}

function run(host: AutomationHost, operations: AutomationOperation[]) {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results;
}

/** The handle of shape `id` on the first slide. */
function shapeHandle(host: AutomationHost, id: string): AutomationHandle {
  const results = run(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { id } },
  ]);
  const value = results[1]!.status === 'ok' ? results[1]!.value : undefined;
  if (value?.kind !== 'handle') throw new Error('no shape');
  return value.handle;
}

/**
 * Replace the chart data with the operation `Chart.setData` sends, after the lookups.
 * `packages/editor-api/src/__tests__/chart-data.test.ts` checks that it sends these.
 */
function setData(editor: PptxEditorInstance, values: number[], extra = false): void {
  const host = createBrowserAutomationHost(editor);
  const categories = ['North', 'South', 'East', 'West', ...(extra ? ['Central'] : [])];
  run(host, [
    {
      op: 'setChartData',
      shape: shapeHandle(host, '3'),
      data: {
        categories,
        series: [
          { name: 'Plan', values: categories.map((_, i) => values[i] ?? 0) },
          { name: 'Actual', values: categories.map((_, i) => (values[i] ?? 0) + 1) },
        ],
      },
    },
  ]);
}

/** Type into the title on the same slide. */
function retitle(editor: PptxEditorInstance, text: string): void {
  const host = createBrowserAutomationHost(editor);
  run(host, [{ op: 'setText', shape: shapeHandle(host, '2'), text }]);
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

function chartValues(e: PptxEditorInstance): unknown {
  const host = createBrowserAutomationHost(e);
  const results = run(host, [{ op: 'readChart', shape: shapeHandle(host, '3') }]);
  const value = results[0]!.status === 'ok' ? results[0]!.value : undefined;
  return value?.kind === 'chart' ? value.chart : undefined;
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('Chart.setData on two peers', () => {
  test('a change on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    setData(b, [1, 2, 3, 4], true);
    expect(chartValues(a)).toEqual(chartValues(b));
    expect(chartValues(a)).toMatchObject({
      categories: ['North', 'South', 'East', 'West', 'Central'],
    });
    expect(fingerprint(a)).toBe(fingerprint(b));
    const reopened = [a, b].map((e) => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen');
      return canonicalPackageFingerprint(r.pkg);
    });
    expect(reopened[0]).toBe(reopened[1]!);
  });

  test('a change while a peer edits the same slide keeps both changes', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setData(a, [5, 6, 7, 8]);
    retitle(b, 'Quarterly report');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(chartValues(b)).toMatchObject({
      series: [
        { name: 'Plan', values: [5, 6, 7, 8] },
        { name: 'Actual', values: [6, 7, 8, 9] },
      ],
    });
    const host = createBrowserAutomationHost(a);
    const text = run(host, [{ op: 'readText', shape: shapeHandle(host, '2') }])[0]!;
    expect(text).toMatchObject({ status: 'ok', value: { kind: 'text', text: 'Quarterly report' } });
  });

  test('concurrent changes to one chart converge to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setData(a, [1, 1, 1, 1]);
    setData(b, [2, 2, 2, 2], true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(chartValues(a)).toEqual(chartValues(b));
  });

  test('undo of the change on one peer restores the chart on both', async () => {
    const { a, b } = await pair();
    const before = chartValues(a);
    setData(b, [9, 9, 9, 9], true);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(chartValues(a)).toEqual(before);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
