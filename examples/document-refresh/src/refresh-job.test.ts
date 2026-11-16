// `runRefreshJob()` against a headless editor, after DOCX Editor's
// examples/document-refresh/src/refresh-job.test.ts: identity checks, cumulative results, the
// out-of-order refusal, the refusal stop, cancellation, and the `finally` that finishes the
// submission. The suite runs this file with no global DOM.

import { afterEach, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import {
  createDocumentRefresh,
  createPptxEditor,
  type EditorModule,
  type RefreshResult,
} from '@pptx-editor.dev/core/editor';
import { reviewModule } from '@pptx-editor.dev/pro';
import { samplePptx, sampleUpdate } from '../sample-presentation.ts';
import { runRefreshJob, type ServerUpdate } from './refresh-job.ts';

const close: (() => void)[] = [];
afterEach(() => close.splice(0).forEach((dispose) => dispose()));

async function open(modules: EditorModule[] = []) {
  const editor = createPptxEditor({ fixedMeasurer: true, modules });
  expect((await editor.load(samplePptx())).ok).toBe(true);
  const refresh = createDocumentRefresh(editor);
  close.push(() => editor.destroy());
  return { editor, refresh };
}

/** Every text run of the saved slides and notes, in one string. */
function savedText(editor: ReturnType<typeof createPptxEditor>): string {
  const files = unzipSync(editor.save()!);
  return Object.keys(files)
    .filter((name) => /^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(name))
    .flatMap((name) =>
      [...strFromU8(files[name]!).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1])
    )
    .join('\n');
}

const never = new AbortController().signal;

test('successive jobs apply cumulative files, summaries, and partial failures; late results cannot overwrite', async () => {
  const { editor, refresh } = await open();
  const results: RefreshResult[] = [];
  for (const round of [1, 2]) {
    await runRefreshJob(
      refresh,
      never,
      async function* (submission) {
        yield sampleUpdate(submission.id, round, 1);
        yield sampleUpdate(submission.id, round, 2);
        yield sampleUpdate(submission.id, round, 1);
      },
      (result) => results.push(result)
    );
    expect(refresh.snapshot().phase).toBe('complete');
    const text = savedText(editor);
    expect(text).toContain(`Revenue reached $4.${round + 1}M this quarter.`);
    expect(text).toContain(`Mention the launch on March ${round + 1}.`);
    expect(text).toContain('Summary');
    expect(text).not.toContain('DRAFT');
    expect(editor.snapshot().slides.length).toBe(4);
    expect(results.at(-1)).toMatchObject({ ok: false, code: 'out-of-order' });
    const accepted = results.at(-2)!;
    expect(accepted.ok).toBe(true);
    if (!accepted.ok) continue;
    expect(accepted.mode).toBe('reload');
    expect(accepted.failures).toEqual(['regional-chart']);
    expect(accepted.changes.map((change) => [change.id, change.status])).toEqual([
      ['revenue', 'available'],
      ['summary-slide', 'available'],
      ['launch-notes', 'available'],
      ['draft-label', 'available'],
    ]);
    expect(refresh.highlightChanges({ includePrevious: true, animation: false })).toBe(4);
  }
});

test('undoable updates add one undo step per accepted result', async () => {
  const { editor, refresh } = await open();
  const original = savedText(editor);
  const results: RefreshResult[] = [];
  await runRefreshJob(
    refresh,
    never,
    async function* (submission) {
      yield sampleUpdate(submission.id, 1, 1);
      yield sampleUpdate(submission.id, 1, 2);
    },
    (result) => results.push(result),
    { mode: 'edit' }
  );
  expect(results.map((result) => result.ok && result.mode)).toEqual(['edit', 'edit']);
  const second = savedText(editor);
  expect(second).toContain('Mention the launch on March 2.');
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(savedText(editor)).not.toBe(second);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(savedText(editor)).toBe(original);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(savedText(editor)).toBe(second);
});

test('without processor changes, the review module computes them', async () => {
  const { refresh } = await open([reviewModule()]);
  const results: RefreshResult[] = [];
  await runRefreshJob(
    refresh,
    never,
    async function* (submission) {
      yield sampleUpdate(submission.id, 1, 1);
      yield sampleUpdate(submission.id, 1, 2);
    },
    (result) => results.push(result),
    { computeChanges: true }
  );
  const last = results.at(-1)!;
  expect(last.ok).toBe(true);
  if (!last.ok) return;
  expect(last.changeInformation).toBe('available');
  expect(last.failures).toEqual(['regional-chart']);
  expect(last.changes.every((change) => change.computed && !change.description)).toBe(true);
  expect(last.changes.map((change) => [change.computed!.scope, change.computed!.action])).toEqual(
    expect.arrayContaining([
      ['text', 'changed'],
      ['slide', 'added'],
      ['notes', 'changed'],
      ['shape', 'removed'],
    ])
  );
});

test('local edits stop a job and survive the refused replacement', async () => {
  const { editor, refresh } = await open();
  const results: RefreshResult[] = [];
  await runRefreshJob(
    refresh,
    never,
    async function* (submission) {
      expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
      yield sampleUpdate(submission.id, 1, 1);
      throw new Error('The client must stop receiving after refusal');
    },
    (result) => results.push(result)
  );
  expect(results).toMatchObject([{ ok: false, code: 'local-edits' }]);
  expect(editor.snapshot().slides.length).toBe(4);
  expect(savedText(editor)).toContain('Revenue reached $4.0M this quarter.');
});

test('cancellation keeps accepted output and prevents later updates', async () => {
  const { editor, refresh } = await open();
  const controller = new AbortController();
  const results: RefreshResult[] = [];
  await runRefreshJob(
    refresh,
    controller.signal,
    async function* (submission) {
      yield sampleUpdate(submission.id, 1, 1);
      controller.abort();
      refresh.cancel();
      yield sampleUpdate(submission.id, 1, 2);
    },
    (result) => results.push(result)
  );
  expect(results).toHaveLength(1);
  const text = savedText(editor);
  expect(text).toContain('Revenue reached $4.2M this quarter.');
  expect(text).toContain('DRAFT');
  expect(text).not.toContain('Mention the launch on March 2.');
});

for (const field of ['documentId', 'submissionId'] as const) {
  test(`rejects a foreign ${field} and closes the submission`, async () => {
    const { editor, refresh } = await open();
    const original = savedText(editor);
    await expect(
      runRefreshJob(
        refresh,
        never,
        async function* (submission) {
          yield { ...sampleUpdate(submission.id, 1, 1), [field]: 'foreign' };
        },
        () => {
          throw new Error('Foreign output must not reach the editor');
        }
      )
    ).rejects.toThrow('identity-mismatch');
    expect(savedText(editor)).toBe(original);
    expect(refresh.snapshot().phase).toBe('idle');
  });
}

test('transport failure closes the submission and permits a fresh request', async () => {
  const { refresh } = await open();
  await expect(
    runRefreshJob(
      refresh,
      never,
      async function* (): AsyncGenerator<ServerUpdate> {
        throw new Error('transport-failed');
      },
      () => {}
    )
  ).rejects.toThrow('transport-failed');
  expect(refresh.snapshot().phase).toBe('idle');
  const submission = await refresh.capture();
  refresh.finish(submission);
});
