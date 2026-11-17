// The controls of the agent replay page in happy-dom, with short timings and the scripted files
// in place of the mock server: Next step, Play all, Pause and Play all again, Undo step and its
// refusals, Restart, and the buttons that do nothing during Play all.

import { unregisterDom } from './test-dom.ts';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { reviewModule } from '@pptx-editor.dev/pro';
import { PptxEditor } from '@pptx-editor.dev/react';
import { AGENT_STEPS, agentPptx } from '../agent-script.ts';
import { AgentPanel, UNDO_REFUSED, type AgentTiming } from './agent-panel.tsx';

afterEach(cleanup);
afterAll(unregisterDom);

const MODULES = [reviewModule()];
const LAST = AGENT_STEPS.length;
// Bun loads the testing library before `test-dom.ts` registers the DOM, so its `screen` has no
// document: query the body at call time.
const screen = {
  getByRole: (role: string, options?: { name?: string | RegExp }) =>
    within(document.body).getByRole(role, options),
};
const sleep = (ms: number) => act(() => new Promise<void>((r) => setTimeout(r, ms)));

/** The page with `timing`; `loads` lists each requested step with the time of the request. */
async function mount(timing: Partial<AgentTiming> = {}) {
  const loads: { step: number; at: number }[] = [];
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditor.Root
      document={agentPptx(0)}
      modules={MODULES}
      onReady={(instance: PptxEditorInstance) => (editor = instance)}
    >
      <PptxEditor.Content />
      <AgentPanel
        timing={{ thinkMs: 0, pauseMs: 20, highlightMs: 50, ...timing }}
        loadStep={async (step) => {
          loads.push({ step, at: performance.now() });
          return agentPptx(step);
        }}
      />
    </PptxEditor.Root>
  );
  await waitFor(() => expect(editor?.snapshot().status).toBe('ready'), { timeout: 10_000 });
  const button = (name: string) => screen.getByRole('button', { name });
  const status = () => document.querySelector('.refresh-status')?.textContent ?? '';
  const applied = () => screen.getByRole('heading', { name: /Steps/ }).textContent ?? '';
  const until = (check: () => void, timeout = 10_000) => waitFor(check, { timeout });
  const appliedCount = (n: number) => until(() => expect(applied()).toContain(`${n} of ${LAST}`));
  const click = (name: string) => act(() => void fireEvent.click(button(name)));
  return { editor: editor!, loads, button, status, appliedCount, until, click };
}

describe('agent replay page', () => {
  test('Next step applies one step and lists its change', async () => {
    const { click, appliedCount, status, loads, editor } = await mount();
    await click('Next step');
    await appliedCount(1);
    expect(status()).toContain('Step 1 applied as an undoable edit. 1 change flashed.');
    expect(loads.map((l) => l.step)).toEqual([1]);
    const changes = screen.getByRole('list', { name: 'Changes' });
    expect(changes.querySelectorAll('li')).toHaveLength(1);
    expect(editor.snapshot().canUndo).toBe(true);
  }, 20_000);

  test('Play all runs every step to the end', async () => {
    const { click, appliedCount, button, loads, editor, until } = await mount();
    await click('Play all');
    expect(button('Pause')).toBeDefined();
    await appliedCount(LAST);
    await until(() => expect(button('Play all').getAttribute('aria-disabled')).toBe('true'));
    expect(loads.map((l) => l.step)).toEqual([1, 2, 3, 4, 5, 6, 7]);
    // Step 4 added a slide; step 7 moved one.
    expect(editor.snapshot().slides).toHaveLength(5);
  }, 20_000);

  test('Pause, then Play all during the pause: one loop runs every step', async () => {
    const pauseMs = 250;
    const { click, appliedCount, loads, until, button } = await mount({ pauseMs });
    await click('Play all');
    await appliedCount(1);
    await click('Pause');
    await sleep(50);
    // The first loop still waits in its pause.
    await click('Play all');
    await appliedCount(LAST);
    await until(() => expect(button('Play all')).toBeDefined());
    expect(loads.map((l) => l.step)).toEqual([1, 2, 3, 4, 5, 6, 7]);
    // One loop: after the new Play all, each step waits for the whole pause after the one before.
    const gaps = loads.slice(2).map((l, i) => l.at - loads[i + 1]!.at);
    for (const gap of gaps) expect(gap).toBeGreaterThan(pauseMs * 0.9);
  }, 20_000);

  test('Undo step returns to the previous step, also twice in a row', async () => {
    const { click, appliedCount, status, editor } = await mount();
    const original = editor.save();
    await click('Next step');
    await appliedCount(1);
    await click('Next step');
    await appliedCount(2);
    await click('Undo step');
    await appliedCount(1);
    expect(status()).toContain('Undid step 2.');
    await click('Undo step');
    await appliedCount(0);
    expect(editor.save()).toEqual(original);
    expect(editor.snapshot().canUndo).toBe(false);
  }, 20_000);

  test('Undo step refuses after a user edit and after a reload step', async () => {
    const { click, appliedCount, status, button, editor, until } = await mount();
    await click('Next step');
    await appliedCount(1);
    expect(button('Undo step').getAttribute('aria-disabled')).toBe('false');
    // A user edit after the step: an undo would revert the edit, not the step.
    act(() => void editor.exec({ type: 'insertTextBox' }));
    await until(() => expect(button('Undo step').getAttribute('aria-disabled')).toBe('true'));
    const revision = editor.getDocumentHandle().revision;
    await click('Undo step');
    expect(status()).toBe(UNDO_REFUSED);
    expect(editor.getDocumentHandle().revision).toBe(revision);
    // A step after the edit can undo, but the step before it then refuses: the edit is between.
    await click('Next step');
    await appliedCount(2);
    await click('Undo step');
    await appliedCount(1);
    expect(button('Undo step').getAttribute('aria-disabled')).toBe('true');
    // A reload step records no undo step.
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Reload' })));
    await click('Next step');
    await appliedCount(2);
    expect(status()).toContain('applied as a reload');
    expect(button('Undo step').getAttribute('aria-disabled')).toBe('true');
    await click('Undo step');
    expect(status()).toBe(UNDO_REFUSED);
    await appliedCount(2);
  }, 20_000);

  test('Restart returns to step 0, also during Play all', async () => {
    const { click, appliedCount, status, editor, button, loads, until } = await mount({
      pauseMs: 300,
    });
    await click('Next step');
    await appliedCount(1);
    await click('Restart');
    await until(() => expect(status()).toContain('Restarted.'));
    await appliedCount(0);
    expect(editor.snapshot().currentSlide).toBe(0);
    // During Play all, Restart pauses first: no step runs after it.
    await click('Play all');
    await appliedCount(1);
    await click('Restart');
    await until(() => expect(status()).toContain('Restarted.'));
    expect(button('Play all')).toBeDefined();
    const requested = loads.length;
    await sleep(450);
    expect(loads).toHaveLength(requested);
    await appliedCount(0);
  }, 20_000);

  test('Next step and Undo step do nothing during Play all', async () => {
    const { click, appliedCount, loads, button, editor } = await mount({ pauseMs: 400 });
    await click('Play all');
    await appliedCount(1);
    expect(button('Next step').getAttribute('aria-disabled')).toBe('true');
    expect(button('Undo step').getAttribute('aria-disabled')).toBe('true');
    const revision = editor.getDocumentHandle().revision;
    await click('Next step');
    await click('Undo step');
    expect(loads.map((l) => l.step)).toEqual([1]);
    expect(editor.getDocumentHandle().revision).toBe(revision);
    await appliedCount(1);
    await click('Pause');
  }, 20_000);
});
