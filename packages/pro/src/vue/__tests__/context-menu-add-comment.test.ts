/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The twin of `../../react/__tests__/context-menu-add-comment.test.tsx`.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component, type VNode } from 'vue';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorSidePanel,
  PptxEditorViewport,
} from '@pptx-editor.dev/vue';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxEditorReview, reviewModule } from '../index.ts';

const IN = 914_400;
const PX = 9_525;

let app: App | null = null;
afterEach(() => {
  app?.unmount();
  app = null;
  document.body.replaceChildren();
});

const settle = async (): Promise<void> => {
  for (let i = 0; i < 3; i += 1) await nextTick();
};
const frame = (): Promise<void> => new Promise((done) => requestAnimationFrame(() => done()));

async function mount(
  modules: readonly EditorModule[] | undefined,
  pane: (() => VNode) | null
): Promise<PptxEditorInstance> {
  app?.unmount();
  document.body.replaceChildren();
  const host = document.createElement('div');
  document.body.append(host);
  let ready: PptxEditorInstance | null = null;
  app = createApp(
    defineComponent({
      setup: () => () =>
        h(
          PptxEditorRoot,
          { onReady: (e: PptxEditorInstance) => (ready = e), author: 'Ada Lovelace', modules },
          () => h(PptxEditorViewport, null, () => [h(PptxEditorContent), pane?.()])
        ),
    }) as Component
  );
  app.mount(host);
  await nextTick();
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: IN, y: IN, cx: 2 * IN, cy: IN }) }],
  });
  const editor = ready as PptxEditorInstance | null;
  expect((await editor!.load(deck)).ok).toBe(true);
  await settle();
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  // happy-dom lays nothing out: one CSS pixel is one screen pixel at zoom 1.
  stage.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 1280,
      bottom: 720,
      width: 1280,
      height: 720,
      x: 0,
      y: 0,
    }) as DOMRect;
  editor!.setZoom(1);
  await settle();
  return editor!;
}

/** Right-click the shape "Title", which selects it and opens the shape menu. */
async function rightClickTitle(): Promise<void> {
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.dispatchEvent(
    new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: (2 * IN) / PX,
      clientY: (1.5 * IN) / PX,
    })
  );
  await settle();
  expect(document.querySelector('[role="menu"][aria-label="Shape actions"]')).not.toBeNull();
}

const row = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('[role="menu"] [data-slot="review.addComment"]');
const field = (): HTMLTextAreaElement | null =>
  document.querySelector<HTMLTextAreaElement>('textarea[aria-label="Add a comment…"]');

describe('New comment in the canvas context menu (Vue)', () => {
  test('the row shows only while a comments pane is mounted with the review module', async () => {
    await mount([reviewModule()], () => h(PptxEditorReview));
    await rightClickTitle();
    expect(row()).not.toBeNull();
    await mount([reviewModule()], null);
    await rightClickTitle();
    expect(row()).toBeNull();
    // Without the review module the pane reads no comments: no row.
    await mount(undefined, () => h(PptxEditorReview));
    await rightClickTitle();
    expect(row()).toBeNull();
  });

  test('the row opens a focused draft that points at the right-clicked element', async () => {
    const editor = await mount([reviewModule()], () => h(PptxEditorReview));
    await rightClickTitle();
    row()!.click();
    await settle();
    expect(document.querySelector('[role="menu"]')).toBeNull();
    await frame();
    expect(document.activeElement).toBe(field());
    expect(document.body.textContent).toContain('Points at Title');
    // The draft keeps its element when the selection moves on.
    editor.setSelection({ kind: 'none' });
    await settle();
    const el = field()!;
    el.value = 'Check the title';
    el.dispatchEvent(new Event('input', { bubbles: true }));
    await settle();
    el.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    );
    await settle();
    const [thread] = editor.snapshot().comments!;
    expect(thread).toMatchObject({ text: 'Check the title', anchor: { kind: 'shape' } });
  });

  test('a closed comments pane opens on its tab with the draft', async () => {
    await mount([reviewModule()], () =>
      h(PptxEditorSidePanel, null, { comments: () => h(PptxEditorReview) })
    );
    expect(document.querySelector('[data-tab="format"]')?.getAttribute('aria-selected')).toBe(
      'true'
    );
    await rightClickTitle();
    row()!.click();
    await settle();
    await frame();
    await settle();
    expect(document.querySelector('[data-tab="comments"]')?.getAttribute('aria-selected')).toBe(
      'true'
    );
    expect(field()).not.toBeNull();
    expect(document.activeElement).toBe(field());
  });

  test('in view mode the row stays, off with the viewing reason', async () => {
    const editor = await mount([reviewModule()], () => h(PptxEditorReview));
    editor.setMode('view');
    await settle();
    await rightClickTitle();
    expect(row()?.getAttribute('aria-disabled')).toBe('true');
    row()!.click();
    await settle();
    expect(field()).toBeNull();
  });
});
