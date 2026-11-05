// The chrome tooltip in the React adapter: the root mounts the bubble, a hover on a toolbar
// control opens it after the delay inside the editor scope, a disabled control shows its reason
// on its own line and keeps it in `aria-describedby`, and no chrome control keeps a native
// `title`. After DOCX Editor's `Tooltip`.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { TOOLTIP_DELAY_MS, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorRoot, PptxEditorViewport } from '../PptxEditorRoot.tsx';
import { StatusBar } from '../StatusBar.tsx';
import { Toolbar } from '../index.ts';

afterEach(() => cleanup());

const wait = (ms: number) => act(() => new Promise<void>((r) => setTimeout(r, ms)));
const over = (el: Element) =>
  act(() => void el.dispatchEvent(new MouseEvent('pointerover', { bubbles: true })));

function mount(colorMode?: 'dark') {
  let editor: PptxEditorInstance | null = null;
  const view = render(
    <PptxEditorRoot onReady={(e) => (editor = e)} {...(colorMode ? { colorMode } : {})}>
      <PptxEditorViewport>
        <Toolbar />
        <StatusBar />
      </PptxEditorViewport>
    </PptxEditorRoot>
  );
  return { ...view, editor: () => editor };
}

describe('the chrome tooltip', () => {
  test('a hover opens the bubble in the editor scope, with the reason on its own line', async () => {
    const { container } = mount();
    const undo = container.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    over(undo);
    expect(document.querySelectorAll('.pptx-tooltip').length).toBe(0);
    await wait(TOOLTIP_DELAY_MS + 40);
    const bubble = document.querySelector<HTMLElement>('.pptx-tooltip')!;
    expect(bubble.parentElement).toBe(container.querySelector('.pptx-editor'));
    expect(bubble.getAttribute('role')).toBe('tooltip');
    expect(bubble.getAttribute('aria-hidden')).toBe('true');
    expect(bubble.getAttribute('data-side')).toBe('bottom');
    expect(bubble.querySelector('.pptx-tooltip__label')!.textContent).toBe(
      undo.getAttribute('aria-label')!
    );
    const reason = bubble.querySelector('.pptx-tooltip__reason')!.textContent!;
    expect(reason.length).toBeGreaterThan(0);
    expect(bubble.textContent).not.toContain('—');
    // Assistive technology keeps the reason through the control's description.
    expect(document.getElementById(undo.getAttribute('aria-describedby')!)!.textContent).toBe(
      reason
    );
    over(container.querySelector('.pptx-editor')!);
    expect(document.querySelectorAll('.pptx-tooltip').length).toBe(0);
  });

  test('no toolbar or status bar control keeps a native title', () => {
    const { container } = mount();
    expect(container.querySelectorAll('[data-pptx-tooltip]').length).toBeGreaterThan(10);
    expect(
      container.querySelectorAll('.pptx-toolbar [title], .pptx-statusbar [title]').length
    ).toBe(0);
  });

  test('in dark mode the bubble sits in the dark scope', async () => {
    const { container } = mount('dark');
    over(container.querySelector('[data-slot="history.redo"]')!);
    await wait(TOOLTIP_DELAY_MS + 40);
    expect(document.querySelector('.pptx-tooltip')!.closest('.pptx-editor.dark') !== null).toBe(
      true
    );
  });
});
