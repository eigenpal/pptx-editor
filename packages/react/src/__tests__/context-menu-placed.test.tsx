import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactElement } from 'react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { PptxEditor } from '../index.ts';

// A `PptxEditor.ContextMenu` placed in the root or the viewport, as DOCX Editor allows
// (docs/site/content/customization/menus.mdx).

const IN = 914_400;
const PX = 9_525;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

const rightClick = (stage: HTMLElement, x: number, y: number) =>
  act(() => {
    fireEvent.contextMenu(stage, { clientX: (x * IN) / PX, clientY: (y * IN) / PX });
  });

describe('a PptxEditor.ContextMenu placed in the root', () => {
  // The Vue twin is packages/vue/test/menus.test.ts ('a PptxEditorContextMenu placed in the root').
  async function placed(
    menu: (props: { disabled: boolean }) => ReactElement,
    popups?: Parameters<typeof PptxEditorRoot>[0]['popups']
  ) {
    let editor: PptxEditorInstance | null = null;
    const ui = (disabled: boolean) => (
      <PptxEditorRoot onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
        <PptxEditor.Viewport>
          <PptxEditorContent />
          {menu({ disabled })}
        </PptxEditor.Viewport>
      </PptxEditorRoot>
    );
    const view = render(ui(false));
    await act(
      async () =>
        void (await editor!.load(
          buildPptx({ slides: [{ shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN }) }] })
        ))
    );
    const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
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
    return { stage, disable: (disabled: boolean) => view.rerender(ui(disabled)) };
  }

  test('it is the canvas menu, with its rows, and onOpenChange hears each open and close', async () => {
    const opens: boolean[] = [];
    let asked = 0;
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { stage } = await placed(() => (
        <PptxEditor.ContextMenu className="placed" onOpenChange={(open) => opens.push(open)}>
          <PptxEditor.ContextMenu.Slot slotId="format.painter" hidden />
          <PptxEditor.ContextMenu.Item label="Ask legal" onSelect={() => (asked += 1)} />
        </PptxEditor.ContextMenu>
      ));
      expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
      rightClick(stage, 1.5, 1.5);
      const menu = document.querySelector<HTMLElement>('.pptx-contextmenu.placed')!;
      expect(menu !== null).toBe(true);
      expect(menu.querySelector('[data-slot="format.painter"]') === null).toBe(true);
      expect(menu.querySelector('[data-slot="arrange.order.front"]') !== null).toBe(true);
      expect(opens).toEqual([true]);
      act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Ask legal' })));
      expect(asked).toBe(1);
      expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
      expect(opens).toEqual([true, false]);
      const said = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(said).not.toContain('PptxEditor.ContextMenu is outside');
    } finally {
      warn.mockRestore();
    }
  });

  test('disabled shows no menu, and turning it on closes an open menu', async () => {
    const opens: boolean[] = [];
    const { stage, disable } = await placed(({ disabled }) => (
      <PptxEditor.ContextMenu disabled={disabled} onOpenChange={(open) => opens.push(open)} />
    ));
    rightClick(stage, 1.5, 1.5);
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(1);
    act(() => disable(true));
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(0);
    expect(opens).toEqual([true, false]);
    rightClick(stage, 1.5, 1.5);
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(0);
    act(() => disable(false));
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(0);
    rightClick(stage, 1.5, 1.5);
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(1);
  });

  test('an explicit popups.contextMenu entry wins over a placed menu', async () => {
    const { stage } = await placed(() => <PptxEditor.ContextMenu className="placed" />, {
      contextMenu: (props) => <PptxEditor.ContextMenu {...props} className="explicit" />,
    });
    rightClick(stage, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu.explicit') !== null).toBe(true);
    expect(document.querySelector('.pptx-contextmenu.placed') === null).toBe(true);
    cleanup();
    const off = await placed(() => <PptxEditor.ContextMenu className="placed" />, {
      contextMenu: false,
    });
    rightClick(off.stage, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });
});
