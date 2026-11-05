import { afterAll, afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactNode } from 'react';
import {
  TOOLBAR_NAMED_PARTS,
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { toolbarRoom } from '../toolbar-measure.ts';
import { Toolbar } from '../index.ts';
import { tooltipText } from './tooltip-text.ts';

// DOCX Editor's toolbar parts (docs/site/content/customization/toolbar.mdx): named parts,
// `Toolbar.Action`, group `priority`, `pinned`, and `after`, and `overflow={false}`.

const IN = 914_400;

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  // The text box is selected: the bar shows its text, shape, and arrange groups.
  editor.setSelection({
    kind: 'shapes',
    nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
  });
  return editor;
}

// happy-dom has no layout: every element measures 9 pixels wide (so a group with its
// separator costs 18, and the More button with its separator 18), and the bar width decides
// how many fit.
let barWidth = 100_000;
const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
const offset = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth');
beforeAll(() => {
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => barWidth,
  });
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', { configurable: true, get: () => 9 });
});
afterAll(() => {
  if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
  if (offset) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', offset);
});
afterEach(() => {
  cleanup();
  barWidth = 100_000;
});

function renderToolbar(editor: PptxEditorInstance, toolbar: ReactNode) {
  return render(<PptxEditorContext.Provider value={editor}>{toolbar}</PptxEditorContext.Provider>);
}

const barGroups = (container: HTMLElement) =>
  [...container.querySelectorAll('.pptx-toolbar > [data-group]')].map((g) =>
    g.getAttribute('data-group')
  );

describe('named parts', () => {
  test('every named part stands for the slot in core’s table', () => {
    for (const [name, slot] of Object.entries(TOOLBAR_NAMED_PARTS))
      expect((Toolbar as unknown as Record<string, { pptxSlot: string }>)[name]!.pptxSlot).toBe(
        slot
      );
  });

  test('a named part overrides its slot in place, and hidden removes it', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Bold className="acme-bold" icon={<b data-testid="b" />} />
        <Toolbar.Strike hidden />
      </Toolbar>
    );
    const bold = container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(bold.classList.contains('acme-bold')).toBe(true);
    expect(bold.querySelector('[data-testid="b"]') !== null).toBe(true);
    expect(bold.closest('[data-group]')!.getAttribute('data-group')).toBe('text');
    expect(container.querySelector('[data-slot="text.strike"]') === null).toBe(true);
  });

  test('Toolbar.Autofit is the Text fitting picker, and hidden removes it', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    const fit = container.querySelector<HTMLElement>('[data-slot="paragraph.autofit"]')!;
    expect(fit.closest('[data-group]')!.getAttribute('data-group')).toBe('paragraph');
    cleanup();
    const hidden = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Autofit hidden />
      </Toolbar>
    );
    expect(hidden.container.querySelector('[data-slot="paragraph.autofit"]') === null).toBe(true);
  });

  test('with preset={false}, named parts render where they stand', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Undo />
        <Toolbar.Separator />
        <Toolbar.Italic asChild>
          <button type="button" data-testid="italic" />
        </Toolbar.Italic>
      </Toolbar>
    );
    expect(
      [...container.querySelectorAll('[data-slot]')].map((b) => b.getAttribute('data-slot'))
    ).toEqual(['history.undo', 'text.italic']);
    expect(container.querySelector('[role="separator"]') !== null).toBe(true);
    expect(screen.getByTestId('italic').getAttribute('aria-label')).toBe('Italic (Ctrl+I)');
  });
});

describe('Toolbar.Action', () => {
  test('a host action gets the button class, state attributes, and the caret guard', () => {
    let runs = 0;
    render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Send" active className="acme" onSelect={() => (runs += 1)}>
          S
        </Toolbar.Action>
        <Toolbar.Action
          label="Lock"
          disabled
          disabledReason="Viewing"
          onSelect={() => (runs += 10)}
        />
      </Toolbar>
    );
    const send = screen.getByRole('button', { name: 'Send' });
    // Text content ("S") gets the text width, as a label does.
    expect(send.className).toBe('pptx-toolbar__button acme pptx-toolbar__button--text');
    expect(send.getAttribute('aria-pressed')).toBe('true');
    expect(send.hasAttribute('data-active')).toBe(true);
    const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    send.dispatchEvent(down);
    expect(down.defaultPrevented).toBe(true);
    act(() => void fireEvent.click(send));
    const lock = screen.getByRole('button', { name: 'Lock' });
    expect(lock.hasAttribute('data-disabled')).toBe(true);
    expect(tooltipText(lock)).toBe('Lock\nViewing');
    // One description: the reason text is hidden, so the tree does not show it again as text.
    const reason = document.getElementById(lock.getAttribute('aria-describedby')!)!;
    expect(reason.textContent).toBe('Viewing');
    expect(reason.hidden).toBe(true);
    act(() => void fireEvent.click(lock));
    expect(runs).toBe(1);
  });

  test('asChild puts the behavior on the host element', () => {
    let runs = 0;
    render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Help" asChild onSelect={() => (runs += 1)}>
          <a href="#help" className="acme-link">
            ?
          </a>
        </Toolbar.Action>
      </Toolbar>
    );
    const link = screen.getByRole('button', { name: 'Help' });
    expect(link.localName).toBe('a');
    expect(link.className).toBe('pptx-toolbar__button acme-link');
    act(() => void fireEvent.click(link));
    expect(runs).toBe(1);
  });

  test('onSelect gets no argument from a click or from Space on a host element', () => {
    const calls: unknown[][] = [];
    render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Send" onSelect={(...args: unknown[]) => void calls.push(args)} />
        <Toolbar.Action
          label="Help"
          asChild
          onSelect={(...args: unknown[]) => void calls.push(args)}
        >
          <span>?</span>
        </Toolbar.Action>
      </Toolbar>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Send' })));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Help' })));
    act(() => void fireEvent.keyDown(screen.getByRole('button', { name: 'Help' }), { key: ' ' }));
    expect(calls).toEqual([[], [], []]);
  });
});

describe('group placement and overflow', () => {
  test('a high priority keeps a host group on a narrow bar while built-in groups leave', async () => {
    barWidth = 60;
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="acme" label="Acme" priority={175} className="acme-group">
          <Toolbar.Action label="Ask" />
        </Toolbar.Group>
      </Toolbar>
    );
    expect(barGroups(container)).toEqual(['history', 'acme']);
    expect(container.querySelector('.acme-group') !== null).toBe(true);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More' })));
    const menu = document.querySelector('.pptx-toolbar__overflow')!;
    expect(menu.querySelector('[data-group="more"]') !== null).toBe(true);
  });

  test('a pinned host group stays at every width, and overflowContent fills its menu entry', async () => {
    barWidth = 20;
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="pin" label="Pin" pinned>
          <Toolbar.Action label="Pinned" />
        </Toolbar.Group>
        <Toolbar.Group id="gone" label="Gone" overflowContent={() => <p data-testid="row" />}>
          <Toolbar.Action label="On bar" />
        </Toolbar.Group>
      </Toolbar>
    );
    expect(barGroups(container)).toContain('pin');
    expect(barGroups(container)).not.toContain('gone');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More' })));
    expect(screen.getByTestId('row')).toBeTruthy();
  });

  test('after places a host group behind the group it names', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="acme" label="Acme" after="history">
          <Toolbar.Action label="Ask" />
        </Toolbar.Group>
      </Toolbar>
    );
    expect(barGroups(container).slice(0, 3)).toEqual(['history', 'acme', 'font']);
  });

  test('overflow={false} wraps the bar and keeps every group on it', async () => {
    barWidth = 20;
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar overflow={false} />);
    expect(container.querySelector('.pptx-toolbar--wrap') !== null).toBe(true);
    expect(screen.queryByRole('button', { name: 'More' }) === null).toBe(true);
    expect(barGroups(container)).toContain('more');
  });
});

describe('toolbar review fixes', () => {
  test('a loose action goes after the groups and never collapses', async () => {
    barWidth = 60;
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Action label="Send" />
      </Toolbar>
    );
    const fixed = container.querySelector('.pptx-toolbar > [data-fixed]')!;
    expect(fixed !== null).toBe(true);
    // With no icon, the action shows its label.
    expect(screen.getByRole('button', { name: 'Send' }).textContent).toBe('Send');
  });

  test('a preset={false} bar wraps, as DOCX Editor', () => {
    const { container } = render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Ask" />
      </Toolbar>
    );
    const bar = container.querySelector('.pptx-toolbar')!;
    expect(bar.classList.contains('pptx-toolbar--wrap')).toBe(true);
    expect(bar.getAttribute('data-overflow')).toBe('wrap');
  });

  test('a built-in group keeps its controls in the More menu before its overflowContent', async () => {
    barWidth = 60;
    const editor = await open();
    renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Group id="more" overflowContent={() => <span data-testid="extra" />} />
      </Toolbar>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More' })));
    const more = document.querySelector('.pptx-toolbar__overflow [data-group="more"]')!;
    expect(more.querySelector('[data-testid="extra"]') !== null).toBe(true);
    expect(more.querySelectorAll('button').length).toBeGreaterThan(0);
    expect(more.querySelector('.pptx-toolbar__overflow-heading')!.textContent).toBeTruthy();
  });

  test('More is a dialog trigger: ArrowDown opens it at its first control', async () => {
    barWidth = 60;
    const editor = await open();
    renderToolbar(editor, <Toolbar />);
    const more = screen.getByRole('button', { name: 'More' });
    expect(more.getAttribute('aria-haspopup')).toBe('dialog');
    more.focus();
    act(() => void fireEvent.keyDown(more, { key: 'ArrowDown' }));
    const panel = document.querySelector<HTMLElement>('.pptx-menu[role="dialog"]')!;
    expect(more.getAttribute('aria-controls')).toBe(panel.id);
    expect(panel.contains(document.activeElement)).toBe(true);
  });

  test('turning overflow off while groups are collapsed brings them back', async () => {
    barWidth = 60;
    const editor = await open();
    const { container, rerender } = renderToolbar(editor, <Toolbar />);
    expect(screen.queryByRole('button', { name: 'More' }) !== null).toBe(true);
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar overflow={false} />
      </PptxEditorContext.Provider>
    );
    expect(screen.queryByRole('button', { name: 'More' }) === null).toBe(true);
    expect(barGroups(container)).toContain('more');
  });

  test('asChild on a non-button element runs on Enter and Space; misuse warns and falls back', () => {
    let runs = 0;
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      render(
        <Toolbar preset={false}>
          <Toolbar.Action label="Help" asChild onSelect={() => (runs += 1)}>
            <span>?</span>
          </Toolbar.Action>
          <Toolbar.Action label="Broken" asChild>
            text
          </Toolbar.Action>
        </Toolbar>
      );
      const help = screen.getByRole('button', { name: 'Help' });
      fireEvent.keyDown(help, { key: 'Enter' });
      fireEvent.keyDown(help, { key: ' ' });
      expect(runs).toBe(2);
      const broken = screen.getByRole('button', { name: 'Broken' });
      expect(broken.localName).toBe('button');
      expect(broken.textContent).toBe('Broken');
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('Toolbar.Action "Broken" has asChild but not exactly one element');
    } finally {
      warn.mockRestore();
    }
  });

  test('a named part warns with its own name', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      renderToolbar(
        editor,
        <Toolbar>
          {/* Misuse on purpose: the types refuse asChild with a text child. */}
          <Toolbar.Bold {...({ asChild: true, children: 'x' } as object)} />
        </Toolbar>
      );
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('Toolbar.Bold "text.bold" has asChild');
    } finally {
      warn.mockRestore();
    }
  });
});

describe('toolbar review fixes, round 2', () => {
  test('groups come back when the parent has room, though the bar itself is narrow', async () => {
    const own = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth')!;
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('pptx-toolbar') ? 60 : 100_000;
      },
    });
    try {
      const editor = await open();
      const { container } = renderToolbar(editor, <Toolbar />);
      expect(barGroups(container)).toContain('more');
      expect(screen.queryByRole('button', { name: 'More' }) === null).toBe(true);
    } finally {
      Object.defineProperty(HTMLElement.prototype, 'clientWidth', own);
    }
  });

  test('parts inside a fragment apply, and empty groups leave the bar', async () => {
    const editor = await open();
    const readOnly = true;
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        {readOnly && (
          <>
            <Toolbar.Find hidden />
            <Toolbar.Replace hidden />
          </>
        )}
        <Toolbar.Group id="acme" label="Acme">
          {false}
        </Toolbar.Group>
        {''}
      </Toolbar>
    );
    expect(barGroups(container)).not.toContain('find');
    expect(barGroups(container)).not.toContain('acme');
    expect(container.querySelector('[data-fixed]') === null).toBe(true);
  });

  test('a part for a control off the preset bar goes after the groups', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.SmartGuides />
      </Toolbar>
    );
    expect(container.querySelector('[data-fixed] [data-slot="view.smartGuides"]') !== null).toBe(
      true
    );
  });

  test('children of a named button part are its icon; a picker part warns with its name', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const { container } = renderToolbar(
        editor,
        <Toolbar>
          <Toolbar.Bold>
            <b data-testid="bold-icon" />
          </Toolbar.Bold>
          <Toolbar.FontFamily {...({ icon: <i /> } as object)} />
        </Toolbar>
      );
      const bold = container.querySelector('[data-slot="text.bold"]')!;
      expect(bold.querySelector('[data-testid="bold-icon"]') !== null).toBe(true);
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('Toolbar.FontFamily');
    } finally {
      warn.mockRestore();
    }
  });

  test('Escape closes a More panel opened with the mouse, even when a handler stops it', async () => {
    barWidth = 60;
    const editor = await open();
    renderToolbar(editor, <Toolbar />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More' })));
    expect(document.querySelector('.pptx-toolbar__overflow') !== null).toBe(true);
    const surface = document.createElement('div');
    surface.addEventListener('keydown', (e) => e.stopPropagation());
    document.body.append(surface);
    act(() => void fireEvent.keyDown(surface, { key: 'Escape' }));
    expect(document.querySelector('.pptx-toolbar__overflow') === null).toBe(true);
    surface.remove();
  });

  test('a text action gets room for its text', () => {
    render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Ask legal">Ask legal</Toolbar.Action>
      </Toolbar>
    );
    expect(
      screen
        .getByRole('button', { name: 'Ask legal' })
        .classList.contains('pptx-toolbar__button--text')
    ).toBe(true);
  });
});

describe('toolbar review fixes, round 3', () => {
  test('the bar is one Tab stop; arrows move between controls, disabled ones too', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    const bar = container.querySelector<HTMLElement>('.pptx-toolbar')!;
    // Natively disabled buttons (picker carets) are no Tab stops anyway.
    const stops = [...bar.querySelectorAll<HTMLButtonElement>('button')].filter(
      (b) => !b.disabled && b.tabIndex === 0
    );
    expect(stops.length).toBe(1);
    const undo = bar.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    undo.focus();
    act(() => void fireEvent.keyDown(undo, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(bar.querySelector('[data-slot="history.redo"]'));
    expect((document.activeElement as HTMLElement).tabIndex).toBe(0);
    expect(undo.tabIndex).toBe(-1);
  });

  test('the arrows leave the Font size field at the ends of its text (round 4)', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    const bar = container.querySelector<HTMLElement>('.pptx-toolbar')!;
    const size = bar.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
    const next = bar.querySelector<HTMLElement>('[aria-label="Increase font size"]')!;
    act(() => void fireEvent.change(size, { target: { value: '18' } }));
    size.focus();
    size.setSelectionRange(0, 0);
    act(() => void fireEvent.keyDown(size, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(size);
    size.setSelectionRange(2, 2);
    act(() => void fireEvent.keyDown(size, { key: 'ArrowRight' }));
    expect(document.activeElement).toBe(next);
  });

  test('the More panel is visible before focus moves into it, and focus leaving closes it', async () => {
    barWidth = 60;
    const editor = await open();
    renderToolbar(editor, <Toolbar />);
    const more = screen.getByRole('button', { name: 'More' });
    more.focus();
    act(() => void fireEvent.keyDown(more, { key: 'ArrowDown' }));
    const panel = document.querySelector<HTMLElement>('.pptx-menu[role="dialog"]')!;
    expect(panel.style.visibility).toBe('');
    const outside = document.createElement('button');
    document.body.append(outside);
    act(() => outside.focus());
    expect(document.querySelector('.pptx-menu[role="dialog"]') === null).toBe(true);
    outside.remove();
  });

  test('an Escape during IME composition leaves the More panel open', async () => {
    barWidth = 60;
    const editor = await open();
    renderToolbar(editor, <Toolbar />);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More' })));
    act(() => void fireEvent.keyDown(document.body, { key: 'Escape', isComposing: true }));
    expect(document.querySelector('.pptx-toolbar__overflow') !== null).toBe(true);
  });

  test('asChild on a color split warns and keeps the packaged control', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const { container } = renderToolbar(
        editor,
        <Toolbar>
          <Toolbar.Slot slotId="text.highlight" asChild>
            <span />
          </Toolbar.Slot>
        </Toolbar>
      );
      expect(container.querySelector('button[data-slot="text.highlight"]') !== null).toBe(true);
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('"text.highlight" renders a picker');
    } finally {
      warn.mockRestore();
    }
  });

  test('an action with an icon and text lays them out side by side', () => {
    render(
      <Toolbar preset={false}>
        <Toolbar.Action label="Send">
          <svg />
          Send
        </Toolbar.Action>
      </Toolbar>
    );
    expect(
      screen.getByRole('button', { name: 'Send' }).classList.contains('pptx-toolbar__button--text')
    ).toBe(true);
  });
});

describe('toolbar review fixes, round 5', () => {
  test('a disabled picker and the size field stay reachable, say why, and open nothing', async () => {
    const editor = await open();
    // Nothing selected: the font controls have no text to act on.
    editor.setSelection({ kind: 'none' });
    const { container } = renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Slot slotId="font.family" />
        <Toolbar.Slot slotId="font.size" />
      </Toolbar>
    );
    const family = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(family.getAttribute('aria-disabled')).toBe('true');
    const reasonId = family.getAttribute('aria-describedby')!;
    expect(document.getElementById(reasonId)!.textContent!.length).toBeGreaterThan(0);
    act(() => void fireEvent.click(family));
    expect(document.querySelector('.pptx-font-picker') === null).toBe(true);
    const size = container.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
    expect(size.getAttribute('aria-disabled')).toBe('true');
    expect(size.readOnly).toBe(true);
  });

  test('a bar in a grid cell measures the cell as its room', () => {
    const grid = document.createElement('div');
    grid.style.display = 'grid';
    const bar = document.createElement('div');
    grid.append(bar);
    document.body.append(grid);
    let width = '';
    Object.defineProperty(bar, 'clientWidth', {
      configurable: true,
      get: () => (bar.style.width === '100%' ? 640 : 300),
    });
    const room = toolbarRoom(bar, getComputedStyle(bar));
    width = bar.style.width;
    grid.remove();
    expect(room).toBe(640);
    expect(width).toBe('');
  });
});

// The React twin of the Vue `ToolbarSlot` class tests: `className` on `Toolbar.Slot` goes on
// the control in the preset bar and in a hand-built bar.
describe('Toolbar.Slot className', () => {
  test('className goes on the control in the preset bar and with preset={false}', async () => {
    const editor = await open();
    const preset = renderToolbar(
      editor,
      <Toolbar>
        <Toolbar.Slot slotId="text.bold" className="acme-bold" />
      </Toolbar>
    );
    const bold = preset.container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(bold.classList.contains('acme-bold')).toBe(true);
    expect(bold.classList.contains('pptx-toolbar__button')).toBe(true);
    expect(bold.closest('[data-group]')!.getAttribute('data-group')).toBe('text');
    cleanup();
    const hand = renderToolbar(
      editor,
      <Toolbar preset={false}>
        <Toolbar.Slot slotId="text.italic" className="acme-italic" />
      </Toolbar>
    );
    const italic = hand.container.querySelector<HTMLElement>('[data-slot="text.italic"]')!;
    expect(italic.classList.contains('acme-italic')).toBe(true);
  });

  test('className or icon on a slot whose children replace the control warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const { container } = renderToolbar(
        editor,
        <Toolbar>
          <Toolbar.Slot slotId="text.bold" className="acme-bold">
            <i data-testid="mine" />
          </Toolbar.Slot>
          <Toolbar.Slot slotId="text.italic" icon={<b />}>
            <i data-testid="mine-2" />
          </Toolbar.Slot>
        </Toolbar>
      );
      expect(container.querySelector('[data-testid="mine"]')).not.toBeNull();
      expect(container.querySelector('.acme-bold')).toBeNull();
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain(
        'Toolbar.Slot "text.bold" has its own content, which replaces the control, so its className is ignored.'
      );
      expect(text).toContain('Toolbar.Slot "text.italic" has its own content');
      expect(text).toContain('so its icon is ignored');
    } finally {
      warn.mockRestore();
    }
  });

  test('data-*, style, and id on a slot warn and stay off the control', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const { container } = renderToolbar(
        editor,
        <Toolbar>
          <Toolbar.Slot slotId="text.bold" data-test="bold" {...{ style: { color: 'red' } }} />
          <Toolbar.Italic {...{ id: 'acme-italic' }} className="acme-italic" />
        </Toolbar>
      );
      const bold = container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
      expect(bold.hasAttribute('data-test')).toBe(false);
      expect(container.querySelector('#acme-italic')).toBeNull();
      expect(container.querySelector('.acme-italic')).not.toBeNull();
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain(
        'Toolbar.Slot "text.bold" does not pass data-test, style to its control, so they are ignored.'
      );
      expect(text).toContain(
        'Toolbar.Italic "text.italic" does not pass id to its control, so it is ignored.'
      );
    } finally {
      warn.mockRestore();
    }
  });
});

describe('the bar for the selection', () => {
  test('Delete removes its own group; the focus stays in the bar', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    const remove = container.querySelector<HTMLElement>('[data-slot="more.delete"]')!;
    act(() => remove.focus());
    act(() => void fireEvent.click(remove));
    // Nothing is selected now: the shape and arrange groups left with the Delete button.
    expect(barGroups(container)).toEqual(['history', 'insert', 'slide', 'find']);
    expect(container.querySelector('.pptx-toolbar')!.contains(document.activeElement)).toBe(true);
  });

  test('focus the user moves off the bar stays off when the groups change', async () => {
    const editor = await open();
    const { container } = renderToolbar(editor, <Toolbar />);
    const remove = container.querySelector<HTMLElement>('[data-slot="more.delete"]')!;
    act(() => remove.focus());
    // A click on the page that takes no focus: the focus goes to the body.
    act(() => remove.blur());
    await act(async () => await Promise.resolve());
    act(() => editor.setSelection({ kind: 'none' }));
    expect(document.activeElement).toBe(document.body);
  });

  test('contextual={false} draws a placed slide slot after the groups', async () => {
    const editor = await open();
    const { container } = renderToolbar(
      editor,
      <Toolbar contextual={false}>
        <Toolbar.Slot slotId="slide.background" />
      </Toolbar>
    );
    expect(container.querySelector('[data-slot="slide.background"]') !== null).toBe(true);
  });

  test('the groups follow the selection; contextual={false} keeps every group', async () => {
    const editor = await open();
    const { container, rerender } = renderToolbar(editor, <Toolbar />);
    expect(barGroups(container)).toContain('shape');
    act(() => editor.setSelection({ kind: 'none' }));
    expect(barGroups(container)).not.toContain('font');
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar contextual={false} />
      </PptxEditorContext.Provider>
    );
    expect(barGroups(container)).toContain('font');
    expect(barGroups(container)).toContain('shape');
  });
});
