import { describe, expect, test } from 'bun:test';
import { StrictMode } from 'react';
import { fireEvent, render } from '@testing-library/react';
import { keepFocus, MenuItem, Popover } from '../chrome.tsx';

// Chrome must not steal the selection: mousedown is prevented everywhere except on the
// controls that need focus to work (AGENTS.md, "React adapter").
describe('keepFocus', () => {
  const prevented = (tag: string): boolean => {
    const { container } = render(
      <div onMouseDown={keepFocus}>{tag === 'BUTTON' ? <button /> : null}</div>
    );
    const root = container.firstElementChild as HTMLElement;
    const target =
      tag === 'BUTTON'
        ? (root.firstElementChild as HTMLElement)
        : root.appendChild(document.createElement(tag.toLowerCase()));
    return !fireEvent.mouseDown(target);
  };

  test('prevents mousedown on buttons and plain elements', () => {
    expect(prevented('BUTTON')).toBe(true);
    expect(prevented('DIV')).toBe(true);
  });

  test('keeps the default on INPUT, SELECT, and TEXTAREA', () => {
    for (const tag of ['INPUT', 'SELECT', 'TEXTAREA']) expect(prevented(tag)).toBe(false);
  });
});

// `examples/vite` renders under StrictMode, which runs the focus effect twice. The second run
// finds focus inside the panel; the element saved by the first run must still get focus back.
describe('Popover focus under StrictMode', () => {
  test('a pick after a pointer open gives focus back to the slide text, not the trigger', () => {
    const slide = document.createElement('textarea');
    const anchor = document.createElement('button');
    document.body.append(slide, anchor);
    slide.focus();
    const Host = ({ open }: { open: boolean }) =>
      open ? (
        <Popover anchor={anchor} onClose={() => {}}>
          <MenuItem onSelect={() => {}}>One</MenuItem>
          <MenuItem onSelect={() => {}}>Two</MenuItem>
        </Popover>
      ) : null;
    const view = render(<Host open />, { wrapper: StrictMode });
    const panel = document.querySelector<HTMLElement>('.pptx-menu')!;
    expect(document.activeElement === panel).toBe(true);
    fireEvent.keyDown(panel, { key: 'ArrowDown' });
    expect(document.activeElement === panel.querySelector('[role="menuitem"]')).toBe(true);
    view.rerender(<Host open={false} />);
    expect(document.activeElement === slide).toBe(true);
    view.unmount();
    slide.remove();
    anchor.remove();
  });
});
