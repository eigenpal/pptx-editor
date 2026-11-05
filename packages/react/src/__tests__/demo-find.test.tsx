// Find and replace on the React demo in happy-dom: Ctrl+F opens the bar, typing selects the
// first match on its slide, Ctrl+H shows the replace row, Enter there replaces the shown match,
// and Escape closes the bar. Ported from the Playwright `e2e/find-export.spec.ts`; the PNG
// exports stay in the browser spec, which has a real canvas.

import { afterEach, describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { closeDemo, fill, openDemo, press, screen, until } from './demo-harness.tsx';

afterEach(closeDemo);

/** Matches of `query` in the open deck, through the engine's own find. */
const count = (editor: PptxEditorInstance, query: string): number =>
  editor.find.matches(query, { matchCase: true }).length;

describe('demo find and replace', () => {
  test('find a word and replace it', async () => {
    const { editor } = await openDemo();
    const total = count(editor, 'HarfBuzz');
    expect(total).toBeGreaterThan(1);

    // Ctrl+F opens find; typing selects the first match on its slide, before Enter.
    await press('Control+f');
    const find = screen.getByRole('textbox', { name: 'Find' });
    expect(document.activeElement).toBe(find);
    await fill(find, 'HarfBuzz');
    const counter = () => document.querySelector('.pptx-find__count')?.textContent;
    await until(() => expect(counter()).toBe(`1 of ${total}`));
    expect(document.activeElement).toBe(find);
    expect(document.querySelectorAll('.pptx-rail__item')[1].getAttribute('aria-current')).toMatch(
      /.+/
    );

    // Ctrl+H shows the replace row; Enter there replaces the shown match.
    await press('Control+h');
    const replace = screen.getByRole('textbox', { name: 'Replace' });
    await fill(replace, 'Shaper');
    await press('Enter');
    await until(() => expect(count(editor, 'Shaper')).toBe(1));
    expect(count(editor, 'HarfBuzz')).toBe(total - 1);
    await until(() => expect(counter()).toBe(`1 of ${total - 1}`));

    // Escape closes the bar and gives the focus back to the slide.
    await press('Escape');
    expect(screen.queryByRole('textbox', { name: 'Find' })).toBeNull();
  });
});
