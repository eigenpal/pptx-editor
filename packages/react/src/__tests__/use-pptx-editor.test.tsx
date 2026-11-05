// `usePptxEditor()` outside `PptxEditor.Root` returns null with a development warning, as the
// other parts do (`warnOutsideRoot`); under a root or a host's editor provider it stays quiet.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { cleanup, render } from '@testing-library/react';
import { createPptxEditor, resetDevWarnings } from '@pptx-editor.dev/core/editor';
import { PptxEditorContext, usePptxEditor } from '../context.ts';
import { PptxEditor } from '../index.ts';

afterEach(cleanup);

const OUTSIDE = 'usePptxEditor is used outside PptxEditor.Root';

function Probe({ seen }: { seen: unknown[] }) {
  seen.push(usePptxEditor());
  return null;
}

const warned = (warn: { mock: { calls: unknown[][] } }): boolean =>
  warn.mock.calls.some((c) => String(c[0]).includes(OUTSIDE));

describe('usePptxEditor', () => {
  test('outside the root it returns null and warns', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    const seen: unknown[] = [];
    render(<Probe seen={seen} />);
    expect(seen.at(-1) === null).toBe(true);
    expect(warned(warn)).toBe(true);
    warn.mockRestore();
  });

  test('inside the root, or with a host editor provider, it does not warn', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    const seen: unknown[] = [];
    render(
      <PptxEditor.Root>
        <Probe seen={seen} />
      </PptxEditor.Root>
    );
    const editor = createPptxEditor({ fixedMeasurer: true });
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe seen={seen} />
      </PptxEditorContext.Provider>
    );
    expect(seen.at(-1) === editor).toBe(true);
    expect(warned(warn)).toBe(false);
    warn.mockRestore();
    editor.destroy();
  });
});
