import { describe, expect, test } from 'bun:test';
import type { Editor, EditorSelection } from '../../contracts/index.ts';
import {
  createHyperlinkPopupController,
  hyperlinkEntryReads,
  hyperlinkPopupFromRequest,
  markReadingPopup,
} from '../link-popup.ts';

const caret = (offset: number): EditorSelection => ({
  kind: 'text',
  nodeId: 'n1',
  anchor: { paragraph: 0, offset },
  focus: { paragraph: 0, offset },
});

/** An editor with one text shape: a caret and the link under it, as the snapshot reports. */
function fakeEditor(start: { selection: EditorSelection; link?: string; editable?: boolean }) {
  let selection = start.selection;
  let link = start.link;
  const listeners = new Set<() => void>();
  const calls: (string | null)[] = [];
  let refuse: string | null = null;
  const editor = {
    snapshot: () => ({ selection, format: link === undefined ? { bold: false } : { link } }),
    subscribe: (l: () => void) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
    can: () =>
      start.editable === false ? { ok: false, reason: 'disabled.readOnly' } : { ok: true },
    exec: (command: { type: string; url: string | null }) => {
      if (refuse) return { ok: false, reason: refuse };
      calls.push(command.url);
      link = command.url ?? undefined;
      for (const l of listeners) l();
      return { ok: true };
    },
  } as unknown as Editor;
  return {
    editor,
    calls,
    move(next: EditorSelection, nextLink?: string) {
      selection = next;
      link = nextLink;
      for (const l of listeners) l();
    },
    refuseWith(reason: string | null) {
      refuse = reason;
    },
  };
}

describe('link popup controller', () => {
  test('reading opens only on a caret inside a link, with the safe target', () => {
    const plain = fakeEditor({ selection: caret(2) });
    const c1 = createHyperlinkPopupController(plain.editor);
    c1.openReading(null);
    expect(c1.getState().mode).toBe('closed');

    const range = fakeEditor({
      selection: {
        kind: 'text',
        nodeId: 'n1',
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 3 },
      },
      link: 'https://example.com/',
    });
    const c2 = createHyperlinkPopupController(range.editor);
    c2.openReading(null);
    expect(c2.getState().mode).toBe('closed');

    const linked = fakeEditor({ selection: caret(2), link: 'https://example.com/a' });
    const c3 = createHyperlinkPopupController(linked.editor);
    c3.openReading(null);
    expect(c3.getState()).toMatchObject({
      mode: 'reading',
      url: 'https://example.com/a',
      href: 'https://example.com/a',
      canEdit: true,
    });
  });

  test('an unsafe target shows but is neither opened nor copied', async () => {
    const linked = fakeEditor({ selection: caret(1), link: 'javascript:alert(1)' });
    const writes: string[] = [];
    const c = createHyperlinkPopupController(linked.editor, () => ({
      writeText: async (t: string) => void writes.push(t),
    }));
    c.openReading(null);
    expect(c.getState().href).toBe(null);
    expect(await c.copy()).toBe(false);
    expect(writes).toEqual([]);
  });

  test('a move of the caret closes reading mode; editing stays open', () => {
    const linked = fakeEditor({ selection: caret(2), link: 'https://example.com/' });
    const c = createHyperlinkPopupController(linked.editor);
    c.openReading(null);
    linked.move(caret(3), 'https://example.com/');
    expect(c.getState().mode).toBe('closed');
    c.openEditing(null);
    linked.move(caret(4), 'https://example.com/');
    expect(c.getState().mode).toBe('editing');
  });

  test('Edit, Apply, a refusal, and Unlink run the one setLink command', () => {
    const linked = fakeEditor({ selection: caret(2), link: 'https://example.com/' });
    const c = createHyperlinkPopupController(linked.editor);
    c.openReading(null);
    c.beginEdit();
    expect(c.getState()).toMatchObject({ mode: 'editing', draft: 'https://example.com/' });
    c.setUrl('bad');
    linked.refuseWith('errors.unsafeLink');
    expect(c.commitEdit()).toBe(false);
    expect(c.getState()).toMatchObject({ mode: 'editing', error: 'errors.unsafeLink' });
    c.setUrl('example.org');
    expect(c.getState().error).toBe(null);
    linked.refuseWith(null);
    expect(c.commitEdit()).toBe(true);
    expect(c.getState().mode).toBe('closed');
    c.openReading(null);
    expect(c.unlink()).toBe(true);
    expect(linked.calls).toEqual(['example.org', null]);
    expect(c.getState().mode).toBe('closed');
  });

  test('Copy writes the safe target and says so', async () => {
    const linked = fakeEditor({ selection: caret(2), link: 'https://example.com/x' });
    const writes: string[] = [];
    const c = createHyperlinkPopupController(linked.editor, () => ({
      writeText: async (t: string) => void writes.push(t),
    }));
    c.openReading(null);
    expect(await c.copy()).toBe(true);
    expect(writes).toEqual(['https://example.com/x']);
    expect(c.getState().copied).toBe(true);
  });

  test('a read-only editor reads the link but cannot edit it', () => {
    const linked = fakeEditor({
      selection: caret(2),
      link: 'https://example.com/',
      editable: false,
    });
    const c = createHyperlinkPopupController(linked.editor);
    c.openReading(null);
    expect(c.getState()).toMatchObject({ mode: 'reading', canEdit: false });
    // Edit does nothing, and the edit form never opens: it could only fail.
    c.beginEdit();
    expect(c.getState().mode).toBe('reading');
    c.close();
    c.openEditing(null);
    expect(c.getState().mode).toBe('closed');
  });

  test('a popup from props alone shows them, and closing calls onClose', () => {
    let closed = 0;
    const c = hyperlinkPopupFromRequest(null, {
      anchor: null,
      url: 'https://example.com/',
      mode: 'reading',
      onClose: () => void (closed += 1),
    });
    expect(c.getState()).toMatchObject({ mode: 'reading', href: 'https://example.com/' });
    c.close();
    expect(closed).toBe(1);
  });

  test('a new link under the same caret (a peer, an undo) refreshes the panel', () => {
    const f = fakeEditor({ selection: caret(2), link: 'https://a.test/' });
    const c = createHyperlinkPopupController(f.editor);
    c.openReading(null);
    f.move(caret(2), 'https://b.test/');
    expect(c.getState()).toMatchObject({
      mode: 'reading',
      url: 'https://b.test/',
      href: 'https://b.test/',
    });
  });

  test('openTarget opens the safe target through a link element, never an unsafe one', () => {
    const opened: Record<string, string>[] = [];
    // A document with no DOM: the link element records its attributes when clicked.
    const doc = {
      createElement: () => {
        const attrs: Record<string, string> = {};
        return {
          setAttribute: (k: string, v: string) => void (attrs[k] = v),
          click: () => void opened.push(attrs),
        };
      },
    };
    const anchor = { ownerDocument: doc } as unknown as HTMLElement;
    const safe = createHyperlinkPopupController(
      fakeEditor({ selection: caret(2), link: 'https://a.test/' }).editor
    );
    safe.openReading(anchor);
    expect(safe.openTarget()).toBe(true);
    expect(opened).toEqual([
      { href: 'https://a.test/', target: '_blank', rel: 'noopener noreferrer' },
    ]);
    const unsafe = createHyperlinkPopupController(
      fakeEditor({ selection: caret(2), link: 'javascript:alert(1)' }).editor
    );
    unsafe.openReading(anchor);
    expect(unsafe.openTarget()).toBe(false);
    expect(opened.length).toBe(1);
  });
});

describe('popups.hyperlink and reading mode', () => {
  test('the packaged panel reads; a host renderer only when it asks', () => {
    expect(hyperlinkEntryReads(undefined)).toBe(true);
    expect(hyperlinkEntryReads(false)).toBe(false);
    expect(hyperlinkEntryReads(() => null)).toBe(false);
    expect(hyperlinkEntryReads(markReadingPopup(() => null))).toBe(true);
  });
});
