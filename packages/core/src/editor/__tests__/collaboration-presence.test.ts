// Presence between the editor and an attached collaboration session: selection conversion,
// publishing every selection change, and painting peers on the canvas as editing chrome.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape, table } from '../../../../../scripts/fixtures/builder.ts';
import { xmlPartOf, type Element, type PptxPackage } from '../../store/index.ts';
import { fromCollaborationSelection, toCollaborationSelection } from '../collaboration-presence.ts';
import {
  createPptxEditor,
  type CollaborationLocalSelection,
  type CollaborationRemoteSelection,
  type EditorCollaborationSession,
  type PptxEditorInstance,
} from '../index.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Hel') + run('lo')) }) +
          shape({ id: 3, x: IN, y: 3 * IN, cx: 4 * IN, cy: IN, text: para('') }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

/** A session that records published selections and serves fixed remote selections. */
function fakeSession() {
  const published: (CollaborationLocalSelection | null)[] = [];
  let remote: readonly CollaborationRemoteSelection[] = [];
  const listeners = new Set<(s: readonly CollaborationRemoteSelection[]) => void>();
  let attached = false;
  const session: EditorCollaborationSession = {
    documentId: 'deck',
    sessionId: 's',
    identity: { actorId: 'me', name: 'Me' },
    get attached() {
      return attached;
    },
    status: () => 'ready',
    statusSnapshot: () => ({ status: 'ready', reason: undefined, lastFailure: undefined }),
    subscribeStatus: () => () => {},
    participants: () => [],
    subscribeParticipants: () => () => {},
    setLocalSelection: (s) => published.push(s),
    remoteSelections: () => remote,
    subscribeRemoteSelections: (l) => (listeners.add(l), () => listeners.delete(l)),
    attach: () => ((attached = true), () => (attached = false)),
    destroy: () => {},
  };
  return {
    session,
    published,
    setRemote(next: readonly CollaborationRemoteSelection[]) {
      remote = next;
      for (const l of listeners) l(remote);
    },
  };
}

async function open(session?: EditorCollaborationSession): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    ...(session ? { modules: [{ id: 'collaboration', collaboration: { session } }] } : {}),
  });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
function elements(pkg: PptxPackage, p: string, match: (el: Element) => boolean): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (match(c)) out.push(c);
      walk(c);
    }
  };
  walk(xmlPartOf(pkg, p)!.root);
  return out;
}
const shapeIds = (e: PptxEditorInstance) =>
  elements(e.session!.store.package, part(e), (el) => el.kind === 'sp').map((el) => el.id);
const textNodeIds = (e: PptxEditorInstance) =>
  elements(e.session!.store.package, part(e), (el) => el.local === 't' && el.kind === 'text').map(
    (t) => t.children[0].id
  );

describe('selection conversion', () => {
  test('a paragraph offset maps to the run text node that holds it, and back', async () => {
    const editor = await open();
    const pkg = editor.session!.store.package;
    const [first] = shapeIds(editor);
    const [hel, lo] = textNodeIds(editor);
    const sel = {
      kind: 'text' as const,
      nodeId: first,
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 4 },
    };
    const sent = toCollaborationSelection(pkg, part(editor), sel);
    expect(sent).toEqual({
      kind: 'text',
      part: part(editor),
      anchor: { node: hel, offset: 1 },
      head: { node: lo, offset: 1 },
    });
    expect(fromCollaborationSelection(pkg, sent!)).toEqual(sel);
  });

  test('a caret in an empty paragraph maps to the paragraph element at offset 0', async () => {
    const editor = await open();
    const pkg = editor.session!.store.package;
    const second = shapeIds(editor)[1];
    const caret = { paragraph: 0, offset: 0 };
    const sent = toCollaborationSelection(pkg, part(editor), {
      kind: 'text',
      nodeId: second,
      anchor: caret,
      focus: caret,
    });
    expect(sent?.kind).toBe('text');
    expect(fromCollaborationSelection(pkg, sent!)).toEqual({
      kind: 'text',
      nodeId: second,
      anchor: caret,
      focus: caret,
    });
  });

  // Two peers who type at once into a cell with no text body leave two bodies in it, which the
  // canvas draws as one (`cellAsSaved`).
  test('a caret in the second text body of a cell maps to the paragraph the canvas draws', async () => {
    const body = (text: string) =>
      `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: table({
            id: 2,
            x: IN,
            y: IN,
            cx: IN,
            cy: 370_840,
            columns: [IN],
            tblPr: '<a:tblPr/>',
            rows: `<a:tr h="370840"><a:tc>${body('xx')}${body('yy')}<a:tcPr/></a:tc></a:tr>`,
          }),
        },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    const pkg = editor.session!.store.package;
    const [frame] = elements(pkg, part(editor), (el) => el.kind === 'graphicFrame');
    const caret = { paragraph: 1, offset: 1 };
    const sel = {
      kind: 'text' as const,
      nodeId: frame.id,
      cell: { row: 0, column: 0 },
      anchor: caret,
      focus: caret,
    };
    const sent = toCollaborationSelection(pkg, part(editor), sel);
    expect(sent?.kind).toBe('text');
    expect(fromCollaborationSelection(pkg, sent!)).toEqual(sel);
  });
});

describe('publishing', () => {
  test('every selection change is published once, and none after detach', async () => {
    const fake = fakeSession();
    const editor = await open(fake.session);
    const [first] = shapeIds(editor);
    editor.setSelection({ kind: 'shapes', nodeIds: [first] });
    editor.setSelection({ kind: 'shapes', nodeIds: [first] });
    const caret = { paragraph: 0, offset: 2 };
    editor.setSelection({ kind: 'text', nodeId: first, anchor: caret, focus: caret });
    const sent = fake.published.filter((s) => s !== null);
    expect(sent).toEqual([
      { kind: 'shapes', part: part(editor), shapes: [first] },
      {
        kind: 'text',
        part: part(editor),
        anchor: { node: textNodeIds(editor)[0], offset: 2 },
        head: { node: textNodeIds(editor)[0], offset: 2 },
      },
    ]);
    const count = fake.published.length;
    editor.destroy();
    editor.setSelection({ kind: 'none' });
    expect(fake.published.length).toBe(count);
  });
});

describe('painting', () => {
  test('peers on this slide are painted in the overlay, never in the slide paint', async () => {
    const g = globalThis as { ResizeObserver?: unknown };
    const had = 'ResizeObserver' in g;
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const fake = fakeSession();
    const editor = await open(fake.session);
    const host = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(host as unknown as Node);
    editor.attach(host);
    const [first, second] = shapeIds(editor);
    const [hel] = textNodeIds(editor);
    fake.setRemote([
      {
        actorId: 'p1',
        name: 'Ada',
        color: '#ff0000',
        selection: { kind: 'shapes', part: part(editor), shapes: [second] },
      },
      {
        actorId: 'p2',
        name: 'Grace',
        color: 'red;}x{', // not a hex color: falls back to the theme accent
        selection: {
          kind: 'text',
          part: part(editor),
          anchor: { node: hel, offset: 1 },
          head: { node: hel, offset: 3 },
        },
      },
      {
        actorId: 'p3',
        name: 'Elsewhere',
        selection: { kind: 'shapes', part: '/ppt/slides/slide2.xml', shapes: [first] },
      },
    ]);
    const overlay = host.querySelector('.pptx-surface__overlay')!;
    const peers = [...overlay.querySelectorAll('.pptx-surface__presence')];
    expect(peers.map((p) => p.getAttribute('data-actor'))).toEqual(['p1', 'p2']);
    expect(peers[0].getAttribute('style')).toBe('--pptx-remote-color: #ff0000');
    expect(peers[0].querySelector('.pptx-surface__remote-outline')).not.toBeNull();
    expect(peers[1].getAttribute('style')).toBeNull();
    expect(peers[1].querySelector('.pptx-surface__remote-caret')).not.toBeNull();
    expect(peers[1].querySelector('.pptx-surface__remote-selection')).not.toBeNull();
    expect(
      peers.map((p) => p.querySelector('.pptx-surface__remote-label text')?.textContent)
    ).toEqual(['Ada', 'Grace']);
    // The label is drawn in screen pixels and scaled to slide units. A font size in EMU is
    // past the largest size a browser draws, so the name would show as a speck.
    for (const p of peers) {
      const tag = p.querySelector('.pptx-surface__remote-label')!;
      expect(tag.getAttribute('transform')).toMatch(/^translate\([^)]+\) scale\([^)]+\)$/);
      expect(tag.querySelector('text')!.getAttribute('font-size')).toBe('11');
    }
    const slide = host.querySelector('.pptx-surface__slide')!;
    expect(slide.querySelector('.pptx-surface__presence')).toBeNull();

    fake.setRemote([]);
    expect(overlay.querySelectorAll('.pptx-surface__presence')).toHaveLength(0);
    editor.destroy();
    if (!had) delete g.ResizeObserver;
  });
});
