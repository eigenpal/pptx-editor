import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { canonicalPackageFingerprint } from '../../store/index.ts';
import {
  collaborationPort,
  createPptxEditor,
  resolveEditorModules,
  type CollaborationDocumentPort,
  type CollaborationLocalTransaction,
  type EditorCollaborationSession,
  type EditorModule,
} from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Hello')) }),
      },
    ],
  });

/** A session that records what the editor gives it. Transport is out of scope here. */
function recordingSession(): EditorCollaborationSession & {
  readonly ports: CollaborationDocumentPort[];
  readonly sent: CollaborationLocalTransaction[];
  detached: number;
} {
  const ports: CollaborationDocumentPort[] = [];
  const sent: CollaborationLocalTransaction[] = [];
  const state = {
    documentId: 'doc',
    sessionId: 'session',
    identity: { actorId: 'a', name: 'A' },
    ports,
    sent,
    detached: 0,
    attached: false,
    status: () => 'ready' as const,
    statusSnapshot: () => ({
      status: 'ready' as const,
      reason: undefined,
      lastFailure: undefined,
      failureCount: 0,
      recovering: false,
    }),
    subscribeStatus: () => () => {},
    participants: () => [],
    subscribeParticipants: () => () => {},
    destroy: () => {},
    attach(port: CollaborationDocumentPort) {
      ports.push(port);
      state.attached = true;
      const off = port.subscribeLocal((tx) => sent.push(tx));
      return () => {
        off();
        state.attached = false;
        state.detached += 1;
      };
    },
  };
  return state;
}

describe('editor modules', () => {
  test('an empty module list resolves to an empty registry', () => {
    expect(resolveEditorModules(undefined)).toEqual({ review: null, collaboration: null });
    expect(resolveEditorModules([])).toEqual({ review: null, collaboration: null });
  });

  test('the first module of each kind wins', () => {
    const first = { displayModes: ['markup'] as const, collectReviewItems: () => [] };
    const second = { displayModes: ['final'] as const, collectReviewItems: () => [] };
    const modules: EditorModule[] = [
      { id: 'one', review: first },
      { id: 'two', review: second },
    ];
    expect(resolveEditorModules(modules).review).toBe(first);
  });

  test('the editor attaches the collaboration session on load and detaches on destroy', async () => {
    const session = recordingSession();
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [{ id: 'collaboration', collaboration: { session } }],
    });
    expect((await editor.load(deck())).ok).toBe(true);
    expect(session.ports.length).toBe(1);
    expect(session.attached).toBe(true);

    // Reopening detaches from the old document before it attaches to the new one.
    expect((await editor.load(deck())).ok).toBe(true);
    expect(session.ports.length).toBe(2);
    expect(session.detached).toBe(1);

    editor.destroy();
    expect(session.detached).toBe(2);
    expect(session.attached).toBe(false);
  });

  test('local ops sent to a peer give the peer the same document', async () => {
    const session = recordingSession();
    const local = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [{ id: 'collaboration', collaboration: { session } }],
    });
    const peer = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = deck();
    expect((await local.load(bytes)).ok).toBe(true);
    expect((await peer.load(bytes)).ok).toBe(true);

    expect(local.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(local.exec({ type: 'insertTextBox' }).ok).toBe(true);
    expect(session.sent.length).toBeGreaterThan(0);

    const port = collaborationPort(peer.session!.store);
    for (const tx of session.sent) expect(port.applyRemote(tx.ops ?? []).ok).toBe(true);
    expect(canonicalPackageFingerprint(peer.session!.store.package)).toBe(
      canonicalPackageFingerprint(local.session!.store.package)
    );
  });

  test('remote ops make no history entry and clear local undo', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const other = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = deck();
    expect((await editor.load(bytes)).ok).toBe(true);
    expect((await other.load(bytes)).ok).toBe(true);
    const sent: CollaborationLocalTransaction[] = [];
    collaborationPort(other.session!.store).subscribeLocal((tx) => sent.push(tx));
    expect(other.exec({ type: 'addSlide' }).ok).toBe(true);

    expect(editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
    expect(editor.session!.store.canUndo).toBe(true);
    const store = editor.session!.store;
    const changes: string[] = [];
    store.subscribe((change) => changes.push(change.origin));
    for (const tx of sent) expect(collaborationPort(store).applyRemote(tx.ops ?? []).ok).toBe(true);
    expect(changes).toEqual(sent.map(() => 'remote'));
    expect(store.canUndo).toBe(false);
    expect(store.canRedo).toBe(false);
  });

  test('with session history, the editor undo and redo go to the session', async () => {
    const session = recordingSession();
    const calls: string[] = [];
    let undoable = false;
    const withHistory = Object.assign(session, {
      canUndo: () => undoable,
      canRedo: () => false,
      undo: () => (calls.push('undo'), true),
      redo: () => (calls.push('redo'), false),
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [{ id: 'collaboration', collaboration: { session: withHistory } }],
    });
    expect((await editor.load(deck())).ok).toBe(true);
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    // The store has history, but the session decides.
    expect(editor.session!.store.canUndo).toBe(true);
    expect(editor.can({ type: 'undo' }).ok).toBe(false);
    undoable = true;
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(false);
    expect(calls).toEqual(['undo', 'redo']);
    expect(editor.session!.info.slides.length).toBe(2);
  });

  test('a refused remote transaction is returned, not dropped', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    const r = collaborationPort(editor.session!.store).applyRemote([
      { op: 'removeNode', part: '/ppt/slides/slide1.xml', node: 'no-such-node' },
    ]);
    expect(r.ok).toBe(false);
  });
});
