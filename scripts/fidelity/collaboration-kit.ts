// Two collaborating editors for the feature edits (`feature-edits.ts`): each opens the deck on
// its own Yjs replica through the collaboration session of `packages/pro`, and a link carries
// the updates between the replicas. The link can pause, to make edits of the two peers
// concurrent. AGENTS.md: "Verify collaboration convergence when changing editing behavior."

import { join } from 'node:path';
import { createPptxEditor } from '../../packages/core/src/editor/index.ts';
import {
  collaborationModule,
  createYjsCollaborationSession,
} from '../../packages/pro/src/index.ts';
import type { OpenPeers, PeerLink } from './feature-edit-kit.ts';

/** The Yjs document type of the session. */
type Doc = Parameters<typeof createYjsCollaborationSession>[0]['doc'];

/** The part of yjs that this kit uses. */
interface Yjs {
  readonly Doc: new () => Doc;
  applyUpdate(doc: Doc, update: Uint8Array, origin?: unknown): void;
}

// The session and this kit must share one copy of yjs: resolve its ES module entry as the pro
// package imports it. yjs is a dependency of the pro package only.
const Y = (await import(
  Bun.resolveSync('yjs', join(import.meta.dir, '../../packages/pro/src'))
)) as Yjs;

/** Open `bytes` in two editors on two linked replicas, with fixed client ids. */
export const openPeers: OpenPeers = async (bytes) => {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  docs[0].clientID = 1_001;
  docs[1].clientID = 2_002;
  const queue: [Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  docs[0].on('update', send(docs[1]));
  docs[1].on('update', send(docs[0]));
  const link: PeerLink = {
    pause: () => {
      paused = true;
    },
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
  const editors = [];
  for (const [i, doc] of docs.entries()) {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'feature-edit',
      identity: { actorId: `peer-${i}`, name: `Peer ${i + 1}` },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    const result = await editor.load(bytes);
    if (!result.ok) throw new Error(`peer ${i + 1} did not open: ${JSON.stringify(result)}`);
    editors.push(editor);
  }
  return { editor: editors[0]!, peer: editors[1]!, link };
};
