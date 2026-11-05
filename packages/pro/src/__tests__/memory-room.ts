// A test-only op-based room: every attached replica sends its local ops to the others, in one
// total order. No transport and no conflicts (the tests edit one replica at a time). The real
// binding is state-based: see `../collaboration/yjs/`.

import type {
  CollaborationDocumentPort,
  CollaborationIdentity,
  CollaborationParticipant,
  EditorCollaborationSession,
} from '@pptx-editor.dev/core/editor';

const READY = Object.freeze({
  status: 'ready' as const,
  reason: undefined,
  lastFailure: undefined,
  failureCount: 0,
  recovering: false,
});

export function memoryRoom(): {
  session(identity: CollaborationIdentity): EditorCollaborationSession;
} {
  const ports = new Set<CollaborationDocumentPort>();
  const identities: CollaborationIdentity[] = [];
  let applying = false;
  return {
    session(identity) {
      identities.push(identity);
      let attached = false;
      return {
        documentId: 'memory',
        sessionId: identity.actorId,
        identity,
        get attached() {
          return attached;
        },
        status: () => 'ready',
        statusSnapshot: () => READY,
        subscribeStatus: () => () => {},
        participants: (): readonly CollaborationParticipant[] =>
          identities.map((p) => ({ ...p, isLocal: p.actorId === identity.actorId })),
        subscribeParticipants: () => () => {},
        destroy: () => {},
        attach(port) {
          ports.add(port);
          attached = true;
          const off = port.subscribeLocal((tx) => {
            // An op-based room: undo and redo carry no ops, so it sends edits only.
            if (applying || !tx.ops) return;
            applying = true;
            try {
              for (const peer of ports) if (peer !== port) peer.applyRemote(tx.ops, tx.label);
            } finally {
              applying = false;
            }
          });
          return () => {
            off();
            ports.delete(port);
            attached = false;
          };
        },
      };
    },
  };
}
