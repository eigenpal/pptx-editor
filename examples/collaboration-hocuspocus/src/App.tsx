// Server-backed collaboration over Hocuspocus. `server/server.ts` admits each connection,
// relays Yjs updates and presence, and stores every room on disk.
//
// The tab that starts a room asks the server to create it, then seeds it from the sample deck
// after the first sync. A tab opened from an invite link only joins: it waits until the room
// holds a deck, so it adopts the room instead of seeding a second copy.

import { useEffect, useState } from 'react';
import * as Y from 'yjs';
import { HocuspocusProvider } from '@hocuspocus/provider';
import { packagedFonts } from '@docx-editor.dev/fonts';
import { createYjsCollaborationSession, type YjsCollaborationSession } from '@pptx-editor.dev/pro';
import { usePptxSource } from '@pptx-editor.dev/react';
import {
  actorIdFor,
  holdsRoom,
  createRoomId,
  roomFromUrl,
  roomUrl,
  shortRoomId,
  type Person,
} from '../../shared/collaboration/room.ts';
import { RoomEditor, RoomMessage, SeatPicker } from '../../shared/collaboration/RoomEditor.tsx';
import { encodeRoomToken } from '../shared/token.ts';

/** Where `server/server.ts` listens. */
const SERVER_URL = import.meta.env.VITE_COLLAB_URL ?? 'ws://127.0.0.1:1234';
/** The shared secret that the server checks. Set `COLLAB_TOKEN` on the server to match. */
const SECRET = import.meta.env.VITE_COLLAB_TOKEN ?? 'demo-token';
const SERVER_COMMAND = 'bun run dev:collaboration-hocuspocus:server';

interface Joined {
  readonly person: Person;
  readonly roomId: string;
  /** This tab started the room, so the server may create it and its deck seeds it. */
  readonly creating: boolean;
}

export function App() {
  const [joined, setJoined] = useState<Joined | null>(null);
  const [session, setSession] = useState<YjsCollaborationSession | null>(null);
  const [failure, setFailure] = useState<string | null>(null);
  const [refused, setRefused] = useState<string | null>(null);
  const [offline, setOffline] = useState(false);
  // Synced, but the room holds no deck yet: its starting tab has not seeded it.
  const [empty, setEmpty] = useState(false);
  const source = usePptxSource('/sample.pptx', { fonts: packagedFonts() });

  useEffect(() => {
    if (!joined) return undefined;
    const { person, roomId, creating } = joined;
    const doc = new Y.Doc();
    let created: YjsCollaborationSession | null = null;
    const start = () => {
      if (created) return;
      try {
        created = createYjsCollaborationSession({
          doc,
          awareness: provider.awareness ?? undefined,
          documentId: roomId,
          identity: { actorId: actorIdFor(person), name: person.name, color: person.color },
        });
      } catch (error) {
        // A room that the session cannot read, for example one that a hostile peer wrote.
        setFailure(error instanceof Error ? error.message : String(error));
        return;
      }
      setSession(created);
      // The room exists now: a reload or a copied address joins it.
      history.replaceState(null, '', roomUrl(roomId));
    };
    const roomArrived = () => {
      if (holdsRoom(doc)) start();
    };
    const provider = new HocuspocusProvider({
      url: SERVER_URL,
      name: roomId,
      document: doc,
      token: encodeRoomToken({ secret: SECRET, createRoom: creating }),
      onSynced: () => {
        if (creating) start();
        else roomArrived();
        setEmpty(!created);
      },
      onStatus: ({ status }) => {
        setOffline(status === 'disconnected');
        created?.setConnected(status === 'connected');
      },
      onAuthenticationFailed: ({ reason }) => setRefused(reason),
    });
    if (!creating) doc.on('update', roomArrived);
    return () => {
      doc.off('update', roomArrived);
      provider.destroy();
      setSession(null);
      setRefused(null);
      setEmpty(false);
      setOffline(false);
      // After the editor unmounts and detaches from it.
      setTimeout(() => {
        created?.destroy();
        doc.destroy();
      });
    };
  }, [joined]);

  if (!joined) {
    const invited = roomFromUrl();
    return (
      <SeatPicker
        title="Server-backed collaboration"
        lede={
          invited
            ? 'Pick a person that nobody in this room uses.'
            : 'Pick a person, then open the invite link in a second browser as somebody else.'
        }
        onPick={(person) =>
          setJoined(
            invited
              ? { person, roomId: invited, creating: false }
              : { person, roomId: createRoomId(), creating: true }
          )
        }
      />
    );
  }

  if (failure) {
    return (
      <RoomMessage title="This room cannot open" alert>
        <p>{failure}</p>
        <p>
          <a href={location.pathname}>Start a new room</a>
        </p>
      </RoomMessage>
    );
  }

  if (source.error) {
    return (
      <RoomMessage title="The sample deck did not load" alert>
        <p>{source.error.message}</p>
      </RoomMessage>
    );
  }

  if (refused) {
    return (
      <RoomMessage title="The server refused this room" alert>
        <p>
          {refused === 'unknown-room'
            ? 'This room does not exist on the server. Start a new room from the example address.'
            : `The server sent "${refused}". Check that COLLAB_TOKEN and VITE_COLLAB_TOKEN match.`}
        </p>
        <p>
          <a href={location.pathname}>Start a new room</a>
        </p>
      </RoomMessage>
    );
  }

  if (!session && empty) {
    return (
      <RoomMessage title={`Room ${shortRoomId(joined.roomId)} has no deck yet`}>
        <p>
          The tab that started this room has not added its deck. Keep that tab open, and this page
          opens the deck when it arrives.
        </p>
        <p>
          <a href={location.pathname}>Start a new room</a>
        </p>
      </RoomMessage>
    );
  }

  if (!session) {
    return (
      <RoomMessage title={`Opening room ${shortRoomId(joined.roomId)}`}>
        <p>
          {offline
            ? `The room server at ${SERVER_URL} does not answer. Start it with:`
            : `Waiting for ${SERVER_URL}. If this does not change, start the server with:`}
        </p>
        <p>
          <code>{SERVER_COMMAND}</code>
        </p>
      </RoomMessage>
    );
  }

  return (
    <RoomEditor
      session={session}
      source={source}
      roomId={joined.roomId}
      inviteUrl={roomUrl(joined.roomId)}
      transport={`Hocuspocus ${SERVER_URL}`}
    />
  );
}
