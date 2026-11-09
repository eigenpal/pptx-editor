// Peer-to-peer collaboration over `y-webrtc`. Browsers exchange Yjs updates directly; the
// signaling server only introduces them. Tabs of one browser also find each other through a
// BroadcastChannel, with no signaling at all.
//
// The tab that starts a room seeds it from the sample deck. A tab opened from an invite link
// waits for the room's state from a peer before it creates its session, so it adopts the room
// instead of seeding a second copy.

import { useEffect, useRef, useState } from 'react';
import * as Y from 'yjs';
import { WebrtcProvider } from 'y-webrtc';
import { packagedFonts } from '@docx-editor.dev/fonts';
import { createYjsCollaborationSession, type YjsCollaborationSession } from '@pptx-editor.dev/pro';
import { usePptxSource } from '@pptx-editor.dev/react';
import {
  actorIdFor,
  holdsRoom,
  createRoomId,
  randomHex,
  roomFromUrl,
  roomUrl,
  shortRoomId,
  type Person,
} from '../../shared/collaboration/room.ts';
import { RoomEditor, RoomMessage, SeatPicker } from '../../shared/collaboration/RoomEditor.tsx';
import { chunkWebrtcMessages } from './webrtc-chunks.ts';

/** The signaling server that `bun run dev:collaboration:signal` starts. */
const SIGNALING_URL = import.meta.env.VITE_SIGNALING_URL ?? 'ws://localhost:4444';

/** The encryption key of an invite link, from its `#key=` fragment. Browsers never send it. */
function keyFromUrl(): string | null {
  const key = new URLSearchParams(location.hash.slice(1)).get('key');
  return key && /^[0-9a-f]{32}$/.test(key) ? key : null;
}

/** How long a found peer may take to connect before the waiting screen says why it may fail. */
const STUCK_MS = 10_000;

/** How far this tab got in finding and reaching a peer. */
interface Link {
  /** The signaling server answers. */
  readonly signaling: boolean;
  /** Peers found through signaling or the BroadcastChannel. */
  readonly found: number;
  /** Peers with an open channel. */
  readonly linked: number;
  /** A peer was found but has not connected for `STUCK_MS`. */
  readonly stuck: boolean;
}

/** The waiting screen's explanation of `link`. */
function waitingText(link: Link): string {
  if (link.linked > 0) return 'Connected to a peer. Receiving the presentation…';
  if (link.stuck)
    return 'A peer was found, but WebRTC could not connect to it. A browser setting or extension that blocks WebRTC or hides local addresses can cause this. Try a private window, where extensions are off.';
  if (link.found > 0) return 'Found a peer. Connecting over WebRTC…';
  if (!link.signaling)
    return `The signaling server at ${SIGNALING_URL} does not answer. Start it with bun run dev:collaboration:signal.`;
  return 'No other tab of this room is open. Open the room in the tab that started it, or start a new room.';
}

interface Joined {
  readonly person: Person;
  readonly roomId: string;
  readonly key: string;
  /** This tab started the room, so its deck seeds it. */
  readonly creating: boolean;
}

export function App() {
  const [joined, setJoined] = useState<Joined | null>(null);
  const [session, setSession] = useState<YjsCollaborationSession | null>(null);
  const [failure, setFailure] = useState<string | null>(null);
  const [link, setLink] = useState<Link>({ signaling: false, found: 0, linked: 0, stuck: false });
  const source = usePptxSource('/sample.pptx', { fonts: packagedFonts() });
  // Rooms this tab seeded. A remount (Fast Refresh, for one) then joins instead of seeding the
  // room a second time from a new, empty Y.Doc.
  const seeded = useRef(new Set<string>());

  useEffect(() => {
    if (!joined) return undefined;
    const { person, roomId, key } = joined;
    const creating = joined.creating && !seeded.current.has(roomId);
    const doc = new Y.Doc();
    // `password` encrypts every message between peers, and the signaling topic is hashed. The
    // peers use y-webrtc's default public STUN servers to find a route. A browser that hides its
    // local addresses, as some privacy settings and extensions do, needs one.
    const provider = new WebrtcProvider(`pptx-editor-${roomId}`, doc, {
      signaling: [SIGNALING_URL],
      password: key,
    });
    // A deck room is larger than one data channel message.
    chunkWebrtcMessages(provider);
    let created: YjsCollaborationSession | null = null;
    // Connected while the signaling server answers or a peer is linked. Edits continue offline
    // and merge when the peers meet again.
    let foundAt = 0;
    const report = () => {
      const room = provider.room;
      const found = room ? room.webrtcConns.size + room.bcConns.size : 0;
      const linked = room
        ? [...room.webrtcConns.values()].filter((conn) => conn.connected).length + room.bcConns.size
        : 0;
      const signaling = provider.signalingConns.some(
        (conn: { connected: boolean }) => conn.connected
      );
      created?.setConnected(signaling || linked > 0);
      // What the waiting screen says: how far the search for a peer got.
      if (found === 0) foundAt = 0;
      else if (foundAt === 0) foundAt = Date.now();
      const stuck = found > 0 && linked === 0 && Date.now() - foundAt > STUCK_MS;
      setLink((was) =>
        was.signaling === signaling &&
        was.found === found &&
        was.linked === linked &&
        was.stuck === stuck
          ? was
          : { signaling, found, linked, stuck }
      );
    };
    // A peer link opens without an event that this example hears, so look again each second.
    const poll = setInterval(report, 1000);
    const start = () => {
      if (created) return;
      try {
        created = createYjsCollaborationSession({
          doc,
          awareness: provider.awareness,
          documentId: roomId,
          identity: { actorId: actorIdFor(person), name: person.name, color: person.color },
        });
      } catch (error) {
        // A room that the session cannot read, for example one that a hostile peer wrote.
        setFailure(error instanceof Error ? error.message : String(error));
        return;
      }
      seeded.current.add(roomId);
      report();
      setSession(created);
      // The room exists now: a reload or a copied address joins it.
      history.replaceState(null, '', roomUrl(roomId, `key=${key}`));
    };
    provider.on('peers', report);
    // The signaling sockets open once the room key is ready, and they report their own state.
    type SignalingSocket = {
      on(event: string, f: () => void): void;
      off(event: string, f: () => void): void;
    };
    let sockets: SignalingSocket[] = [];
    void provider.key.then(() => {
      sockets = [...provider.signalingConns];
      for (const socket of sockets) {
        socket.on('connect', report);
        socket.on('disconnect', report);
      }
      report();
    });
    const roomArrived = () => {
      if (holdsRoom(doc)) start();
    };
    if (creating) start();
    else doc.on('update', roomArrived);
    return () => {
      clearInterval(poll);
      doc.off('update', roomArrived);
      for (const socket of sockets) {
        socket.off('connect', report);
        socket.off('disconnect', report);
      }
      provider.destroy();
      setSession(null);
      // After the editor unmounts and detaches from it.
      setTimeout(() => {
        created?.destroy();
        doc.destroy();
      });
    };
  }, [joined]);

  if (!joined) {
    const invited = roomFromUrl();
    const key = keyFromUrl();
    return (
      <SeatPicker
        title="Peer-to-peer collaboration"
        lede={
          invited && key
            ? 'Pick a person that nobody in this room uses.'
            : 'Pick a person, then open the invite link in a second browser as somebody else.'
        }
        onPick={(person) =>
          setJoined(
            invited && key
              ? { person, roomId: invited, key, creating: false }
              : { person, roomId: createRoomId(), key: randomHex(16), creating: true }
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

  if (!session) {
    return (
      <RoomMessage title={`Waiting for a peer in room ${shortRoomId(joined.roomId)}`}>
        <p>{waitingText(link)}</p>
        <p>
          If every tab of this room closed, the room is gone.{' '}
          <a href={location.pathname}>Start a new room</a>.
        </p>
      </RoomMessage>
    );
  }

  return (
    <RoomEditor
      session={session}
      source={source}
      roomId={joined.roomId}
      inviteUrl={roomUrl(joined.roomId, `key=${joined.key}`)}
      transport={`WebRTC, signaling ${SIGNALING_URL}`}
    />
  );
}
