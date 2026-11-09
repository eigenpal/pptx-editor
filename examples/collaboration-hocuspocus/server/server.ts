// A Hocuspocus room server for PPTX Editor collaboration.
//
// The server treats each room as an opaque Y.Doc. It checks a shared token, creates a room only
// for the tab that asks to start it, and stores each room in `server/.data/` so it survives a
// restart. The tab that starts a room seeds it from its deck; everyone else joins.
//
// Run it with Node.js 22.18 or later: `node server/server.ts`. Node strips the types.

import { access, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { Server } from '@hocuspocus/server';
import * as Y from 'yjs';
import { readRoomToken } from '../shared/token.ts';

const PORT = Number(process.env.PORT ?? 1234);

/**
 * The shared secret that the app sends. A real server verifies a signed token instead, and
 * derives the user and the room permissions from it.
 */
const SECRET = process.env.COLLAB_TOKEN ?? 'demo-token';

const DATA_DIR = path.join(import.meta.dirname, '.data');

/** The app's room ids. Anything else never reaches the file system. */
const ROOM_ID = /^[a-z0-9]{16,64}$/;

/** Rooms that a tab started in this run, before their first store. */
const started = new Set<string>();
/** At most this many started rooms wait for their first store. */
const MAX_STARTED = 100;

let stores = 0;

const roomFile = (room: string): string => path.join(DATA_DIR, `${room}.ydoc`);

const exists = (file: string): Promise<boolean> =>
  access(file).then(
    () => true,
    () => false
  );

/** A refusal that reaches the client as `onAuthenticationFailed({ reason })`. */
function refuse(reason: string): Error {
  return Object.assign(new Error(reason), { reason });
}

const server = new Server({
  port: PORT,
  // This machine only. Set HOST=0.0.0.0, and a real token, to accept other devices.
  address: process.env.HOST ?? '127.0.0.1',
  // Messages are bounded before authentication. A deck room's first sync is well under this.
  websocketOptions: { maxPayload: 32 * 1024 * 1024 },
  name: 'pptx-editor-collaboration',
  // Store a room at most two seconds after its last change.
  debounce: 2000,

  // Every connection waits for this, so nothing syncs before the server admits the client.
  async onAuthenticate({ token, documentName }) {
    const claim = readRoomToken(token);
    if (!claim || claim.secret !== SECRET) throw refuse('unauthorized');
    if (!ROOM_ID.test(documentName)) throw refuse('unknown-room');
    if (started.has(documentName) || (await exists(roomFile(documentName)))) return;
    // An invite link only joins. Only the starting tab can create a room.
    if (!claim.createRoom) throw refuse('unknown-room');
    if (started.size >= MAX_STARTED) throw refuse('too-many-rooms');
    started.add(documentName);
  },

  // A started room that closed before anyone seeded it is not a room: forget it.
  async afterUnloadDocument({ documentName }) {
    if (!(await exists(roomFile(documentName)))) started.delete(documentName);
  },

  async onLoadDocument({ document, documentName }) {
    try {
      Y.applyUpdate(document, new Uint8Array(await readFile(roomFile(documentName))));
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
    }
    return document;
  },

  // Write a temporary file, then rename it, so a crash leaves the old state or the new one.
  // Each store has its own temporary file, so two overlapping stores never mix.
  async onStoreDocument({ document, documentName }) {
    await mkdir(DATA_DIR, { recursive: true });
    const file = roomFile(documentName);
    const temporary = `${file}.${process.pid}.${++stores}.tmp`;
    await writeFile(temporary, Y.encodeStateAsUpdate(document));
    await rename(temporary, file);
  },
});

await server.listen();
console.log(`Room server listening on ws://127.0.0.1:${PORT}`);
