# Server-backed PPTX collaboration

Edit one presentation from two browsers through a [Hocuspocus](https://tiptap.dev/docs/hocuspocus) room server. The server admits each connection and stores every room, so a room survives a restart.

For rooms with no server, see the [peer-to-peer collaboration example](../collaboration/README.md).

## Run the example

Use Node.js 22.18 or later for the server. From the repository root, start the server in the first terminal:

```bash
bun install
bun run dev:collaboration-hocuspocus:server
```

Start the app in the second terminal:

```bash
bun run dev:collaboration-hocuspocus
```

Open `http://localhost:5278`. The server listens on `ws://127.0.0.1:1234`.

## Try it

1. Pick a person.
2. Select **Copy invite link** in the room bar.
3. Open the link in another browser or a private window, and pick a different person.
4. Edit a slide in one window. The other window shows the edit and your caret with your name.
5. Stop and start the server, then open the invite link again. The room keeps your edits.

## How it works

`src/App.tsx` connects a `Y.Doc` with `HocuspocusProvider`, then creates a session with `createYjsCollaborationSession()` and passes `collaborationModule({ session })` to the editor. The tab that starts a room asks the server to create it and seeds it from the sample deck. An invite link only joins: `server/server.ts` refuses a room that it does not know.

The server checks the shared token `demo-token`. Set `COLLAB_TOKEN` on the server and `VITE_COLLAB_TOKEN` in the app to change it, and `VITE_COLLAB_URL` to move the server. A real server verifies a signed token and derives the user and room permissions from it. Rooms are stored in `server/.data/`.

For more information, see [Collaboration](../../docs/site/content/pro/collaboration.mdx).
