# Peer-to-peer PPTX collaboration

Edit one presentation from two browsers over `y-webrtc`. Browsers exchange edits directly; a local signaling server only introduces them. There is no room storage.

For server-backed rooms, see the [Hocuspocus collaboration example](../collaboration-hocuspocus/README.md).

## Run the example

From the repository root, start the signaling server in the first terminal:

```bash
bun install
bun run dev:collaboration:signal
```

Start the app in the second terminal:

```bash
bun run dev:collaboration
```

Open `http://localhost:5277`. The signaling server listens on `ws://localhost:4444`.

## Try it

1. Pick a person.
2. Select **Copy invite link** in the room bar.
3. Open the link in another browser or a private window, and pick a different person.
4. Edit a slide in one window. The other window shows the edit and your caret with your name.

A room lives only in its open tabs. When every tab of a room closes, the room is gone.

The invite key needs `crypto.subtle`, which browsers offer on `localhost` and on HTTPS only. To join from a second device, serve the app over HTTPS and set `VITE_SIGNALING_URL` to a signaling server that the device can reach.

## How it works

`src/App.tsx` connects a `Y.Doc` with `WebrtcProvider`, then creates a session with `createYjsCollaborationSession()` and passes `collaborationModule({ session })` to the editor. The tab that starts a room seeds it from the sample deck. A tab from an invite link waits for the room from a peer, so it never seeds a second copy.

A deck room is larger than one WebRTC data channel message, so `src/webrtc-chunks.ts` splits each message into chunks. The `#key=` fragment of the invite link encrypts the room; anyone with the full link can join. The example uses no STUN or TURN servers, so peers must reach each other directly.

For more information, see [Collaboration](../../docs/site/content/pro/collaboration.mdx).
