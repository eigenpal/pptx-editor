// Send large Yjs messages over `y-webrtc`.
//
// `y-webrtc` sends each sync message as one WebRTC data channel message, and browsers refuse
// messages over about 256 KiB. A presentation room is larger than that: the sample deck's
// first sync is about 500 KiB. The send fails silently and the joining peer never syncs.
//
// This splits each outgoing message into 16 KiB chunks, a size that every WebRTC stack
// accepts, and joins them on arrival. Each chunk starts with one byte: 0 when more chunks
// follow, 1 on the last chunk. A data channel is
// ordered and reliable by default, so chunks arrive in order. Every peer of a room must use
// this, because it changes what travels on the channel. Tabs that sync through the
// BroadcastChannel are not affected.

import type { WebrtcProvider } from 'y-webrtc';

const CHUNK_BYTES = 16 * 1024;
/** A peer that sends more than this in one message is dropped. */
const MAX_MESSAGE_BYTES = 64 * 1024 * 1024;

/** The `simple-peer` instance of one `y-webrtc` connection. */
interface PeerLike {
  /** True once the data channel is open. */
  readonly connected: boolean;
  send(data: Uint8Array): void;
  /** The stream write: queued, and held while the channel's send buffer is full. */
  write(data: Uint8Array): boolean;
  _write(chunk: Uint8Array, encoding: string, done: (error?: Error | null) => void): void;
  emit(event: string | symbol, ...args: unknown[]): boolean;
  destroy(): void;
}

function chunkPeer(peer: PeerLike): void {
  // Frames go through the stream's write queue, not straight to the channel: a large deck is
  // many frames, and a browser refuses a send while its buffer is full (about 16 MiB).
  const rawSend = peer.send.bind(peer);
  const chunkedSend = (data: Uint8Array) => {
    // Before the channel opens, refuse as the raw send does: `y-webrtc` drops the message and
    // syncs in full on connect. A queued frame would be sent again, framed twice, on connect.
    if (!peer.connected) throw new Error('The data channel is not open.');
    for (let start = 0; start < data.length || start === 0; start += CHUNK_BYTES) {
      const part = data.subarray(start, start + CHUNK_BYTES);
      const frame = new Uint8Array(part.length + 1);
      frame[0] = start + CHUNK_BYTES >= data.length ? 1 : 0;
      frame.set(part, 1);
      peer.write(frame);
    }
  };
  peer.send = chunkedSend;
  // The queue sends each frame with `this.send`: give it the raw send for that call.
  const rawWrite = peer._write.bind(peer);
  peer._write = (chunk, encoding, done) => {
    peer.send = rawSend;
    try {
      rawWrite(chunk, encoding, done);
    } finally {
      peer.send = chunkedSend;
    }
  };
  let parts: Uint8Array[] = [];
  let size = 0;
  const emit = peer.emit.bind(peer);
  peer.emit = (event, ...args) => {
    if (event === 'close' || event === 'error') {
      parts = [];
      size = 0;
    }
    if (event !== 'data') return emit(event, ...args);
    const frame = args[0];
    if (!(frame instanceof Uint8Array) || frame.length === 0) return false;
    parts.push(frame.subarray(1));
    size += frame.length - 1;
    if (size > MAX_MESSAGE_BYTES) {
      peer.destroy();
      return false;
    }
    if (frame[0] !== 1) return true;
    const whole = new Uint8Array(size);
    let offset = 0;
    for (const part of parts) {
      whole.set(part, offset);
      offset += part.length;
    }
    parts = [];
    size = 0;
    return emit('data', whole);
  };
}

/** Chunk every WebRTC connection that `provider` opens from now on. */
export function chunkWebrtcMessages(provider: WebrtcProvider): void {
  // The provider opens its room when its key is ready, and peers connect after that.
  void provider.key.then(() => {
    const conns = provider.room?.webrtcConns;
    if (!conns) return;
    const set = conns.set.bind(conns);
    conns.set = (id, conn) => {
      chunkPeer(conn.peer as PeerLike);
      return set(id, conn);
    };
  });
}
