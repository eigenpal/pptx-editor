import { describe, expect, test } from 'bun:test';
import type { WebrtcProvider } from 'y-webrtc';
import { chunkWebrtcMessages } from './webrtc-chunks.ts';

/** A stand-in for one simple-peer: `write` queues through `_write`, which sends raw frames. */
function fakePeer() {
  const sent: Uint8Array[] = [];
  const received: Uint8Array[] = [];
  const peer = {
    connected: false,
    destroyed: false,
    send(data: Uint8Array) {
      sent.push(data);
    },
    write(data: Uint8Array) {
      peer._write(data, 'buffer', () => {});
      return true;
    },
    _write(chunk: Uint8Array, _encoding: string, done: () => void) {
      peer.send(chunk);
      done();
    },
    emit(event: string | symbol, ...args: unknown[]) {
      if (event === 'data') received.push(args[0] as Uint8Array);
      return true;
    },
    destroy() {
      peer.destroyed = true;
    },
  };
  return { peer, sent, received };
}

/** Patch a fake peer as `chunkWebrtcMessages` patches each new connection. */
async function chunked() {
  const conns = new Map<string, { peer: unknown }>();
  const provider = { key: Promise.resolve(null), room: { webrtcConns: conns } };
  chunkWebrtcMessages(provider as unknown as WebrtcProvider);
  await provider.key;
  const fake = fakePeer();
  conns.set('peer', { peer: fake.peer });
  return fake;
}

const bytes = (length: number) => Uint8Array.from({ length }, (_, i) => i % 251);

describe('chunkWebrtcMessages', () => {
  test('a send before the channel opens throws and queues nothing', async () => {
    const { peer, sent } = await chunked();
    expect(() => peer.send(bytes(10))).toThrow();
    expect(sent).toHaveLength(0);
  });

  test('a large message travels in frames and arrives whole, once', async () => {
    const { peer, sent } = await chunked();
    peer.connected = true;
    const message = bytes(40_000);
    peer.send(message);
    expect(sent.map((f) => f.length)).toEqual([16_385, 16_385, 7_233]);
    expect(sent.map((f) => f[0])).toEqual([0, 0, 1]);

    const receiver = await chunked();
    for (const frame of sent) receiver.peer.emit('data', frame);
    expect(receiver.received).toHaveLength(1);
    expect(receiver.received[0]).toEqual(message);
  });

  test('an empty message and an exact multiple end on a final frame', async () => {
    const { peer, sent } = await chunked();
    peer.connected = true;
    peer.send(new Uint8Array(0));
    peer.send(bytes(16_384));
    expect(sent.map((f) => [f.length, f[0]])).toEqual([
      [1, 1],
      [16_385, 1],
    ]);
  });

  test('a closed channel drops a half-received message', async () => {
    const { peer, received } = await chunked();
    peer.emit('data', Uint8Array.of(0, 1, 2));
    peer.emit('close');
    peer.emit('data', Uint8Array.of(1, 9));
    expect(received).toEqual([Uint8Array.of(9)]);
  });
});
