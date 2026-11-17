// The editor's listener relay: synchronous notifications, and one deferred notification for a
// refresh commit, which keeps the cause of the change and the order of states.

import { describe, expect, test } from 'bun:test';
import { changeSourceOf, withChangeSource } from '../change-source.ts';
import { createListenerRelay, deferListeners } from '../listener-relay.ts';

const nextTask = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

describe('listener relay', () => {
  test('notifies synchronously outside deferListeners', () => {
    const editor = {};
    const relay = createListenerRelay(() => editor);
    let calls = 0;
    relay.add(() => void calls++);
    relay.notify();
    expect(calls).toBe(1);
  });

  test('a deferred commit notifies once, in the next task, with its change source', async () => {
    const editor = {};
    const relay = createListenerRelay(() => editor);
    const seen: (string | undefined)[] = [];
    relay.add(() => void seen.push(changeSourceOf(editor)));
    withChangeSource(editor, 'refresh', () =>
      deferListeners(editor, () => {
        relay.notify();
        relay.notify();
      })
    );
    expect(seen).toEqual([]);
    await nextTask();
    expect(seen).toEqual(['refresh']);
    expect(changeSourceOf(editor)).toBeUndefined();
  });

  test('a notification due while one waits delivers the waiting one first', async () => {
    const editor = {};
    const relay = createListenerRelay(() => editor);
    const seen: (string | undefined)[] = [];
    relay.add(() => void seen.push(changeSourceOf(editor) ?? 'edit'));
    withChangeSource(editor, 'refresh', () => deferListeners(editor, () => relay.notify()));
    relay.notify();
    expect(seen).toEqual(['refresh', 'edit']);
    await nextTask();
    expect(seen).toEqual(['refresh', 'edit']);
  });

  test('deferring one editor leaves another synchronous', () => {
    const a = {};
    const b = {};
    const relayA = createListenerRelay(() => a);
    const relayB = createListenerRelay(() => b);
    let heardA = 0;
    let heardB = 0;
    relayA.add(() => void heardA++);
    relayB.add(() => void heardB++);
    deferListeners(a, () => {
      relayA.notify();
      relayB.notify();
    });
    expect([heardA, heardB]).toEqual([0, 1]);
  });

  test('clear drops a waiting notification and the listeners', async () => {
    const editor = {};
    const relay = createListenerRelay(() => editor);
    let calls = 0;
    const off = relay.add(() => void calls++);
    deferListeners(editor, () => relay.notify());
    relay.clear();
    await nextTask();
    relay.notify();
    expect(calls).toBe(0);
    off();
  });
});
