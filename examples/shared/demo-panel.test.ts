// The demo's side panel and speaker notes: their open state at the start, and what the demo
// keeps for the next visit. A narrow window shows the panel as a sheet that starts closed.
import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { notesStartOpen, panelStartsOpen, rememberNotes, rememberPanel } from './demo-files.ts';

const g = globalThis as unknown as Record<string, unknown>;
let narrow = false;
let store = new Map<string, string>();

beforeEach(() => {
  narrow = false;
  store = new Map();
  g.window = globalThis;
  g.matchMedia = () => ({ matches: narrow }) as MediaQueryList;
  g.localStorage = {
    getItem: (k: string) => store.get(k) ?? null,
    setItem: (k: string, v: string) => void store.set(k, v),
  };
});
afterEach(() => {
  delete g.window;
  delete g.matchMedia;
  delete g.localStorage;
});

describe('the demo panel and notes', () => {
  test('the panel starts open on a wide window and closed on a narrow one', () => {
    expect(panelStartsOpen()).toBe(true);
    narrow = true;
    expect(panelStartsOpen()).toBe(false);
  });

  test('the panel starts as the user left it on a wide window', () => {
    rememberPanel(false);
    expect(panelStartsOpen()).toBe(false);
    rememberPanel(true);
    narrow = true;
    expect(panelStartsOpen()).toBe(true);
  });

  test('closing the sheet on a narrow window keeps the wide choice', () => {
    rememberPanel(true);
    narrow = true;
    rememberPanel(false);
    narrow = false;
    expect(panelStartsOpen()).toBe(true);
  });

  test('the notes start closed, then as the user left them', () => {
    expect(notesStartOpen()).toBe(false);
    rememberNotes(true);
    expect(notesStartOpen()).toBe(true);
  });

  test('storage that throws gives the defaults', () => {
    g.localStorage = {
      getItem: () => {
        throw new Error('off');
      },
      setItem: () => {
        throw new Error('off');
      },
    };
    expect(() => rememberNotes(true)).not.toThrow();
    expect(notesStartOpen()).toBe(false);
    expect(panelStartsOpen()).toBe(true);
  });
});
