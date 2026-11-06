import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { prototypeChanges, shadowingCopies, snapshotPrototypes } from './dom-prototypes.ts';

// The guards of `preload-dom.ts`, on a private happy-dom window (this file has no global DOM).
type Classes = { Element: typeof Element; HTMLElement: typeof HTMLElement };

describe('the DOM prototype guards', () => {
  test('a restore by assignment leaves a copy that hides later patches; the guard deletes it', () => {
    const win = new Window();
    const { Element: El, HTMLElement: Html } = win as unknown as Classes;
    const before = snapshotPrototypes(win);
    const own = Object.getOwnPropertyDescriptor(El.prototype, 'getBoundingClientRect')!;
    const patch = () => ({ top: 7 }) as DOMRect;
    // The copy hides a later patch of Element.prototype from every HTML element.
    Html.prototype.getBoundingClientRect = Html.prototype.getBoundingClientRect;
    El.prototype.getBoundingClientRect = patch;
    expect(win.document.createElement('div').getBoundingClientRect().top).not.toBe(7);
    Object.defineProperty(El.prototype, 'getBoundingClientRect', own);
    // At the end of the test that made it, the guard finds the copy and deletes it.
    expect(shadowingCopies(win, before)).toEqual(['HTMLElement.prototype.getBoundingClientRect']);
    El.prototype.getBoundingClientRect = patch;
    expect(win.document.createElement('div').getBoundingClientRect().top).toBe(7);
    Object.defineProperty(El.prototype, 'getBoundingClientRect', own);
    expect(prototypeChanges(win, before)).toEqual([]);
    void win.happyDOM.close();
  });

  test('a patch that is in effect is no copy; the end-of-run check reports and restores it', () => {
    const win = new Window();
    const proto = (win as unknown as Classes).HTMLElement.prototype;
    const before = snapshotPrototypes(win);
    const focus = proto.focus;
    proto.focus = () => undefined;
    expect(shadowingCopies(win, before)).toEqual([]);
    expect(prototypeChanges(win, before)).toEqual(['HTMLElement.prototype.focus']);
    expect(proto.focus).toBe(focus);
    Reflect.deleteProperty(proto, 'focus');
    expect(prototypeChanges(win, before)).toEqual(['HTMLElement.prototype.focus']);
    expect(proto.focus).toBe(focus);
    void win.happyDOM.close();
  });
});
