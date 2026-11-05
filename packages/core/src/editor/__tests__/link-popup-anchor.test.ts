// The link popup opens at its own editor's open menu bar trigger: with two editors on one page,
// the other editor's open menu never anchors it.

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { linkPopupAnchor } from '../index.ts';
import { setSurfaceRoot } from '../surface-registry.ts';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;

/** One editor scope with a menu bar trigger and a slide surface, laid out. */
function editorScope(open: boolean): { trigger: HTMLElement; surface: HTMLElement } {
  const scope = document.createElement('div');
  scope.className = 'pptx-editor';
  scope.innerHTML =
    `<div class="pptx-menubar"><button class="pptx-menubar__trigger" aria-expanded="${open}">` +
    'Insert</button></div><div class="pptx-surface"></div>';
  document.body.appendChild(scope);
  const trigger = scope.querySelector<HTMLElement>('.pptx-menubar__trigger')!;
  // happy-dom lays nothing out: the trigger is shown.
  trigger.getClientRects = () => [{}] as unknown as DOMRectList;
  return { trigger, surface: scope.querySelector<HTMLElement>('.pptx-surface')! };
}

describe('linkPopupAnchor', () => {
  test("opens at this editor's trigger, never at another editor's open menu", () => {
    const first = editorScope(true);
    const second = editorScope(false);
    const editorA = {};
    const editorB = {};
    setSurfaceRoot(editorA, first.surface);
    setSurfaceRoot(editorB, second.surface);
    expect(linkPopupAnchor({ editor: editorA }, editorA)).toBe(first.trigger);
    // B has no open menu of its own: its slide anchors the popup.
    expect(linkPopupAnchor({ editor: editorB }, editorB)).toBe(second.surface);
    setSurfaceRoot(editorA, null);
    setSurfaceRoot(editorB, null);
  });
});
