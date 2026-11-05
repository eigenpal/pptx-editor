// Bullets and numbering (`setBulletFormat`, the galleries, and the Bullets and Numbering
// dialog) through two editors on two linked Y.Docs, as AGENTS.md "Collaboration verification"
// asks: concurrent edits with the link paused, undo while a peer edits, and a deletion while a
// peer edits the same shape. The replicas converge, before and after save and reopen. One
// value is an attribute write; a pick that is one unit (a character in a new font, a new kind,
// a color with its luminance) is a new child, so two picks at once show one peer's whole pick.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  createPptxEditor,
  runToolbarCommand,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  readPptx,
  writePptx,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const LIST =
  '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text: para(run('one'), LIST) + para(run('two'), LIST),
        }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

async function replicas() {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  const sync = link(docs[0], docs[1]);
  const editors: PptxEditorInstance[] = [];
  for (const [i, doc] of docs.entries()) {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: `p${i}`, name: `p${i}` },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(deck())).ok).toBe(true);
    editors.push(editor);
  }
  return { a: editors[0]!, b: editors[1]!, sync };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
/** Both replicas match, live and after save and reopen. */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = (e: PptxEditorInstance) => {
    const reopened = readPptx(writePptx(e.session!.store.package));
    if (!reopened.ok) throw new Error('reopen failed');
    return canonicalPackageFingerprint(reopened.pkg);
  };
  expect(saved(a)).toBe(saved(b));
}
const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeOf = (e: PptxEditorInstance): Element | undefined =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T');
const pPrs = (e: PptxEditorInstance): Element[] =>
  elements(child(shapeOf(e)!, 'p:txBody')!)
    .filter((c) => c.local === 'p')
    .map((p) => child(p, 'a:pPr')!);
/** The bullet children of each paragraph, in order. */
const kids = (e: PptxEditorInstance): string[][] =>
  pPrs(e).map((p) => elements(p).map((c) => c.local));
function selectText(e: PptxEditorInstance): void {
  e.setSelection({
    kind: 'text',
    nodeId: shapeOf(e)!.id,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 1, offset: 3 },
  });
}

describe('Bullets and numbering under collaboration', () => {
  test('two peers pick different gallery bullets at once: one bullet, one child per slot', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    sync.pause();
    expect(runToolbarCommand(a, 'list.bulletStyle', 'hollowRound').ok).toBe(true);
    expect(runToolbarCommand(b, 'list.bulletStyle', 'star').ok).toBe(true);
    sync.flush();
    converged(a, b);
    for (const p of pPrs(a)) {
      expect(elements(p).map((c) => c.local)).toEqual(['buFont', 'buChar']);
      const pair = [attr(find(p, 'a:buFont'), 'typeface'), attr(find(p, 'a:buChar'), 'char')];
      expect([JSON.stringify(['Courier New', 'o']), JSON.stringify(['Wingdings', 'v'])]).toContain(
        JSON.stringify(pair)
      );
    }
  });

  test('two peers change the bullet color at once: one color', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    expect(a.exec({ type: 'setBulletFormat', format: { color: '#00B050' } }).ok).toBe(true);
    sync.pause();
    expect(a.exec({ type: 'setBulletFormat', format: { color: '#C00000' } }).ok).toBe(true);
    expect(b.exec({ type: 'setBulletFormat', format: { color: '#0070C0' } }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    for (const p of pPrs(a)) {
      const clr = find(p, 'a:buClr')!;
      expect(elements(clr).map((c) => c.local)).toEqual(['srgbClr']);
      expect(['C00000', '0070C0']).toContain(attr(find(clr, 'a:srgbClr'), 'val') ?? '');
    }
  });

  test('one peer sets the size and the other the color: both survive', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    sync.pause();
    expect(
      a.exec({ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 150 } }).ok
    ).toBe(true);
    expect(
      b.exec({ type: 'setBulletFormat', format: { kind: 'bullet', color: '#C00000' } }).ok
    ).toBe(true);
    sync.flush();
    converged(a, b);
    for (const p of pPrs(a)) {
      expect(attr(find(p, 'a:buSzPct'), 'val')).toBe('150000');
      expect(attr(find(p, 'a:srgbClr'), 'val')).toBe('C00000');
      expect(attr(find(p, 'a:buChar'), 'char')).toBe('•');
    }
    // Both children survive; their order follows the client ids (the second `todo` below).
    expect([...kids(a)[0]!].sort()).toEqual(['buChar', 'buClr', 'buFont', 'buSzPct']);
  });

  test("undo reverts this peer's change and keeps the other peer's", async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    sync.pause();
    expect(
      a.exec({ type: 'setBulletFormat', format: { kind: 'number', scheme: 'romanUcPeriod' } }).ok
    ).toBe(true);
    expect(
      b.exec({ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 80 } }).ok
    ).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    for (const p of pPrs(a)) {
      expect(find(p, 'a:buAutoNum')).toBeUndefined();
      expect(attr(find(p, 'a:buChar'), 'char')).toBe('•');
      expect(attr(find(p, 'a:buSzPct'), 'val')).toBe('80000');
    }
    // Redo brings this peer's numbering back on both replicas.
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(attr(find(pPrs(b)[0]!, 'a:buAutoNum'), 'type')).toBe('romanUcPeriod');
  });

  test('a peer deletes the shape while the other sets its bullets: both end without it', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    b.setSelection({ kind: 'shapes', nodeIds: [shapeOf(b)!.id] });
    sync.pause();
    expect(a.exec({ type: 'setBulletFormat', format: { kind: 'bullet', char: '–' } }).ok).toBe(
      true
    );
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(shapeOf(a)).toBeUndefined();
  });

  /** The bullet children of each paragraph with the values that pick the bullet. */
  const bullets = (e: PptxEditorInstance) =>
    pPrs(e).map((p) => ({
      kids: elements(p).map((c) => c.local),
      font: attr(find(p, 'a:buFont'), 'typeface'),
      char: attr(find(p, 'a:buChar'), 'char'),
      scheme: attr(find(p, 'a:buAutoNum'), 'type'),
    }));

  test("two peers add a paragraph's first bullet color at once: one child", async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    sync.pause();
    expect(a.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent1' } } }).ok).toBe(
      true
    );
    expect(b.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent2' } } }).ok).toBe(
      true
    );
    sync.flush();
    converged(a, b);
    for (const k of kids(a)) expect(k).toEqual(['buClr', 'buFont', 'buChar']);
  });

  test('two peers add a bullet size and a bullet color at once: schema order kept', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    selectText(b);
    sync.pause();
    expect(a.exec({ type: 'setBulletFormat', format: { sizePercent: 150 } }).ok).toBe(true);
    expect(b.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent2' } } }).ok).toBe(
      true
    );
    sync.flush();
    converged(a, b);
    for (const k of kids(a)) expect(k).toEqual(['buClr', 'buSzPct', 'buFont', 'buChar']);
  });

  test("two peers change the kind at once (numbers and a symbol bullet): one peer's whole bullet", async () => {
    for (let trial = 0; trial < 8; trial++) {
      const { a, b, sync } = await replicas();
      selectText(a);
      selectText(b);
      sync.pause();
      expect(
        a.exec({ type: 'setBulletFormat', format: { kind: 'number', scheme: 'romanUcPeriod' } }).ok
      ).toBe(true);
      expect(
        b.exec({
          type: 'setBulletFormat',
          format: { kind: 'bullet', char: 'v', font: 'Wingdings' },
        }).ok
      ).toBe(true);
      sync.flush();
      converged(a, b);
      for (const p of bullets(a)) {
        expect(p.kids).toEqual(p.scheme ? ['buFont', 'buAutoNum'] : ['buFont', 'buChar']);
        if (p.scheme) expect([p.font, p.scheme]).toEqual(['+mj-lt', 'romanUcPeriod']);
        else expect([p.font, p.char]).toEqual(['Wingdings', 'v']);
      }
    }
  });

  test("two peers pick bullets at once, one with the text font: one peer's whole bullet", async () => {
    for (let trial = 0; trial < 8; trial++) {
      const { a, b, sync } = await replicas();
      selectText(a);
      selectText(b);
      sync.pause();
      expect(runToolbarCommand(a, 'list.bulletStyle', 'hollowRound').ok).toBe(true);
      expect(b.exec({ type: 'setBulletFormat', format: { kind: 'bullet', char: '»' } }).ok).toBe(
        true
      );
      sync.flush();
      converged(a, b);
      for (const p of pPrs(a)) {
        const font = elements(p).find((c) => c.local === 'buFont' || c.local === 'buFontTx');
        const pair = [font?.local, attr(font, 'typeface'), attr(find(p, 'a:buChar'), 'char')];
        expect([
          JSON.stringify(['buFont', 'Courier New', 'o']),
          JSON.stringify(['buFontTx', null, '»']),
        ]).toContain(JSON.stringify(pair));
      }
    }
  });

  test("two peers pick theme colors with luminance at once: one peer's whole color", async () => {
    for (let trial = 0; trial < 8; trial++) {
      const { a, b, sync } = await replicas();
      selectText(a);
      expect(
        a.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent1', shift: 40 } } }).ok
      ).toBe(true);
      sync.flush();
      selectText(b);
      sync.pause();
      expect(
        a.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent2', shift: 40 } } }).ok
      ).toBe(true);
      expect(
        b.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent1', shift: -25 } } }).ok
      ).toBe(true);
      sync.flush();
      converged(a, b);
      for (const p of pPrs(a)) {
        const clr = find(find(p, 'a:buClr')!, 'a:schemeClr')!;
        const pick = [
          attr(clr, 'val'),
          ...elements(clr).map((c) => `${c.local}=${attr(c, 'val')}`),
        ];
        const aPick = JSON.stringify(['accent2', 'lumMod=60000', 'lumOff=40000']);
        const bPick = JSON.stringify(['accent1', 'lumMod=75000']);
        expect([aPick, bPick]).toContain(JSON.stringify(pick));
      }
    }
  });

  test('two peers change a theme bullet color to an RGB one and to another theme color: one child', async () => {
    const { a, b, sync } = await replicas();
    selectText(a);
    expect(a.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent1' } } }).ok).toBe(
      true
    );
    sync.flush();
    selectText(b);
    sync.pause();
    expect(a.exec({ type: 'setBulletFormat', format: { color: '#FF0000' } }).ok).toBe(true);
    expect(b.exec({ type: 'setBulletFormat', format: { color: { theme: 'accent2' } } }).ok).toBe(
      true
    );
    sync.flush();
    converged(a, b);
    for (const p of pPrs(a)) expect(elements(find(p, 'a:buClr')!).length).toBe(1);
  });
});
