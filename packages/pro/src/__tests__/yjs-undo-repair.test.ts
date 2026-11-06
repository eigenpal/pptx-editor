/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// An undo must not take what a peer still needs: an attribute both peers wrote at once,
// and a relationship and media part that a peer's copy of a picture still uses.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('4472C4') }),
      },
    ],
  });

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

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

async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

async function pair(first: number, second: number, bytes = deck()) {
  const da = new Y.Doc();
  const db = new Y.Doc();
  da.clientID = first;
  db.clientID = second;
  const l = link(da, db);
  const a = await replica(da, 'a', bytes);
  const b = await replica(db, 'b', bytes);
  return { a, b, l };
}

const typesXml = (pkg: PptxPackage) => serializeXmlPart(xmlPartOf(pkg, '/[Content_Types].xml')!);
/** The content type of a part: its `Override`, else the `Default` of its extension. */
function contentTypeFor(pkg: PptxPackage, part: string): string | undefined {
  const types = typesXml(pkg);
  for (const m of types.matchAll(/<Override\b[^>]*>/g)) {
    if (/PartName="([^"]+)"/.exec(m[0])?.[1]?.toLowerCase() === part.toLowerCase())
      return /ContentType="([^"]+)"/.exec(m[0])?.[1];
  }
  const ext = part.slice(part.lastIndexOf('.') + 1).toLowerCase();
  for (const m of types.matchAll(/<Default\b[^>]*>/g)) {
    if (/Extension="([^"]+)"/.exec(m[0])?.[1]?.toLowerCase() === ext)
      return /ContentType="([^"]+)"/.exec(m[0])?.[1];
  }
  return undefined;
}
/** Every picture a slide part embeds has its part and the png content type. */
function expectEmbedsTyped(pkg: PptxPackage, slide: string): void {
  const xmlText = serializeXmlPart(xmlPartOf(pkg, slide)!);
  const embeds = [...xmlText.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]!);
  expect(embeds.length).toBeGreaterThan(0);
  for (const id of embeds) {
    const target = relationshipsOf(pkg, slide).find((r) => r.id === id)?.target;
    expect(target).toBeDefined();
    expect(pkg.parts.get(target!)?.kind).toBe('binary');
    expect(contentTypeFor(pkg, target!)).toBe('image/png');
  }
}
/** The ids of a part's relationships, as the saved `.rels` lists them. */
const relIds = (e: PptxEditorInstance) => {
  const rels = xmlPartOf(e.session!.store.package, '/ppt/slides/_rels/slide1.xml.rels')!;
  return [...serializeXmlPart(rels).matchAll(/ Id="([^"]+)"/g)].map((m) => m[1]);
};
/** The shape's offset, `a:off` inside `p:sp`. */
const offset = (e: PptxEditorInstance) => /<p:sp>[\s\S]*?(<a:off\b[^>]*\/>)/.exec(xml(e))![1]!;

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const slidePart = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, slidePart(e))!);
const select = (e: PptxEditorInstance, index: number) =>
  e.setSelection({
    kind: 'shapes',
    nodeIds: [e.session!.layout(0)!.items[index]!.element.nodeId],
  });

describe('undo after a concurrent attribute write', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`the winner's undo keeps every required attribute (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      select(a, 0);
      select(b, 0);
      l.pause();
      expect(a.exec({ type: 'nudge', dx: IN / 4, dy: IN / 4 }).ok).toBe(true);
      expect(b.exec({ type: 'nudge', dx: -IN / 4, dy: IN / 2 }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      const winner = xml(a).includes(`x="${IN + IN / 4}"`) ? a : b;
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      // CT_Point2D needs both coordinates.
      const off = offset(a);
      expect(off).toMatch(/ x="-?\d+"/);
      expect(off).toMatch(/ y="-?\d+"/);
      expect(readPptx(a.save()!).ok).toBe(true);
      // Redo gives the winner's move back, over the value the repair wrote.
      const moved = winner === a ? `x="${IN + IN / 4}"` : `x="${IN - IN / 4}"`;
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(offset(a)).toContain(moved);
    });
  }
});

describe('undo of a picture a peer copied', () => {
  test("the author's undo keeps the relationship and the image the copy uses", async () => {
    const { a, b } = await pair(1, 2);
    expect(
      a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
    ).toBe(true);
    select(b, 1);
    expect(b.exec({ type: 'duplicate' }).ok).toBe(true);
    // A undoes its insert: its own picture goes, B's copy stays.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const pkg = a.session!.store.package;
    const embed = /r:embed="([^"]+)"/.exec(xml(a))?.[1];
    expect(embed).toBeDefined();
    const rel = relationshipsOf(pkg, slidePart(a)).find((r) => r.id === embed);
    expect(rel?.target).toBeDefined();
    expect(pkg.parts.get(rel!.target!)?.kind).toBe('binary');
    // The image's content type: a session's first picture of a type gets its own Override
    // (`ensureContentType` in core), a deck that has one keeps its Default.
    expect(contentTypeFor(pkg, rel!.target!)).toBe('image/png');
    const reopened = readPptx(a.save()!);
    expect(reopened.ok).toBe(true);
    // Redo of the insert: no relationship id twice.
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const ids = relIds(a);
    expect(new Set(ids).size).toBe(ids.length);
    expect(readPptx(a.save()!).ok).toBe(true);
  });
});

describe('undo of a single author', () => {
  test('typing into a shrink-on-overflow box, then undo, gives the body properties back', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN / 2,
            text: para(run('Hello world')),
            bodyPr: '<a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr>',
          }),
        },
      ],
    });
    const { a, b } = await pair(1, 2, bytes);
    const body = () => /<p:txBody>[\s\S]*?<\/p:txBody>/.exec(xml(a))![0];
    const before = body();
    const item = a.session!.layout(0)!.items[0]!;
    const end = {
      paragraph: 0,
      offset: item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0,
    };
    a.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: end, focus: end });
    for (let i = 0; i < 60; i += 1) expect(a.input('text', ' more words here').ok).toBe(true);
    expect(body()).not.toBe(before);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    // The same body: the box's own `a:normAutofit`, with no scale left from the typing.
    expect(body()).toContain('<a:normAutofit/>');
    expect(body()).toContain('<a:t>Hello world</a:t>');
    expect(body()).not.toContain('more words');
    expect(before).toContain('<a:normAutofit/>');
  });
});

describe('undo and redo of a repaired step', () => {
  test('undo after a repaired redo reverts the step again', async () => {
    const { a, b, l } = await pair(2, 1);
    select(a, 0);
    select(b, 0);
    const original = offset(a);
    l.pause();
    expect(a.exec({ type: 'nudge', dx: IN / 4, dy: IN / 4 }).ok).toBe(true);
    expect(b.exec({ type: 'nudge', dx: -IN / 4, dy: IN / 2 }).ok).toBe(true);
    l.flush();
    const winner = xml(a).includes(`x="${IN + IN / 4}"`) ? a : b;
    const moved = offset(winner);
    expect(winner.exec({ type: 'undo' }).ok).toBe(true);
    expect(winner.exec({ type: 'redo' }).ok).toBe(true);
    expect(offset(a)).toBe(moved);
    expect(winner.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(offset(a)).toBe(original);
  });
});

describe('a redo that Yjs applies to a later step', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`one undo after it reverts one step (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      select(a, 0);
      select(b, 0);
      l.pause();
      expect(a.exec({ type: 'nudge', dx: IN / 4, dy: IN / 4 }).ok).toBe(true);
      expect(b.exec({ type: 'nudge', dx: -IN / 4, dy: IN / 2 }).ok).toBe(true);
      l.flush();
      const winner = xml(a).includes(`x="${IN + IN / 4}"`) ? a : b;
      expect(winner.exec({ type: 'setLineWidth', points: 4 }).ok).toBe(true);
      const moved = offset(winner);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      // Yjs cannot redo the move (a repair wrote its offset); one redo still gives back one
      // step, the move, by its repair, and the next redo gives back the width.
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(offset(a)).toBe(moved);
      expect(xml(a)).not.toContain('w="50800"');
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(xml(a)).toContain('w="50800"');
      // One undo now reverts only the width.
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).not.toContain('w="50800"');
      expect(offset(a)).toBe(moved);
    });
  }
});

describe('redo of steps that Yjs cannot redo', () => {
  /** Both peers move the first shape at once; returns the replica whose move shows. */
  function conflict(
    a: PptxEditorInstance,
    b: PptxEditorInstance,
    l: { pause(): void; flush(): void }
  ) {
    l.pause();
    a.exec({ type: 'nudge', dx: IN / 4, dy: IN / 4 });
    b.exec({ type: 'nudge', dx: -IN / 4, dy: IN / 2 });
    l.flush();
  }

  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`no step applies: the repair is its own step (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      select(a, 0);
      select(b, 0);
      expect(a.exec({ type: 'setLineWidth', points: 4 }).ok).toBe(true);
      conflict(a, b, l);
      const winner = xml(a).includes(`x="${IN + IN / 4}"`) ? a : b;
      if (winner === b) expect(b.exec({ type: 'setLineWidth', points: 4 }).ok).toBe(true);
      const moved = offset(winner);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(offset(a)).toBe(moved);
      // The next undo reverts the move only; the width stays.
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(offset(a)).not.toBe(moved);
      expect(xml(a)).toContain('w="50800"');
    });

    test(`a skipped redo comes back in order after a second conflict (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      select(a, 0);
      select(b, 0);
      const start = offset(a);
      conflict(a, b, l);
      const winner = xml(a).includes(`x="${IN + IN / 4}"`) ? a : b;
      const once = offset(a);
      // Consecutive nudges are one step, also across a peer's edit: select again, so the
      // second move is a step of its own.
      for (const e of [a, b]) (e.setSelection({ kind: 'none' }), select(e, 0));
      conflict(a, b, l);
      const twice = offset(a);
      expect(winner.exec({ type: 'setLineWidth', points: 4 }).ok).toBe(true);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(offset(a)).toBe(once);
      // Yjs cannot undo the first move (the second conflict took its values), and says so.
      expect(winner.exec({ type: 'undo' }).ok).toBe(false);
      expect(offset(a)).toBe(once);
      // Each redo gives back one step: the second move (its repair), then the width.
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(offset(a)).toBe(twice);
      expect(xml(a)).not.toContain('w="50800"');
      expect(winner.exec({ type: 'redo' }).ok).toBe(true);
      expect(xml(a)).toContain('w="50800"');
      // Undo goes back one step at a time: the width, then the second move.
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(xml(a)).not.toContain('w="50800"');
      expect(offset(a)).toBe(twice);
      expect(winner.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(offset(a)).toBe(once);
      expect(start).not.toBe(once);
    });
  }
});

describe('two pictures inserted at once', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`one author's undo keeps the other picture's image (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      l.pause();
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      expect(
        b.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      l.flush();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      const pkg = a.session!.store.package;
      for (const id of [...xml(a).matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1])) {
        const rel = relationshipsOf(pkg, slidePart(a)).find((r) => r.id === id);
        expect(rel?.target && pkg.parts.get(rel.target)?.kind).toBe('binary');
      }
    });
  }
});

describe('content types after an undo and a redo', () => {
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`a picture a peer copied keeps one Default per extension (client ids ${first}, ${second})`, async () => {
      const { a, b } = await pair(first, second);
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      select(b, 1);
      expect(b.exec({ type: 'duplicate' }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      for (const e of [a, b]) {
        const saved = readPptx(e.save()!);
        expect(saved.ok).toBe(true);
        if (!saved.ok) continue;
        // Each picture keeps its part and the png content type.
        expectEmbedsTyped(saved.pkg, slidePart(e));
        const types = typesXml(saved.pkg);
        // At most one png Default: a session's first picture of a type gets its own Override
        // (`ensureContentType` in core), and no entry shows twice.
        expect((types.match(/<Default [^>]*Extension="png"/gi) ?? []).length).toBeLessThanOrEqual(
          1
        );
        const parts = [...types.matchAll(/<Override [^>]*PartName="([^"]+)"/g)].map((m) =>
          m[1]!.toLowerCase()
        );
        expect(new Set(parts).size).toBe(parts.length);
      }
    });
  }
});

describe('undo repair over the whole package', () => {
  const twoSlides = (): Uint8Array =>
    buildPptx({
      slides: [
        { layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN }) },
        { layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN }) },
      ],
    });
  /** Every internal relationship target is a part, and every part has a content type. */
  const whole = (pkg: PptxPackage) => {
    const types = typesXml(pkg);
    const defaults = new Set(
      [...types.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)].map((m) => m[1]!.toLowerCase())
    );
    const overrides = new Set(
      [...types.matchAll(/<Override\b[^>]*PartName="([^"]+)"/g)].map((m) => m[1]!.toLowerCase())
    );
    for (const name of pkg.parts.keys()) {
      if (name === '/[Content_Types].xml') continue;
      const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
      expect(overrides.has(name.toLowerCase()) || defaults.has(ext)).toBe(true);
      if (pkg.parts.get(name)?.kind !== 'xml' || name.endsWith('.rels')) continue;
      for (const rel of relationshipsOf(pkg, name))
        if (!rel.external && rel.target) expect(pkg.parts.has(rel.target)).toBe(true);
    }
  };
  for (const [first, second] of [
    [1, 2],
    [2, 1],
  ] as const) {
    test(`a picture a peer pasted on another slide keeps its image (client ids ${first}, ${second})`, async () => {
      const { a, b } = await pair(first, second, twoSlides());
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      select(b, 1);
      const payload = b.copy()!;
      b.goToSlide(1);
      expect(b.exec({ type: 'paste', payload }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      whole(a.session!.store.package);
      const reopened = readPptx(a.save()!);
      expect(reopened.ok).toBe(true);
      if (reopened.ok) whole(reopened.pkg);
    });

    test(`a peer's picture keeps the content type that the author's first picture added (client ids ${first}, ${second})`, async () => {
      const { a, b } = await pair(first, second);
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      const other = new Uint8Array([...PNG, 0]);
      expect(
        b.exec({ type: 'insertImage', bytes: other, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      whole(a.session!.store.package);
      expect(readPptx(a.save()!).ok).toBe(true);
    });

    test(`an undo that Yjs can only take from an older step changes nothing (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      select(a, 0);
      select(b, 0);
      // Both peers set a fill first, so each has an older step.
      expect(a.exec({ type: 'setLineWidth', points: 4 }).ok).toBe(true);
      expect(b.exec({ type: 'setLineWidth', points: 2 }).ok).toBe(true);
      l.pause();
      expect(a.exec({ type: 'nudge', dx: IN / 4, dy: IN / 4 }).ok).toBe(true);
      expect(b.exec({ type: 'nudge', dx: -IN / 4, dy: IN / 2 }).ok).toBe(true);
      l.flush();
      const loser = xml(a).includes(`x="${IN + IN / 4}"`) ? b : a;
      const width = /<a:ln w="(\d+)"/.exec(xml(a))?.[1];
      const where = offset(a);
      const done = loser.exec({ type: 'undo' });
      expect(fingerprint(a)).toBe(fingerprint(b));
      // The older step stays: the width and the position do not change.
      expect(/<a:ln w="(\d+)"/.exec(xml(a))?.[1]).toBe(width);
      expect(offset(a)).toBe(where);
      expect(done.ok).toBe(false);
      // The reason says that a peer's change took the step, not that there is nothing to undo.
      expect(done).toMatchObject({ reason: 'disabled.historyConflict' });
    });

    test(`an undo while a peer copies the picture at the same time keeps the copy's image (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      select(b, 1);
      l.pause();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(b.exec({ type: 'duplicate' }).ok).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      // Each picture's relationship and media part stay.
      const embeds = [...xml(a).matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]!);
      expect(embeds.length).toBeGreaterThan(0);
      const rels = relationshipsOf(a.session!.store.package, slidePart(a));
      for (const id of embeds) expect(rels.some((r) => r.id === id && !!r.target)).toBe(true);
      whole(a.session!.store.package);
      const reopened = readPptx(a.save()!);
      expect(reopened.ok).toBe(true);
      if (reopened.ok) whole(reopened.pkg);
    });

    test(`an undo while a peer inserts another picture keeps that picture's content type (client ids ${first}, ${second})`, async () => {
      const { a, b, l } = await pair(first, second);
      expect(
        a.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      l.pause();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      const other = new Uint8Array([...PNG, 0]);
      expect(
        b.exec({ type: 'insertImage', bytes: other, mime: 'image/png', width: 1, height: 1 }).ok
      ).toBe(true);
      l.flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      whole(a.session!.store.package);
      expect(readPptx(a.save()!).ok).toBe(true);
    });
  }
});
