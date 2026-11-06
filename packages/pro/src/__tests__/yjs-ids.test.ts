/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Ids that two peers mint at the same time: two editors on two linked
// Y.Docs, the link paused, the same kind of insert on both, then a flush and a save. Each new
// relationship, part, slide, and shape keeps its own id, so the saved file has no duplicate and
// no peer's slide, image, or link is lost.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as Y from 'yjs';
import {
  createPptxEditor,
  type EditorCommand,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  elements,
  readPptx,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../scripts/fixtures/slide-structure.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import {
  ACTOR_ID_STRIPE,
  actorStripe,
} from '../../../core/src/automation/commands/actor-scoped-ids.ts';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        ext: '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1000"/></p:ext></p:extLst>',
        shapes:
          shape({ id: 2, name: 'A', x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('one')) }) +
          shape({ id: 3, name: 'B', x: 4 * IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('two')) }),
      },
    ],
  });

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78,
  0x9c, 0x63, 0, 1, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);
/** A second picture: the same PNG with one pixel byte changed. */
const PNG2 = PNG.map((b, i) => (i === 44 ? 0x64 : b));

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

async function replicas(
  bytes: (actors: readonly string[]) => Uint8Array = deck,
  clientIds: readonly [number, number] = [1_001, 2_002]
) {
  // Fixed client ids keep the peers apart. The session id also holds a module-wide counter, so
  // a stripe depends on test order; tests read it from `sessionId`, and two stripes match once
  // in 65,536 runs.
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  docs[0].clientID = clientIds[0];
  docs[1].clientID = clientIds[1];
  const sync = link(docs[0], docs[1]);
  const sessions = docs.map((doc, i) =>
    createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: `p${i}`, name: `p${i}` },
    })
  );
  const file = bytes(sessions.map((s) => s.sessionId));
  const editors: PptxEditorInstance[] = [];
  for (const session of sessions) {
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(file)).ok).toBe(true);
    editors.push(editor);
  }
  return { a: editors[0]!, b: editors[1]!, sync, sessions, docs };
}

/** Both replicas match, and the saved file reopens with the same content. */
function saved(a: PptxEditorInstance, b: PptxEditorInstance): PptxPackage {
  const pkg = a.session!.store.package;
  expect(canonicalPackageFingerprint(pkg)).toBe(
    canonicalPackageFingerprint(b.session!.store.package)
  );
  const r = readPptx(a.save()!);
  if (!r.ok) throw new Error('reopen failed');
  return r.pkg;
}

const relsOf = (pkg: PptxPackage, part: string): Element[] => {
  const slash = part.lastIndexOf('/');
  const rels = xmlPartOf(pkg, `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`);
  return rels ? elements(rels.root) : [];
};
const unique = (values: readonly (string | undefined)[]): boolean =>
  new Set(values).size === values.length;
const byName = (e: PptxEditorInstance, name: string): string => {
  const item = e
    .session!.layout(0)!
    .items.find((i) => i.kind === 'shape' && i.element.name === name);
  return item!.element.nodeId;
};
const all = (root: Element, local: string): Element[] => {
  const out: Element[] = [];
  const walk = (el: Element): void => {
    if (el.local === local) out.push(el);
    for (const c of elements(el)) walk(c);
  };
  walk(root);
  return out;
};
const slidePart = (pkg: PptxPackage, i = 0) => {
  const pres = xmlPartOf(pkg, '/ppt/presentation.xml')!;
  const ids = all(pres.root, 'sldId');
  const rid = ids[i]!.attrs.find((x) => x.local === 'id' && x.ns !== '')!.value;
  const rel = relsOf(pkg, '/ppt/presentation.xml').find((r) => attr(r, 'Id') === rid)!;
  return `/ppt/${attr(rel, 'Target')}`;
};

describe('ids that two peers mint at once', () => {
  test('two peers add a slide: both slides, distinct slide ids, relationships, and parts', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'addSlide' }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const pres = xmlPartOf(pkg, '/ppt/presentation.xml')!;
    const ids = all(pres.root, 'sldId');
    expect(ids).toHaveLength(3);
    expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
    const rels = relsOf(pkg, '/ppt/presentation.xml');
    expect(unique(rels.map((r) => attr(r, 'Id')))).toBe(true);
    const parts = [0, 1, 2].map((i) => slidePart(pkg, i));
    expect(unique(parts)).toBe(true);
    for (const p of parts) expect(xmlPartOf(pkg, p)).toBeDefined();
  });

  test('two peers insert different pictures: two media parts, each picture its own', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    const pic = { type: 'insertImage', mime: 'image/png', width: 10, height: 10 } as const;
    expect(a.exec({ ...pic, bytes: PNG }).ok).toBe(true);
    expect(b.exec({ ...pic, bytes: PNG2 }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const slide = slidePart(pkg);
    const rels = relsOf(pkg, slide);
    expect(unique(rels.map((r) => attr(r, 'Id')))).toBe(true);
    const embeds = all(xmlPartOf(pkg, slide)!.root, 'blip').map(
      (blip) => blip.attrs.find((x) => x.local === 'embed')!.value
    );
    expect(embeds).toHaveLength(2);
    expect(unique(embeds)).toBe(true);
    const targets = embeds.map((id) =>
      attr(
        rels.find((r) => attr(r, 'Id') === id)!,
        'Target'
      )
    );
    expect(unique(targets)).toBe(true);
    const media = targets.map((t) => pkg.parts.get(`/ppt/slides/${t}`.replace('/slides/../', '/')));
    const bytes = media.map((m) => (m?.kind === 'binary' ? m.bytes[44] : undefined));
    expect(new Set(bytes)).toEqual(new Set([PNG[44], PNG2[44]]));
    // Each picture's content type is its own override, never a second `Default` for png.
    const ct = elements(xmlPartOf(pkg, '/[Content_Types].xml')!.root);
    const pngDefaults = ct.filter(
      (e) => e.local === 'Default' && (attr(e, 'Extension') ?? '').toLowerCase() === 'png'
    );
    expect(pngDefaults.length).toBeLessThanOrEqual(1);
    // Every picture's media part has a content type: its own `Override`, or the png `Default`.
    for (const t of targets) {
      const part = `/ppt/slides/${t}`.replace('/slides/../', '/').toLowerCase();
      const typed = ct.some(
        (e) =>
          (e.local === 'Override' && (attr(e, 'PartName') ?? '').toLowerCase() === part) ||
          (e.local === 'Default' && (attr(e, 'Extension') ?? '').toLowerCase() === 'png')
      );
      expect(typed).toBe(true);
    }
  });

  test('two peers link two shapes: each link keeps its own address', async () => {
    const { a, b, sync } = await replicas();
    const select = (e: PptxEditorInstance, name: string) =>
      e.setSelection({ kind: 'shapes', nodeIds: [byName(e, name)] });
    select(a, 'A');
    select(b, 'B');
    sync.pause();
    expect(a.exec({ type: 'setLink', url: 'https://a.example/' }).ok).toBe(true);
    expect(b.exec({ type: 'setLink', url: 'https://b.example/' }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const slide = slidePart(pkg);
    const rels = relsOf(pkg, slide);
    expect(unique(rels.map((r) => attr(r, 'Id')))).toBe(true);
    const links = all(xmlPartOf(pkg, slide)!.root, 'hlinkClick').map((h) => {
      const id = h.attrs.find((x) => x.local === 'id')!.value;
      return attr(
        rels.find((r) => attr(r, 'Id') === id)!,
        'Target'
      );
    });
    expect(links.sort()).toEqual(['https://a.example/', 'https://b.example/']);
  });

  test('two peers copy one slide: each copy keeps its own creation id', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    expect(a.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(b.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const ids = [0, 1, 2].map((i) =>
      attr(all(xmlPartOf(pkg, slidePart(pkg, i))!.root, 'creationId')[0]!, 'val')
    );
    expect(unique(ids)).toBe(true);
  });

  test("two peers copy a slide with comments: each author's comment indexes stay unique", async () => {
    const { a, b, sync } = await replicas(buildSlideStructureDeck);
    sync.pause();
    expect(a.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(b.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const idx: string[] = [];
    // Comment parts are kept as bytes; read each `p:cm` start tag.
    for (const [name, part] of pkg.parts)
      if (name.startsWith('/ppt/comments/') && part.kind === 'binary')
        for (const m of new TextDecoder().decode(part.bytes).matchAll(/<p:cm\s[^>]*>/g)) {
          const tag = m[0];
          idx.push(`${/authorId="(\d+)"/.exec(tag)?.[1]}:${/\bidx="(\d+)"/.exec(tag)?.[1]}`);
        }
    expect(idx.length).toBe(6);
    expect(unique(idx)).toBe(true);
  });

  test("a slide a peer deleted: its id is not minted again, so the peer's undo adds no duplicate", async () => {
    const { a, b } = await replicas();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    const pkg = saved(a, b);
    const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
    expect(ids).toHaveLength(3);
    expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
    expect(unique(relsOf(pkg, '/ppt/presentation.xml').map((r) => attr(r, 'Id')))).toBe(true);
  });

  test('a reload in the same editor keeps the record of minted numbers', async () => {
    const { a, b } = await replicas();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    // The same editor opens the deck again and adopts the room.
    expect((await a.load(deck())).ok).toBe(true);
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    const pkg = saved(a, b);
    const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
    expect(ids).toHaveLength(3);
    expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
    expect(unique([0, 1, 2].map((i) => slidePart(pkg, i)))).toBe(true);
  });

  test('a new editor on the same session keeps the record of minted numbers (a remount)', async () => {
    const { a, b, sessions } = await replicas();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
    // The host destroys the editor and mounts a new one on the same collaboration session.
    a.destroy();
    const a2 = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session: sessions[0]! })],
    });
    expect((await a2.load(deck())).ok).toBe(true);
    expect(a2.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    const pkg = saved(a2, b);
    const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
    expect(ids).toHaveLength(3);
    expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
    expect(unique(relsOf(pkg, '/ppt/presentation.xml').map((r) => attr(r, 'Id')))).toBe(true);
    expect(unique([0, 1, 2].map((i) => slidePart(pkg, i)))).toBe(true);
  });

  // A reload or a remount keeps the mirror's node keys apart from the ones the room holds,
  // so a peer's undo and a late joiner see the same deck.
  for (const who of [0, 1] as const)
    for (const mode of ['reload', 'remount'] as const)
      test(`a ${mode} by peer ${who}: the other peer undo brings the slide back, and a third replica matches`, async () => {
        const r = await replicas();
        const eds = [r.a, r.b];
        expect(eds[who]!.exec({ type: 'addSlide' }).ok).toBe(true);
        expect(eds[1 - who]!.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
        let ed = eds[who]!;
        if (mode === 'remount') {
          ed.destroy();
          ed = createPptxEditor({
            fixedMeasurer: true,
            substitutes: false,
            modules: [collaborationModule({ session: r.sessions[who]! })],
          });
        }
        expect((await ed.load(deck())).ok).toBe(true);
        expect(ed.exec({ type: 'addSlide' }).ok).toBe(true);
        expect(eds[1 - who]!.exec({ type: 'undo' }).ok).toBe(true);
        const pkg = saved(ed, eds[1 - who]!);
        const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
        expect(ids).toHaveLength(3);
        expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
        const d3 = new Y.Doc();
        d3.clientID = 3_003;
        Y.applyUpdate(d3, Y.encodeStateAsUpdate(r.docs[who]));
        const s3 = createYjsCollaborationSession({
          doc: d3,
          documentId: 'deck',
          identity: { actorId: 'p3', name: 'p3' },
        });
        const e3 = createPptxEditor({
          fixedMeasurer: true,
          substitutes: false,
          modules: [collaborationModule({ session: s3 })],
        });
        expect((await e3.load(deck())).ok).toBe(true);
        expect(canonicalPackageFingerprint(e3.session!.store.package)).toBe(
          canonicalPackageFingerprint(ed.session!.store.package)
        );
      });

  test('two peers insert a shape: distinct shape ids on the slide', async () => {
    const { a, b, sync } = await replicas();
    sync.pause();
    expect(a.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(b.exec({ type: 'insertShape', preset: 'ellipse' }).ok).toBe(true);
    sync.flush();
    const pkg = saved(a, b);
    const ids = all(xmlPartOf(pkg, slidePart(pkg))!.root, 'cNvPr').map((c) => attr(c, 'id'));
    expect(ids).toHaveLength(5);
    expect(unique(ids)).toBe(true);
  });

  // A number of the opened file that falls in a replica's stripe: the replica never mints it,
  // so a peer that deletes its object and then undoes adds no duplicate.
  for (const clientIds of [
    [1_001, 2_002],
    [2_002, 1_001],
  ] as const) {
    test(`a file shape id in the stripe, deleted by a peer: not minted again (clients ${clientIds})`, async () => {
      const fileDeck = (actors: readonly string[]): Uint8Array => {
        const own = actorStripe(actors[0]!);
        expect(own).toBeGreaterThan(3);
        return buildPptx({
          slides: [
            {
              layout: 'blank',
              shapes:
                shape({ id: 2, name: 'A', x: IN, y: IN, cx: IN, cy: IN, text: para(run('a')) }) +
                shape({
                  id: own,
                  name: 'X',
                  x: 3 * IN,
                  y: IN,
                  cx: IN,
                  cy: IN,
                  text: para(run('x')),
                }),
            },
          ],
        });
      };
      const { a, b } = await replicas(fileDeck, clientIds);
      b.setSelection({ kind: 'shapes', nodeIds: [byName(b, 'X')] });
      expect(b.exec({ type: 'delete' }).ok).toBe(true);
      expect(a.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      const pkg = saved(a, b);
      const ids = all(xmlPartOf(pkg, slidePart(pkg))!.root, 'cNvPr').map((c) => attr(c, 'id'));
      expect(ids).toHaveLength(4);
      expect(unique(ids)).toBe(true);
    });

    test(`a file slide id and relationship id in the stripe, deleted by a peer: not minted again (clients ${clientIds})`, async () => {
      const fileDeck = (actors: readonly string[]): Uint8Array => {
        const own = actorStripe(actors[0]!);
        // The first slide id and relationship number of the stripe, the ones a mint takes first.
        const slideId =
          256 + ((((own - 256) % ACTOR_ID_STRIPE) + ACTOR_ID_STRIPE) % ACTOR_ID_STRIPE);
        const files = unzipSync(
          buildPptx({
            slides: [
              { layout: 'blank', shapes: '' },
              { layout: 'blank', shapes: '' },
            ],
          })
        );
        const edit = (name: string, change: (xml: string) => string): void => {
          files[name] = strToU8(change(strFromU8(files[name]!)));
        };
        edit('ppt/presentation.xml', (x) =>
          x.replace(
            '<p:sldId id="257" r:id="rId102"/>',
            `<p:sldId id="${slideId}" r:id="rId${own}"/>`
          )
        );
        edit('ppt/_rels/presentation.xml.rels', (x) => x.replace('Id="rId102"', `Id="rId${own}"`));
        return zipSync(files);
      };
      const { a, b } = await replicas(fileDeck, clientIds);
      expect(b.exec({ type: 'deleteSlide', index: 1 }).ok).toBe(true);
      expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      const pkg = saved(a, b);
      const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
      expect(ids).toHaveLength(3);
      expect(unique(ids.map((s) => attr(s, 'id')))).toBe(true);
      expect(unique(relsOf(pkg, '/ppt/presentation.xml').map((r) => attr(r, 'Id')))).toBe(true);
      expect(unique([0, 1, 2].map((i) => slidePart(pkg, i)))).toBe(true);
    });
  }
});

describe('every shape mint, on both peers at once', () => {
  /** Four text shapes, so each peer can group its own pair. */
  const fourShapes = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: ['A', 'B', 'C', 'D']
            .map((name, i) =>
              shape({
                id: i + 2,
                name,
                x: (i * 2 + 1) * IN,
                y: IN,
                cx: IN,
                cy: IN,
                text: para(run(name)),
              })
            )
            .join(''),
        },
      ],
    });
  const select = (e: PptxEditorInstance, ...names: string[]): void =>
    e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => byName(e, n)) });
  const paste = (e: PptxEditorInstance): EditorCommand => {
    const payload = e.copy();
    if (!payload) throw new Error('nothing to copy');
    return { type: 'paste', payload };
  };
  // Each case: what peer A and peer B select first (their own shapes), and the command.
  const cases: readonly {
    readonly name: string;
    readonly own: readonly [readonly string[], readonly string[]];
    readonly command: (e: PptxEditorInstance) => EditorCommand;
  }[] = [
    { name: 'insert a text box', own: [[], []], command: () => ({ type: 'insertTextBox' }) },
    { name: 'insert a line', own: [[], []], command: () => ({ type: 'insertLine', kind: 'line' }) },
    {
      name: 'insert a table',
      own: [[], []],
      command: () => ({ type: 'insertTable', rows: 2, cols: 2 }),
    },
    { name: 'duplicate', own: [['A'], ['A']], command: () => ({ type: 'duplicate' }) },
    { name: 'copy and paste', own: [['A', 'B'], ['C']], command: paste },
    {
      name: 'group',
      own: [
        ['A', 'B'],
        ['C', 'D'],
      ],
      command: () => ({ type: 'group' }),
    },
  ];

  /** Both replicas match, the merged slide passes the store part checks, and ids are unique. */
  const converged = (a: PptxEditorInstance, b: PptxEditorInstance, shapes: number): void => {
    const pkg = saved(a, b);
    expect(packageViolations(a.session!.store.package)).toEqual([]);
    expect(packageViolations(pkg)).toEqual([]);
    const ids = all(xmlPartOf(pkg, slidePart(pkg))!.root, 'cNvPr').map((c) => attr(c, 'id'));
    expect(ids).toHaveLength(shapes);
    expect(unique(ids)).toBe(true);
  };

  for (const c of cases)
    test(`${c.name}: unique ids that pass the part checks, also after undo and redo`, async () => {
      const { a, b, sync } = await replicas(fourShapes);
      // The tree's own id and four shapes.
      const before = 5;
      select(a, ...c.own[0]);
      select(b, ...c.own[1]);
      sync.pause();
      expect(a.exec(c.command(a)).ok).toBe(true);
      expect(b.exec(c.command(b)).ok).toBe(true);
      sync.flush();
      const pkg = saved(a, b);
      const after = all(xmlPartOf(pkg, slidePart(pkg))!.root, 'cNvPr').length;
      expect(after).toBeGreaterThan(before);
      converged(a, b, after);
      // Each peer undoes its own mint: the slide is back to the file's shapes.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      converged(a, b, before);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      converged(a, b, after);
    });

  test('a peer that joins late mints in its own stripe while the others mint too', async () => {
    const { a, b, docs, sync } = await replicas(fourShapes);
    const late = new Y.Doc();
    late.clientID = 3_003;
    Y.applyUpdate(late, Y.encodeStateAsUpdate(docs[0]));
    const syncA = link(docs[0], late);
    const syncB = link(docs[1], late);
    const session = createYjsCollaborationSession({
      doc: late,
      documentId: 'deck',
      identity: { actorId: 'p3', name: 'p3' },
    });
    const c = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await c.load(fourShapes())).ok).toBe(true);
    for (const s of [sync, syncA, syncB]) s.pause();
    for (const e of [a, b, c]) {
      expect(e.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
      expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    }
    for (const s of [sync, syncA, syncB]) s.flush();
    converged(a, c, 8);
    converged(b, c, 8);
    const pkg = saved(a, c);
    const slideIds = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId');
    expect(slideIds).toHaveLength(4);
    expect(unique(slideIds.map((s) => attr(s, 'id')))).toBe(true);
    expect(unique([0, 1, 2, 3].map((i) => slidePart(pkg, i)))).toBe(true);
  });
});

describe('ids minted while the session attaches', () => {
  // A host listener that edits as soon as the session is ready runs inside `attach`. The
  // editor binds the stripe before the modules attach, so that edit mints striped ids too.
  for (const clientIds of [
    [1_001, 2_002],
    [2_002, 1_001],
  ] as const) {
    test(`a slide added from a status listener during attach gets a striped id (clients ${clientIds})`, async () => {
      const docs = [new Y.Doc(), new Y.Doc()] as const;
      docs[0].clientID = clientIds[0];
      docs[1].clientID = clientIds[1];
      const sync = link(docs[0], docs[1]);
      const editors: PptxEditorInstance[] = [];
      const sessions = docs.map((doc, i) =>
        createYjsCollaborationSession({
          doc,
          documentId: 'deck',
          identity: { actorId: `p${i}`, name: `p${i}` },
        })
      );
      for (const [i, session] of sessions.entries()) {
        // The first replica seeds the room; the second adopts it while the peer edits at once.
        if (i === 1) sync.pause();
        const editor = createPptxEditor({
          fixedMeasurer: true,
          substitutes: false,
          modules: [collaborationModule({ session })],
        });
        let added = false;
        session.subscribeStatus((s) => {
          if (added || s === 'initializing') return;
          added = true;
          expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
        });
        expect((await editor.load(deck())).ok).toBe(true);
        expect(added).toBe(true);
        editors.push(editor);
      }
      expect(editors[0]!.exec({ type: 'addSlide' }).ok).toBe(true);
      sync.flush();
      const pkg = saved(editors[0]!, editors[1]!);
      const ids = all(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'sldId').map((s) =>
        Number(attr(s, 'id'))
      );
      expect(ids).toHaveLength(4);
      expect(unique(ids.map(String))).toBe(true);
      // Each added slide's id lies in its own replica's stripe.
      for (const session of sessions) {
        const stripe = actorStripe(session.sessionId);
        expect(ids.some((id) => id % ACTOR_ID_STRIPE === stripe)).toBe(true);
      }
      expect(unique(relsOf(pkg, '/ppt/presentation.xml').map((r) => attr(r, 'Id')))).toBe(true);
    });
  }
});
