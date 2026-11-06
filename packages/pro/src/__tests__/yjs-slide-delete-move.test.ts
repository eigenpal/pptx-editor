/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A slide that one peer deletes while another moves it. The move inserts the slide
// entry again, but the delete wins, as a deletion wins over a concurrent move in DOCX Editor:
// every replica that sees an entry whose relationship is gone removes it. Before, the
// repair gave the slide part back under the `xml` Default, with no relationships, and the file
// named a slide that no reader counted.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { lostReferences } from '../collaboration/yjs/undo-repair.ts';

const IN = 914_400;

/** Six slides; the last one has text and speaker notes. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      ...Array.from({ length: 5 }, () => ({ shapes: '' })),
      {
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Last')) }),
        notes: 'Say goodbye.',
      },
    ],
  });

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

async function replica(doc: Y.Doc, actorId: string, bytes = deck()): Promise<PptxEditorInstance> {
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

/** Two replicas; `deleterFirst` gives the deleting peer the lower Yjs client id. */
async function pair(deleterFirst: boolean, bytes?: Uint8Array) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  docA.clientID = deleterFirst ? 1 : 2;
  docB.clientID = deleterFirst ? 2 : 1;
  const sync = link(docA, docB);
  const a = await replica(docA, 'a', bytes);
  const b = await replica(docB, 'b', bytes);
  return { sync, a, b };
}

/** Every check a saved file of a slide delete and move must pass. Returns its slide count. */
function expectValidSave(e: PptxEditorInstance): number {
  const bytes = e.save()!;
  const zip = unzipSync(bytes);
  const names = new Set(Object.keys(zip));
  const types = strFromU8(zip['[Content_Types].xml']!);
  const overrides = (types.match(/<Override [^>]*>/g) ?? []).map(
    (o) => /PartName="\/([^"]+)"/.exec(o)![1]!
  );
  // No duplicate Override, and each one names a part of the zip.
  expect(new Set(overrides).size).toBe(overrides.length);
  // Each Override names a part of the zip. A `.rels` Override is left out: a plain slide delete
  // in core keeps the fixture's Override of a deleted chart's relationships, with or without a
  // peer (a core defect of its own, not of the collaboration merge).
  for (const name of overrides)
    if (!name.endsWith('.rels')) expect(names.has(name) ? name : `missing ${name}`).toBe(name);
  // Every slide and notes slide has its own Override and relationships.
  for (const name of names) {
    if (!/^ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml$/.test(name)) continue;
    expect(overrides).toContain(name);
    expect(names.has(name.replace(/\/([^/]+)$/, '/_rels/$1.rels'))).toBe(true);
  }
  // Every internal relationship target exists.
  for (const name of names) {
    if (!name.endsWith('.rels')) continue;
    const dir = name.replace(/_rels\/[^/]+$/, '');
    for (const m of strFromU8(zip[name]!).matchAll(/<Relationship ([^>]*)>/g)) {
      if (/TargetMode="External"/.test(m[1]!)) continue;
      const target = /Target="([^"]+)"/.exec(m[1]!)![1]!;
      const parts: string[] = [];
      for (const seg of (target.startsWith('/') ? target.slice(1) : dir + target).split('/'))
        if (seg === '..') parts.pop();
        else if (seg !== '.' && seg !== '') parts.push(seg);
      const folded = new Set([...names].map((n) => n.toLowerCase()));
      expect(folded.has(parts.join('/').toLowerCase())).toBe(true);
    }
  }
  // The slide list: unique ids and r:ids, one entry for each slide.
  const entries = [...strFromU8(zip['ppt/presentation.xml']!).matchAll(/<p:sldId ([^>]*)>/g)];
  const ids = entries.map((m) => /\bid="(\d+)"/.exec(m[1]!)![1]);
  const rIds = entries.map((m) => /r:id="([^"]+)"/.exec(m[1]!)![1]);
  expect(new Set(ids).size).toBe(ids.length);
  expect(new Set(rIds).size).toBe(rIds.length);
  // The sections name each slide of the list once, in list order (a presentation app repairs
  // a file that names a slide twice or names a slide that is gone).
  const pres = strFromU8(zip['ppt/presentation.xml']!);
  const named = [...pres.matchAll(/<p14:sldId id="(\d+)"/g)].map((m) => m[1]);
  if (named.length > 0) expect(named).toEqual(ids);
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('reopen');
  const slides = presentationInfo(r.pkg).slides.length;
  expect(entries.length).toBe(slides);
  expect([...names].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length).toBe(slides);
  return slides;
}

function expectConverged(...editors: PptxEditorInstance[]): void {
  const prints = editors.map((e) => canonicalPackageFingerprint(e.session!.store.package));
  for (const p of prints) expect(p).toBe(prints[0]!);
}

const hasLast = (e: PptxEditorInstance): boolean =>
  e.snapshot().slides.length > 0 &&
  presentationInfo(e.session!.store.package).slides.some((s) => s.partName.endsWith('slide6.xml'));

/** A concurrent delete of slide 6 on `a` and a move of it to the front on `b`. */
function deleteAndMove(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(a.exec({ type: 'deleteSlides', indices: [5] }).ok).toBe(true);
  expect(b.exec({ type: 'moveSlides', indices: [5], to: 0 }).ok).toBe(true);
}

describe('a slide that one peer deletes while another moves it', () => {
  for (const deleterFirst of [true, false]) {
    const order = deleterFirst
      ? 'deleter has the lower client id'
      : 'mover has the lower client id';

    test(`stays deleted on both replicas, and the files stay valid (${order})`, async () => {
      const { sync, a, b } = await pair(deleterFirst);
      sync.pause();
      deleteAndMove(a, b);
      sync.flush();
      expectConverged(a, b);
      for (const e of [a, b]) {
        expect(e.snapshot().slides.length).toBe(5);
        expect(hasLast(e)).toBe(false);
        expect(expectValidSave(e)).toBe(5);
      }
      expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
      expectConverged(a, b);
      for (const e of [a, b]) expect(expectValidSave(e)).toBe(6);
    });

    test(`the deleter that gets only the move saves a valid file (${order})`, async () => {
      const docA = new Y.Doc();
      const docB = new Y.Doc();
      docA.clientID = deleterFirst ? 1 : 2;
      docB.clientID = deleterFirst ? 2 : 1;
      const sync = link(docA, docB);
      const a = await replica(docA, 'a');
      const b = await replica(docB, 'b');
      sync.pause();
      // Only the mover's update reaches the deleter; the mover goes offline.
      const fromB: Uint8Array[] = [];
      docB.on(
        'update',
        (u: Uint8Array, origin: unknown) => void (origin !== 'link' && fromB.push(u))
      );
      deleteAndMove(a, b);
      for (const u of fromB) Y.applyUpdate(docA, u, 'link');
      expect(a.snapshot().slides.length).toBe(5);
      expect(expectValidSave(a)).toBe(5);
    });

    test(`undo and redo on either side keep the replicas equal and the files valid (${order})`, async () => {
      for (const undoer of ['a', 'b', 'both'] as const) {
        const { sync, a, b } = await pair(deleterFirst);
        sync.pause();
        deleteAndMove(a, b);
        sync.flush();
        if (undoer === 'both') {
          sync.pause();
          a.exec({ type: 'undo' });
          b.exec({ type: 'undo' });
          sync.flush();
        } else {
          const e = undoer === 'a' ? a : b;
          e.exec({ type: 'undo' });
          expectConverged(a, b);
          expectValidSave(a);
          e.exec({ type: 'redo' });
          expectConverged(a, b);
          expectValidSave(a);
          e.exec({ type: 'undo' });
        }
        expectConverged(a, b);
        const counts = [a, b].map((x) => expectValidSave(x));
        expect(counts[0]).toBe(counts[1]);
        expect(counts[0]).toBe(a.snapshot().slides.length);
        // An undo of the delete gives the slide back; an undo of the move alone does not.
        expect(counts[0]).toBe(undoer === 'b' ? 5 : 6);
      }
    });
  }

  test('three replicas: one deletes, two move; all agree', async () => {
    const docs = [new Y.Doc(), new Y.Doc(), new Y.Doc()];
    docs.forEach((d, i) => (d.clientID = i + 1));
    // Every pair is linked: a link does not pass on what it received.
    const links = [link(docs[0]!, docs[1]!), link(docs[1]!, docs[2]!), link(docs[0]!, docs[2]!)];
    const [a, b, c] = [
      await replica(docs[0]!, 'a'),
      await replica(docs[1]!, 'b'),
      await replica(docs[2]!, 'c'),
    ];
    for (const l of links) l.pause();
    a.exec({ type: 'deleteSlides', indices: [5] });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    c.exec({ type: 'moveSlides', indices: [5], to: 2 });
    for (const l of links) l.flush();
    expectConverged(a!, b!, c!);
    for (const e of [a!, b!, c!]) expect(expectValidSave(e)).toBe(5);
  });
});

describe('the slide-structure fixture (charts, embeddings, comments, shared media)', () => {
  const fixture = readFileSync(join(import.meta.dir, '../../../../fixtures/slide-structure.pptx'));
  const undoers = { A: ['a'], B: ['b'], both: ['a', 'b'] } as const;
  for (const deleterFirst of [true, false])
    for (const index of [0, 1, 3])
      for (const [who, sides] of Object.entries(undoers)) {
        const order = deleterFirst ? 'deleter first' : 'mover first';
        test(`a delete and a move of slide ${index + 1}, then an undo by ${who}, stay valid (${order})`, async () => {
          const { sync, a, b } = await pair(deleterFirst, new Uint8Array(fixture));
          const count = a.snapshot().slides.length;
          sync.pause();
          expect(a.exec({ type: 'deleteSlides', indices: [index] }).ok).toBe(true);
          expect(
            b.exec({ type: 'moveSlides', indices: [index], to: index === 0 ? count - 1 : 0 }).ok
          ).toBe(true);
          sync.flush();
          expectConverged(a, b);
          for (const e of [a, b]) expect(expectValidSave(e)).toBe(count - 1);
          if (sides.length === 2) sync.pause();
          for (const side of sides) (side === 'a' ? a : b).exec({ type: 'undo' });
          sync.flush();
          expectConverged(a, b);
          // The deleter's undo brings the slide back; the mover's undo alone does not.
          const back = (sides as readonly string[]).includes('a');
          for (const e of [a, b]) expect(expectValidSave(e)).toBe(back ? count : count - 1);
        });
      }
});

describe('slide lists that arrive broken, and updates that are still pending', () => {
  const fixture = readFileSync(join(import.meta.dir, '../../../../fixtures/slide-structure.pptx'));
  const S = (name: string, ids: number[]) =>
    `<p14:section name="${name}" id="{${name}}"><p14:sldIdLst>${ids
      .map((i) => `<p14:sldId id="${i}"/>`)
      .join('')}</p14:sldIdLst></p14:section>`;
  /** The fixture with other sections, and another slide list when `list` is given. */
  const variant = (sections: string, list?: string, dropRels = false): Uint8Array => {
    const z = { ...unzipSync(new Uint8Array(fixture)) };
    let x = strFromU8(z['ppt/presentation.xml']!);
    x = x.replace(
      /(<p14:sectionLst[^>]*>)[^]*?(<\/p14:sectionLst>)/,
      (_m, open: string, close: string) => open + sections + close
    );
    if (list) x = x.replace(/<p:sldIdLst>[^]*?<\/p:sldIdLst>/, list);
    z['ppt/presentation.xml'] = strToU8(x);
    if (dropRels) delete z['ppt/_rels/presentation.xml.rels'];
    return zipSync(z);
  };
  const pkgOf = (bytes: Uint8Array) => {
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('read');
    return r.pkg;
  };
  const presentationOut = (before: Uint8Array, after: Uint8Array) =>
    lostReferences(pkgOf(before), pkgOf(after)).find((p) => p.partName === '/ppt/presentation.xml');

  test('a file that arrives with a slide in two sections is made valid on the next change', async () => {
    const bytes = variant(S('Intro', [256, 257, 259]) + S('Body', [258, 259]));
    const { a, b } = await pair(true, bytes);
    expect(b.exec({ type: 'moveSlides', indices: [1], to: 2 }).ok).toBe(true);
    expectConverged(a, b);
    for (const e of [a, b]) expect(expectValidSave(e)).toBe(4);
  });

  test('a dead entry goes, and a live entry with the same slide id stays', () => {
    const list =
      '<p:sldIdLst><p:sldId id="256" r:id="rId101"/><p:sldId id="257" r:id="rId102"/>' +
      '<p:sldId id="259" r:id="rId999"/><p:sldId id="258" r:id="rId103"/>' +
      '<p:sldId id="259" r:id="rId104"/></p:sldIdLst>';
    const sections = S('Intro', [256, 257]) + S('Body', [258, 259]);
    const out = presentationOut(variant(sections), variant(sections, list));
    expect(out?.kind).toBe('xml');
    const xml = out!.kind === 'xml' ? out!.xml.root : undefined;
    const entries: string[] = [];
    const walk = (el: typeof xml): void => {
      for (const c of el!.children)
        if (c.type === 'element') {
          if (c.local === 'sldId' && c.ns !== xml!.ns) continue;
          if (c.local === 'sldId') entries.push(c.attrs.map((x) => x.value).join(' '));
          else walk(c);
        }
    };
    walk(xml);
    expect(entries).toEqual(['256 rId101', '257 rId102', '258 rId103', '259 rId104']);
  });

  test('without the presentation relationships, no slide entry is removed', () => {
    const sections = S('Intro', [256, 257]) + S('Body', [258, 259]);
    const list =
      '<p:sldIdLst><p:sldId id="257" r:id="rId102"/><p:sldId id="256" r:id="rId101"/>' +
      '<p:sldId id="258" r:id="rId103"/><p:sldId id="259" r:id="rId104"/></p:sldIdLst>';
    expect(presentationOut(variant(sections), variant(sections, list, true))).toBeUndefined();
  });

  test('a move of a new slide that arrives before the slide itself keeps the slide', async () => {
    const [docA, docC, docD] = [new Y.Doc(), new Y.Doc(), new Y.Doc()];
    docA.clientID = 1;
    docC.clientID = 2;
    docD.clientID = 3;
    const a = await replica(docA, 'a');
    Y.applyUpdate(docC, Y.encodeStateAsUpdate(docA), 'link');
    Y.applyUpdate(docD, Y.encodeStateAsUpdate(docA), 'link');
    const c = await replica(docC, 'c');
    const d = await replica(docD, 'd');
    const capture = (doc: Y.Doc) => {
      const out: Uint8Array[] = [];
      doc.on('update', (u: Uint8Array, origin: unknown) => {
        if (origin !== 'link') out.push(u);
      });
      return out;
    };
    const fromA = capture(docA);
    const fromC = capture(docC);
    const fromD = capture(docD);
    d.goToSlide(0);
    expect(d.exec({ type: 'addSlide' }).ok).toBe(true);
    for (const u of fromD) Y.applyUpdate(docC, u, 'link');
    // C moves D's new slide to the front. The move reaches A before the slide does: A sees the
    // moved entry, whose relationship is not there yet, and Yjs keeps the rest pending.
    expect(c.exec({ type: 'moveSlides', indices: [1], to: 0 }).ok).toBe(true);
    for (const u of fromC) Y.applyUpdate(docA, u, 'link');
    expect(!!docA.store.pendingStructs || !!docA.store.pendingDs).toBe(true);
    for (const u of fromD) Y.applyUpdate(docA, u, 'link');
    for (const u of fromA) Y.applyUpdate(docC, u, 'link');
    for (const u of fromA) Y.applyUpdate(docD, u, 'link');
    for (const u of fromC.slice()) Y.applyUpdate(docD, u, 'link');
    for (const e of [a, c, d]) expect(expectValidSave(e)).toBe(7);
    expectConverged(a, c, d);
  });
});

/**
 * Six slides; slide 1 has a slide jump (`ppaction://hlinksldjump`) to slide 6, in a run's
 * `kind` hyperlink, or only a relationship to slide 6 that no element uses (`kind` empty).
 */
const jumpDeck = (kind = 'hlinkClick', target = 'slide6.xml'): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          text: para(
            run(
              'Go to the end',
              '',
              kind ? `<a:${kind} r:id="rId9" action="ppaction://hlinksldjump"/>` : ''
            )
          ),
        }),
        rels: [
          [
            'rId9',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
            target,
          ],
        ],
      },
      ...Array.from({ length: 5 }, () => ({ shapes: '' })),
    ],
  });

describe('a slide that one peer deletes while another copies a slide that links to it', () => {
  for (const kind of ['hlinkClick', ''])
    for (const deleterFirst of [true, false]) {
      const order = deleterFirst ? 'deleter first' : 'copier first';
      test(`the delete wins: the link goes, and the files stay valid (${kind || 'unused'}, ${order})`, async () => {
        const { sync, a, b } = await pair(deleterFirst, jumpDeck(kind));
        sync.pause();
        expect(a.exec({ type: 'deleteSlides', indices: [5] }).ok).toBe(true);
        expect(b.exec({ type: 'duplicateSlides', indices: [0] }).ok).toBe(true);
        sync.flush();
        expectConverged(a, b);
        for (const e of [a, b]) {
          expect(expectValidSave(e)).toBe(6);
          expect(hasLast(e)).toBe(false);
          const zip = unzipSync(e.save()!);
          for (const [name, data] of Object.entries(zip))
            if (/^ppt\/slides\/slide\d+\.xml$/.test(name))
              expect(strFromU8(data)).not.toContain('hlinksldjump');
        }
        // The deleter's undo brings the slide back; the copy keeps no link to it.
        a.exec({ type: 'undo' });
        expectConverged(a, b);
        for (const e of [a, b]) expect(expectValidSave(e)).toBe(7);
      });
    }
});

describe('a copy of a slide jump that reaches the deleter, while the copier never gets the delete', () => {
  for (const kind of ['hlinkClick', 'hlinkHover', '']) {
    const what = kind ? `an ${kind} jump` : 'an unused slide relationship';
    for (const deleterFirst of [true, false])
      test(`the deleter drops ${what} of the copy (${deleterFirst ? 'deleter' : 'copier'} first)`, async () => {
        const docA = new Y.Doc();
        const docB = new Y.Doc();
        docA.clientID = deleterFirst ? 1 : 2;
        docB.clientID = deleterFirst ? 2 : 1;
        const a = await replica(docA, 'a', jumpDeck(kind));
        Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), 'link');
        const b = await replica(docB, 'b', jumpDeck(kind));
        const fromB: Uint8Array[] = [];
        docB.on('update', (u: Uint8Array, origin: unknown) => {
          if (origin !== 'link') fromB.push(u);
        });
        expect(a.exec({ type: 'deleteSlides', indices: [5] }).ok).toBe(true);
        expect(b.exec({ type: 'duplicateSlides', indices: [0] }).ok).toBe(true);
        // B goes offline: only its copy reaches A.
        for (const u of fromB) Y.applyUpdate(docA, u, 'link');
        expect(expectValidSave(a)).toBe(6);
        const zip = unzipSync(a.save()!);
        for (const [name, data] of Object.entries(zip))
          if (/^ppt\/slides\/(_rels\/)?slide\d+\.xml(\.rels)?$/.test(name)) {
            expect(strFromU8(data)).not.toContain('hlinksldjump');
            expect(strFromU8(data)).not.toContain('slide6.xml');
          }
      });
  }
});

describe('slide jumps that the repair keeps, and one it loses', () => {
  test('a jump whose target differs in ASCII case from the slide part stays on another delete', async () => {
    const { sync, a, b } = await pair(true, jumpDeck('hlinkClick', 'Slide3.xml'));
    sync.pause();
    expect(a.exec({ type: 'deleteSlides', indices: [4] }).ok).toBe(true);
    sync.flush();
    expectConverged(a, b);
    for (const e of [a, b]) {
      expect(expectValidSave(e)).toBe(5);
      expect(strFromU8(unzipSync(e.save()!)['ppt/slides/slide1.xml']!)).toContain('hlinksldjump');
    }
  });

  // Known loss: a replica cannot tell a deleted slide from one whose undo has not
  // reached it yet. B gets C's copy of a jump to slide 3 before A's undo brings slide 3 back,
  // so B drops the copy's jump, and every replica converges without it. Slide 1 gets its jump
  // back from A's undo.
  test('three replicas: a copy that arrives before the undo of its target loses the jump', async () => {
    const docs = [new Y.Doc(), new Y.Doc(), new Y.Doc()] as const;
    docs.forEach((d, i) => (d.clientID = i + 1));
    const [docA, docB, docC] = docs;
    const a = await replica(docA, 'a', jumpDeck('hlinkClick', 'slide3.xml'));
    Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), 'link');
    Y.applyUpdate(docC, Y.encodeStateAsUpdate(docA), 'link');
    const b = await replica(docB, 'b', jumpDeck('hlinkClick', 'slide3.xml'));
    const c = await replica(docC, 'c', jumpDeck('hlinkClick', 'slide3.xml'));
    const sent = new Map<Y.Doc, Uint8Array[]>(docs.map((d) => [d, []]));
    for (const d of docs)
      d.on('update', (u: Uint8Array, origin: unknown) => {
        if (origin !== 'link') sent.get(d)!.push(u);
      });
    const send = (from: Y.Doc, to: Y.Doc) => {
      for (const u of sent.get(from)!) Y.applyUpdate(to, u, 'link');
    };
    expect(a.exec({ type: 'deleteSlides', indices: [2] }).ok).toBe(true);
    send(docA, docB);
    send(docA, docC);
    sent.get(docA)!.length = 0;
    a.exec({ type: 'undo' });
    send(docA, docC);
    expect(c.exec({ type: 'duplicateSlides', indices: [0] }).ok).toBe(true);
    const copy = sent.get(docC)!.slice();
    sent.get(docC)!.length = 0;
    for (const u of copy) Y.applyUpdate(docB, u, 'link');
    // Then every replica gets everything.
    const full = () => {
      for (const from of docs)
        for (const to of docs)
          if (from !== to) Y.applyUpdate(to, Y.encodeStateAsUpdate(from), 'link');
    };
    full();
    full();
    expectConverged(a, b, c);
    for (const e of [a, b, c]) {
      expect(expectValidSave(e)).toBe(7);
      // Slide 1 has its jump again (A's undo); the copy has none.
      const zip = unzipSync(e.save()!);
      const jumps = Object.entries(zip)
        .filter(([n]) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .filter(([, d]) => strFromU8(d).includes('hlinksldjump'))
        .map(([n]) => n);
      expect(jumps).toEqual(['ppt/slides/slide1.xml']);
    }
  });
});
