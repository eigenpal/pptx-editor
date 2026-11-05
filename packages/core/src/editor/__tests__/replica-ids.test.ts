// New relationship ids, part names, slide ids, and shape ids: a file's dense numbers when
// the editor works alone, the actor's stripe while a collaboration session is attached, and
// never a duplicate of an id the file already holds, however large.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import { treeOf } from '../../automation/commands/slide-structure.ts';
import { attr, elements, readPptx, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor } from '../index.ts';
import { CommandRefusal, OpBatch, type MintedIds } from '../../automation/commands/op-batch.ts';
import {
  addRelationship,
  freePartName,
  nextShapeId,
  shapeIds,
  spTreeOf,
  stripedIds,
} from '../../automation/commands/package-edits.ts';
import { ACTOR_ID_STRIPE, actorStripe } from '../../automation/commands/actor-scoped-ids.ts';
import { bindReplicaIds, replicaBindingOf } from '../replica-ids.ts';
import { resetDevWarnings } from '../dev-warnings.ts';
import type { EditorCollaborationSession } from '../modules.ts';

const IN = 914_400;
const SLIDE = '/ppt/slides/slide1.xml';
const RELS = '/ppt/slides/_rels/slide1.xml.rels';
const deck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('x')) }),
      },
    ],
  });

function pkgWith(relIds: readonly string[]) {
  const r = readPptx(deck());
  if (!r.ok) throw new Error('unreadable');
  // Add relationships with the given ids to the slide's rels, as a file from another tool can.
  const batch = new OpBatch(r.pkg);
  const root = xmlPartOf(batch.pkg, RELS)!.root;
  for (const id of relIds)
    batch.push({
      op: 'insertNode',
      part: RELS,
      parent: root.id,
      index: root.children.length,
      spec: {
        ns: 'http://schemas.openxmlformats.org/package/2006/relationships',
        local: 'Relationship',
        attrs: [
          { ns: '', local: 'Id', value: id },
          { ns: '', local: 'Type', value: 'http://example.com/x' },
          { ns: '', local: 'Target', value: 'x.xml' },
        ],
      },
    });
  return batch.pkg;
}

describe('new ids', () => {
  test('alone: dense names, the next number after the largest', () => {
    const batch = new OpBatch(pkgWith(['rId7']));
    expect(addRelationship(batch, SLIDE, 'http://example.com/y', '/ppt/y.xml')).toBe('rId8');
    expect(freePartName(batch, '/ppt/slides/slide', '.xml')).toBe('/ppt/slides/slide2.xml');
    expect(nextShapeId(batch, spTreeOf(batch, SLIDE))).toBe(3);
  });

  test('an rId too large to count from never gives a duplicate or an invalid id', () => {
    const huge = 'rId99999999999999999999';
    const batch = new OpBatch(pkgWith([huge, 'rId2']));
    const id = addRelationship(batch, SLIDE, 'http://example.com/y', '/ppt/y.xml');
    expect(id).toMatch(/^rId\d{1,9}$/);
    const ids = elements(xmlPartOf(batch.pkg, RELS)!.root).map((r) => attr(r, 'Id'));
    expect(new Set(ids).size).toBe(ids.length);
  });

  test("in a session: every new number is in the actor's stripe, and never minted twice", () => {
    const actor = 'k3f2-1';
    const stripe = actorStripe(actor);
    const inStripe = (n: number) => n % ACTOR_ID_STRIPE === stripe;
    const minted = new Map() as MintedIds;
    const batch = new OpBatch(pkgWith(['rId7']), { allocationActor: actor, minted });
    const ids = [1, 2].map((i) =>
      Number(addRelationship(batch, SLIDE, `http://example.com/${i}`, `/ppt/${i}.xml`).slice(3))
    );
    expect(ids.every(inStripe)).toBe(true);
    expect(ids[1]).toBe(ids[0]! + ACTOR_ID_STRIPE);
    // Part names take the full stripe too.
    const name = freePartName(batch, '/ppt/slides/slide', '.xml');
    const n = Number(/slide(\d+)\.xml$/.exec(name)![1]);
    expect(n % ACTOR_ID_STRIPE).toBe(stripe);
    // Several shapes at once (paste, duplicate) stay in the stripe, not in a run from the first.
    const next = shapeIds(batch, spTreeOf(batch, SLIDE));
    const shapes = [next(), next(), next()];
    expect(shapes.every((x) => inStripe(x) && x < 2 ** 31)).toBe(true);
    // A number handed out once is not handed out again in the session, even when the tree no
    // longer holds it (a peer deleted that shape, and its undo can bring it back).
    const fresh = new OpBatch(batch.pkg, { allocationActor: actor, minted });
    expect(nextShapeId(fresh, spTreeOf(fresh, SLIDE))).toBe(shapes[2]! + ACTOR_ID_STRIPE);
    expect(
      Number(addRelationship(fresh, SLIDE, 'http://example.com/3', '/ppt/3.xml').slice(3))
    ).not.toBe(ids[0]);
    // Each id kind has its own record, so shape ids never use up relationship numbers.
    const kinds = new OpBatch(pkgWith([]), { allocationActor: actor, minted: new Map() });
    const shapeSource = shapeIds(kinds, spTreeOf(kinds, SLIDE));
    for (let i = 0; i < 3; i += 1) shapeSource();
    expect(addRelationship(kinds, SLIDE, 'http://example.com/k', '/ppt/k.xml')).toBe(
      `rId${stripe}`
    );
    // Numbers that are taken are skipped, and the cursor stops at its ceiling, never wrapping.
    expect(stripe).toBeGreaterThan(0);
    const N = ACTOR_ID_STRIPE;
    const cursor = stripedIds(actor, new Set([stripe + N]), 1, stripe + 2 * N);
    expect([cursor(), cursor()]).toEqual([stripe, stripe + 2 * N]);
    expect(() => cursor()).toThrow(CommandRefusal);
  });

  test('a can() probe uses up no number of the stripe', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(deck())).ok).toBe(true);
    // A stand-in collaboration session: only its id decides the stripe.
    bindReplicaIds(e.session!.store, {
      sessionId: 'k3f2-1',
    } as unknown as EditorCollaborationSession);
    for (let i = 0; i < 50; i += 1) expect(e.can({ type: 'addSlide' }).ok).toBe(true);
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    const pres = xmlPartOf(e.session!.store.package, '/ppt/presentation.xml')!;
    const ids: number[] = [];
    const walk = (el: (typeof pres)['root']): void => {
      if (el.local === 'sldId') ids.push(Number(attr(el, 'id')));
      for (const c of elements(el)) walk(c);
    };
    walk(pres.root);
    const first = stripedIds('k3f2-1', new Set([ids[0]!]), 256, 2_147_483_647)();
    expect(ids[1]).toBe(first);
  });

  test('an editor with no collaboration module keeps the plain names', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(deck())).ok).toBe(true);
    // A picture alone writes a `Default` for its extension, not an override.
    const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);
    expect(
      e.exec({ type: 'insertImage', bytes: png, mime: 'image/png', width: 1, height: 1 }).ok
    ).toBe(true);
    const ct = elements(xmlPartOf(e.session!.store.package, '/[Content_Types].xml')!.root);
    expect(ct.filter((x) => x.local === 'Default' && attr(x, 'Extension') === 'png')).toHaveLength(
      1
    );
    expect(ct.some((x) => x.local === 'Override' && /media/.test(attr(x, 'PartName') ?? ''))).toBe(
      false
    );
    expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(e.session!.info.slides.map((s) => s.partName)).toEqual([
      SLIDE,
      '/ppt/slides/slide2.xml',
    ]);
  });
});

describe('comment indexes of a slide copy in a session', () => {
  test("the copies take free indexes of the actor's stripe after lastIdx, never one in use", async () => {
    const actor = 'k3f2-1';
    const stripe = actorStripe(actor);
    // The first number of the stripe after `lastIdx="2"`: the file already uses it.
    const taken = 3 + ((stripe - 3 + ACTOR_ID_STRIPE) % ACTOR_ID_STRIPE);
    const files = unzipSync(buildSlideStructureDeck());
    const name = 'ppt/comments/comment1.xml';
    files[name] = strToU8(strFromU8(files[name]!).replace('idx="1"', `idx="${taken}"`));
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(zipSync(files))).ok).toBe(true);
    bindReplicaIds(e.session!.store, { sessionId: actor } as unknown as EditorCollaborationSession);
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const pkg = e.session!.store.package;
    const idx = [...pkg.parts.keys()]
      .filter((n) => /\/comments\/comment\d+\.xml$/.test(n))
      .flatMap((n) => elements(treeOf(pkg, n)!.root).map((cm) => Number(attr(cm, 'idx'))));
    expect(idx).toEqual([taken, 2, taken + ACTOR_ID_STRIPE, taken + 2 * ACTOR_ID_STRIPE]);
    const author = elements(treeOf(pkg, '/ppt/commentAuthors.xml')!.root)[0];
    expect(attr(author, 'lastIdx')).toBe(String(taken + 2 * ACTOR_ID_STRIPE));
  });
});

describe('replica binding', () => {
  test('an empty session id binds no stripe and prints a development warning', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [{ layout: 'blank', shapes: '' }] }))).ok).toBe(true);
    resetDevWarnings();
    const warn = console.warn;
    const seen: unknown[] = [];
    console.warn = (...args: unknown[]) => void seen.push(args[0]);
    try {
      bindReplicaIds(e.session!.store, { sessionId: '' } as unknown as EditorCollaborationSession);
    } finally {
      console.warn = warn;
    }
    expect(replicaBindingOf(e.session!.store)).toBeUndefined();
    expect(String(seen[0])).toContain('empty sessionId');
  });
});
