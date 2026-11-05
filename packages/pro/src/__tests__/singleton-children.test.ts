/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The pure rule of `singleton-children.ts`, as DOCX Editor's `singleton-children.test.ts`:
// slots, schema order, seed keys, and the limit on concurrent copies.

import { describe, expect, test } from 'bun:test';
import {
  fromSeed,
  groupOf,
  inSchemaOrder,
  withCompanions,
  type MergeView,
} from '../collaboration/yjs/singleton-children.ts';
import { ROOM_LIMITS } from '../collaboration/yjs/schema.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

describe('groupOf', () => {
  const slot = (parent: { ns: string; local: string }, child: { ns: string; local: string }) =>
    groupOf(parent, child, () => undefined)?.slot;
  test('names one slot for a choice group and none for repeated children', () => {
    const spPr = { ns: NS_P, local: 'spPr' };
    expect(slot(spPr, { ns: NS_A, local: 'solidFill' })).toBe('fill');
    expect(slot(spPr, { ns: NS_A, local: 'noFill' })).toBe('fill');
    expect(slot(spPr, { ns: NS_A, local: 'effectLst' })).toBe('effect');
    // Table cells: one fill and one of each border per cell.
    const tcPr = { ns: NS_A, local: 'tcPr' };
    expect(slot(tcPr, { ns: NS_A, local: 'solidFill' })).toBe('fill');
    expect(slot(tcPr, { ns: NS_A, local: 'lnL' })).toBe('lnL');
    expect(slot({ ns: NS_A, local: 'tr' }, { ns: NS_A, local: 'tc' })).toBeUndefined();
    expect(slot({ ns: NS_P, local: 'blipFill' }, { ns: NS_A, local: 'srcRect' })).toBe('srcRect');
    expect(slot({ ns: NS_P, local: 'spTree' }, { ns: NS_P, local: 'sp' })).toBeUndefined();
    expect(slot({ ns: NS_A, local: 'txBody' }, { ns: NS_A, local: 'p' })).toBeUndefined();
    expect(slot({ ns: NS_P, local: 'cSld' }, { ns: NS_P, local: 'bg' })).toBe('bg');
    // Content types: one Default per extension and one Override per part name, without case.
    const types = {
      ns: 'http://schemas.openxmlformats.org/package/2006/content-types',
      local: 'Types',
    };
    const def = { ns: types.ns, local: 'Default' };
    const over = { ns: types.ns, local: 'Override' };
    expect(groupOf(types, def, () => 'PNG')?.slot).toBe(groupOf(types, def, () => 'png')?.slot);
    expect(groupOf(types, over, () => '/ppt/media/Image1.png')?.slot).toBe(
      groupOf(types, over, () => '/PPT/media/image1.png')?.slot
    );
    expect(groupOf(types, def, () => 'png')?.slot).not.toBe(
      groupOf(types, over, () => 'png')?.slot
    );
    // The fold is ASCII only, as `partNameKey`: the Kelvin sign (U+212A) is no `k`.
    expect(groupOf(types, def, () => '\u212Amz')?.slot).not.toBe(
      groupOf(types, def, () => 'kmz')?.slot
    );
    // A relationship id is compared as written (ECMA-376 Part 2): rId1 and RID1 are two.
    const rels = {
      ns: 'http://schemas.openxmlformats.org/package/2006/relationships',
      local: 'Relationships',
    };
    const rel = { ns: rels.ns, local: 'Relationship' };
    expect(groupOf(rels, rel, () => 'rId1')?.slot).not.toBe(groupOf(rels, rel, () => 'RID1')?.slot);
    // A DrawingML graphic frame (GVML) is not PresentationML's: no order or slot from p:.
    expect(slot({ ns: NS_A, local: 'graphicFrame' }, { ns: NS_A, local: 'xfrm' })).toBeUndefined();
    expect(slot({ ns: NS_A, local: 'bodyPr' }, { ns: NS_A, local: 'flatTx' })).toBe('text3d');
    expect(slot({ ns: NS_A, local: 'rPr' }, { ns: NS_A, local: 'uFill' })).toBe('uFill');
    // Prototype keys are not slots.
    expect(slot({ ns: NS_A, local: 'constructor' }, { ns: NS_A, local: 'x' })).toBeUndefined();
    expect(slot(spPr, { ns: NS_A, local: 'constructor' })).toBeUndefined();
    expect(slot({ ns: 'urn:other', local: 'spPr' }, { ns: NS_A, local: 'ln' })).toBeUndefined();
  });
});

describe('fromSeed', () => {
  test('only the seed marks its keys', () => {
    expect(fromSeed('2._1')).toBe(true);
    // Base-36 counters can start with any letter: a minted `2.s1` is no seed key.
    expect(fromSeed('2.s1')).toBe(false);
    expect(fromSeed('2.1c')).toBe(false);
  });
});

describe('inSchemaOrder', () => {
  const spPr = { ns: NS_P, local: 'spPr' };
  const el = (local: string, ns = NS_A) => ({ ns, local });
  const name = (c: { ns: string; local: string }) => c;
  test('sorts known children and keeps an unknown child after its neighbor', () => {
    const ext = el('thing', 'urn:ext');
    const out = inSchemaOrder(spPr, [el('xfrm'), el('effectLst'), ext, el('ln')], name);
    expect(out.map((c) => c.local)).toEqual(['xfrm', 'ln', 'effectLst', 'thing']);
  });
  test('puts paragraph properties first and keeps the runs in order', () => {
    const out = inSchemaOrder(el('p'), [el('r'), el('br'), el('pPr'), el('endParaRPr')], name);
    expect(out.map((c) => c.local)).toEqual(['pPr', 'r', 'br', 'endParaRPr']);
  });
  test('returns the same list when it is in order, and leaves other parents alone', () => {
    const list = [el('xfrm'), el('prstGeom'), el('ln')];
    expect(inSchemaOrder(spPr, list, name)).toBe(list);
    const shapes = [el('sp', NS_P), el('nvGrpSpPr', NS_P)];
    expect(inSchemaOrder(el('spTree', NS_P), shapes, name)).toBe(shapes);
  });
});

describe('companion limits', () => {
  const NS = NS_A;
  const p = { ns: NS, local: 'p' };
  const copies = (n: number) =>
    Array.from({ length: n }, (_, i) => ({
      key: `${(i + 1).toString(36)}.1`,
      ns: NS,
      local: 'pPr',
      nsDecls: [],
      attrs: [{ ns: '', local: `x${i}`, value: '1' }],
      children: [],
    }));

  test('many concurrent copies of one property container merge in one pass', () => {
    const n = ROOM_LIMITS.maxCopiesPerGroup + 1;
    const views = new Map<string, MergeView>();
    const start = performance.now();
    const out = withCompanions(p, copies(n), views);
    expect(performance.now() - start).toBeLessThan(500);
    expect(out).toHaveLength(1);
    const merged = out![0]!;
    expect('attrs' in merged ? merged.attrs : []).toHaveLength(n);
    expect(views.get('1.1')!.companions).toHaveLength(n - 1);
  });

  test('a group over the limit reads as null (the mirror refuses the part), so a hostile peer cannot stall every read', () => {
    const start = performance.now();
    expect(withCompanions(p, copies(ROOM_LIMITS.maxCopiesPerGroup + 2), new Map())).toBeNull();
    expect(withCompanions(p, copies(20_000), new Map())).toBeNull();
    expect(performance.now() - start).toBeLessThan(500);
  });
});
