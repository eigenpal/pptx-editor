// `PaintMeter` records what each item paint spends and reads of the slide budgets, and decides
// whether a record still fits them.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { MAX_SLIDE_TILE_IMAGES, paintBudgets } from '../paint-budgets.ts';
import { PaintMeter, type PaintedItem } from '../painted-items.ts';
import { Defs, svg } from '../svg.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

/** A paint that reads the tile budget and spends `n` tiles, as a tiled fill does. */
const readTiles = (defs: Defs, n: number) => () => {
  paintBudgets(defs).tiles.take(() => new Array<number>(n));
  return svg(doc, 'g');
};

/** A paint that reads the band budget, as a shape gradient does: 10 commands. */
const readBands = (defs: Defs) => () => {
  paintBudgets(defs).bands.allow(1, 10);
  return svg(doc, 'g');
};

test('a record that read one budget fits while that budget has as much left', () => {
  const defs = new Defs(doc, 't');
  const first = PaintMeter.resetBudgets(defs);
  const record = first.recordWhole(readTiles(defs, 10));
  expect(record.reads).toEqual({ tiles: true, bands: false });
  expect(record.spent.tiles).toBe(10);
  expect(record.startLeft.tiles).toBe(MAX_SLIDE_TILE_IMAGES);

  // A new paint starts whole budgets: the record fits at the start.
  const next = PaintMeter.resetBudgets(defs);
  expect(next.fits(record)).toBe(true);
  // A budget that it did not read does not matter.
  next.recordWhole(readBands(defs));
  expect(next.fits(record)).toBe(true);
  // A tile spent before its turn makes it wrong.
  next.recordWhole(readTiles(defs, 1));
  expect(next.fits(record)).toBe(false);
});

test('a record that read no budget always fits', () => {
  const defs = new Defs(doc, 't');
  const record = PaintMeter.resetBudgets(defs).recordWhole(() => svg(doc, 'rect'));
  expect(record.reads).toEqual({ tiles: false, bands: false });
  const next = PaintMeter.resetBudgets(defs);
  next.recordWhole(readTiles(defs, 100));
  next.recordWhole(readBands(defs));
  expect(next.fits(record)).toBe(true);
});

test('spendShare spends a kept record once, and never a record that this paint made', () => {
  const defs = new Defs(doc, 't');
  const record = PaintMeter.resetBudgets(defs).recordWhole(readTiles(defs, 10));
  const next = PaintMeter.resetBudgets(defs);
  next.spendShare(record);
  next.spendShare(record);
  expect(paintBudgets(defs).tiles.spent).toBe(10);
  const fresh = next.recordWhole(readTiles(defs, 5));
  next.spendShare(fresh);
  expect(paintBudgets(defs).tiles.spent).toBe(15);
});

test('a group start is partial only when children are left', () => {
  const defs = new Defs(doc, 't');
  const meter = PaintMeter.resetBudgets(defs);
  const some = meter.recordGroupStart(svg(doc, 'g'), 5, () => ({ next: 3, weight: 3 }));
  expect(some.painted.rest).toBe(3);
  expect(some.weight).toBe(4);
  const all = meter.recordGroupStart(svg(doc, 'g'), 5, () => ({ next: 5, weight: 5 }));
  expect('rest' in all.painted).toBe(false);
  expect(all.weight).toBe(6);
});

test('extend starts a first budget read at what was left at the item turn', () => {
  const defs = new Defs(doc, 't');
  const node = svg(doc, 'g');
  const start = PaintMeter.resetBudgets(defs).recordGroupStart(node, 4, () => ({
    next: 1,
    weight: 1,
  }));
  expect(start.painted.reads).toEqual({ tiles: false, bands: false });

  const meter = PaintMeter.resetBudgets(defs);
  meter.recordWhole(readTiles(defs, 100)); // An earlier item spends 100 tiles.
  const more = meter.extend(start.painted, 4, (into, from) => {
    expect(into).toBe(node);
    expect(from).toBe(1);
    readTiles(defs, 7)();
    return { next: 3, weight: 2 };
  });
  const extended = more.painted;
  expect(more.weight).toBe(2);
  expect(extended.rest).toBe(3);
  expect(extended.reads).toEqual({ tiles: true, bands: false });
  expect(extended.startLeft.tiles).toBe(MAX_SLIDE_TILE_IMAGES - 100);
  expect(extended.spent.tiles).toBe(7);
  // The earlier record does not change.
  expect(start.painted.rest).toBe(1);
  expect(start.painted.reads.tiles).toBe(false);
});

test('extend keeps the start of a budget that the record read before', () => {
  const defs = new Defs(doc, 't');
  const node = svg(doc, 'g');
  const start = PaintMeter.resetBudgets(defs).recordGroupStart(node, 4, () => {
    readTiles(defs, 5)();
    return { next: 2, weight: 2 };
  });
  expect(start.painted.startLeft.tiles).toBe(MAX_SLIDE_TILE_IMAGES);
  const meter = PaintMeter.resetBudgets(defs);
  const more = meter.extend(start.painted, 4, () => {
    readTiles(defs, 3)();
    return { next: 4, weight: 2 };
  });
  expect(more.painted.startLeft.tiles).toBe(MAX_SLIDE_TILE_IMAGES);
  expect(more.painted.spent.tiles).toBe(8);
  // The last children painted: the record is whole.
  expect('rest' in more.painted).toBe(false);
  // Its share is spent once: the earlier part when its turn began, the new part as it painted.
  expect(paintBudgets(defs).tiles.spent).toBe(8);
  meter.spendShare(more.painted);
  expect(paintBudgets(defs).tiles.spent).toBe(8);
});

test('extend of a whole record only spends its share', () => {
  const defs = new Defs(doc, 't');
  const record: PaintedItem = PaintMeter.resetBudgets(defs).recordWhole(readTiles(defs, 4));
  const meter = PaintMeter.resetBudgets(defs);
  const more = meter.extend(record, 1, () => {
    throw new Error('a whole record paints nothing more');
  });
  expect(more.painted).toBe(record);
  expect(more.weight).toBe(0);
  expect(paintBudgets(defs).tiles.spent).toBe(4);
});
