/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Custom document properties end to end on a server runtime, with a save and reopen.

import { describe, expect, test } from 'bun:test';
import { DocumentPropertyType, isPptxEditorError } from '../index.ts';
import { open, refusedAtSync, rejection, reopen, thrown } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

describe('custom properties', () => {
  test('add of each type survives a save and reopens with its type', async () => {
    const due = new Date('2026-05-06T07:08:09Z');
    const runtime = await open();
    await runtime.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      custom.add('Client', 'Northwind');
      custom.add('Budget', 1250.5);
      custom.add('Approved', true);
      const added = custom.add('Due', due);
      added.load('type,value');
      await context.sync();
      expect([added.type, (added.value as Date).toISOString()]).toEqual([
        DocumentPropertyType.date,
        due.toISOString(),
      ]);
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      custom.load('items');
      const count = custom.getCount();
      await context.sync();
      expect(count.value).toBe(4);
      expect(custom.items.map((p) => [p.key, p.type])).toEqual([
        ['Client', 'String'],
        ['Budget', 'Number'],
        ['Approved', 'Boolean'],
        ['Due', 'Date'],
      ]);
      expect(custom.items[1]!.value).toBe(1250.5);
      expect(custom.items[3]!.value).toBeInstanceOf(Date);
    });
  });

  test('getItem ignores case; a value write changes value and type', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      custom.add('Stage', 'draft');
      await context.sync();
      const stage = custom.getItem('STAGE');
      stage.value = 3;
      await context.sync();
      const read = custom.getItem('stage').load();
      await context.sync();
      expect([read.key, read.type, read.value]).toEqual(['Stage', 'Number', 3]);
    });
  });

  test('lookups that find nothing, delete, and deleteAll', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      custom.add('A', 'one');
      custom.add('B', 'two');
      await context.sync();
      const missing = custom.getItemOrNullObject('Z');
      const at = custom.getItemAt(1).load('key');
      await context.sync();
      expect(missing.isNullObject).toBe(true);
      expect(at.key).toBe('B');
      custom.getItem('A').delete();
      await context.sync();
      const count = custom.getCount();
      await context.sync();
      expect(count.value).toBe(1);
      custom.getItem('Z');
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
      custom.deleteAll();
      const after = custom.getCount();
      await context.sync();
      expect(after.value).toBe(0);
    });
  });

  test('a write on a null object refuses at the call and writes nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      const missing = custom.getItemOrNullObject('Nope');
      await context.sync();
      expect(code(thrown(() => (missing.value = 'x')))).toBe('InvalidObjectPath');
      expect(code(thrown(() => missing.delete()))).toBe('InvalidObjectPath');
      const count = custom.getCount();
      await context.sync();
      expect(count.value).toBe(0);
    });
  });

  test('invalid keys and values refuse at the sync, before anything is sent', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const custom = context.presentation.properties.customProperties;
      expect(code(await refusedAtSync(context, () => custom.add('', 'x')))).toBe('InvalidArgument');
      expect(code(await refusedAtSync(context, () => custom.add('k'.repeat(256), 'x')))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => custom.add('k', Number.NaN)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => custom.add('k', new Date('bad'))))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => custom.add('k', {} as never)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => custom.getItemAt(-1)))).toBe(
        'InvalidArgument'
      );
      expect(code(thrown(() => custom.load('items/nope')))).toBe('InvalidArgument');
      // Characters XML 1.0 cannot carry reach the host and are refused there.
      custom.add('k', 'a\u0001');
      expect(code(await rejection(context.sync()))).toBe('InvalidArgument');
    });
  });

  test('properties load as a navigation of DocumentProperties', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.properties.customProperties.add('X', 'y');
      await context.sync();
      const props = context.presentation.properties.load('customProperties');
      await context.sync();
      const json = props.toJSON() as { customProperties?: { items?: unknown[] } };
      expect(json.customProperties?.items?.length).toBe(1);
    });
  });
});
