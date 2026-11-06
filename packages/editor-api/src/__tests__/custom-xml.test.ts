/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Custom XML parts end to end on a server runtime, on every owner kind, with a save and reopen.

import { describe, expect, test } from 'bun:test';
import { isPptxEditorError, type CustomXmlPartCollection, type RequestContext } from '../index.ts';
import { open, refusedAtSync, rejection, reopen, thrown } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

/** The collection of each owner kind, by name. */
function owners(context: RequestContext): Record<string, CustomXmlPartCollection> {
  const slide = context.presentation.slides.getItemAt(0);
  return {
    presentation: context.presentation.customXmlParts,
    slide: slide.customXmlParts,
    layout: slide.layout.customXmlParts,
    master: slide.slideMaster.customXmlParts,
    shape: slide.shapes.getItemAt(0).customXmlParts,
  };
}

describe('custom XML parts', () => {
  test('add on every owner survives a save; getXml returns the text', async () => {
    const runtime = await open();
    const ids: Record<string, string> = {};
    await runtime.run(async (context) => {
      for (const [name, parts] of Object.entries(owners(context))) {
        const part = parts.add(`<data xmlns="urn:test:${name}"><v>${name}</v></data>`);
        ids[name] = part.id;
      }
      await context.sync();
    });
    for (const id of Object.values(ids)) expect(id).toMatch(GUID);
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const all = owners(context);
      const reads = Object.entries(all).map(([name, parts]) => {
        parts.load('items');
        return { name, parts, xml: parts.getItem(ids[name]!).getXml() };
      });
      await context.sync();
      for (const { name, parts, xml } of reads) {
        expect(parts.items.map((p) => [p.id, p.namespaceUri])).toEqual([
          [ids[name], `urn:test:${name}`],
        ]);
        expect(xml.value).toBe(`<data xmlns="urn:test:${name}"><v>${name}</v></data>`);
      }
    });
  });

  test('getByNamespace scopes the list; getOnlyItem needs exactly one', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const parts = context.presentation.customXmlParts;
      parts.add('<a xmlns="urn:one"/>');
      parts.add('<b xmlns="urn:two"/>');
      parts.add('<c xmlns="urn:two"/>');
      await context.sync();
      const one = parts.getByNamespace('urn:one');
      const two = parts.getByNamespace('urn:two');
      const count = two.getCount();
      const only = one.getOnlyItem().load('namespaceUri');
      const none = two.getOnlyItemOrNullObject();
      const at = two.getItemAt(1).load('namespaceUri');
      two.load('items/id');
      await context.sync();
      expect(count.value).toBe(2);
      expect(only.namespaceUri).toBe('urn:one');
      expect(none.isNullObject).toBe(true);
      expect(at.namespaceUri).toBe('urn:two');
      expect(two.items).toHaveLength(2);
      two.getOnlyItem();
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
    });
  });

  test('setXml and delete on an added part in the same sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const parts = context.presentation.slides.getItemAt(0).customXmlParts;
      const part = parts.add('<x xmlns="urn:x">1</x>');
      part.setXml('<y xmlns="urn:y">2</y>');
      const xml = part.getXml();
      part.load('namespaceUri');
      await context.sync();
      expect([xml.value, part.namespaceUri]).toEqual(['<y xmlns="urn:y">2</y>', 'urn:y']);
      part.delete();
      const count = parts.getCount();
      await context.sync();
      expect(count.value).toBe(0);
    });
  });

  test('lookups that find nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const parts = context.presentation.customXmlParts;
      const missing = parts.getItemOrNullObject('{00000000-0000-4000-8000-000000000000}');
      await context.sync();
      expect(missing.isNullObject).toBe(true);
      expect(code(thrown(() => missing.setXml('<a/>')))).toBe('InvalidObjectPath');
      parts.getItem('00000000-0000-4000-8000-000000000000');
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
    });
  });

  test('untrusted XML refuses with InvalidArgument and writes nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const parts = context.presentation.customXmlParts;
      expect(code(await refusedAtSync(context, () => parts.add('')))).toBe('InvalidArgument');
      expect(code(await refusedAtSync(context, () => parts.add(42 as never)))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => parts.getItem('not-a-guid')))).toBe(
        'InvalidArgument'
      );
      expect(code(await refusedAtSync(context, () => parts.getByNamespace(1 as never)))).toBe(
        'InvalidArgument'
      );
      // Resolve the presentation first: a refused batch also refuses the lookups in it.
      parts.getCount();
      await context.sync();
      for (const xml of ['<a>', '<!DOCTYPE a [<!ENTITY e "x">]><a>&e;</a>', '<a/><b/>']) {
        parts.add(xml);
        const error = await rejection(context.sync());
        expect(code(error)).toBe('InvalidArgument');
        expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
          'edit.invalid-custom-xml'
        );
      }
      const count = parts.getCount();
      await context.sync();
      expect(count.value).toBe(0);
    });
  });
});
