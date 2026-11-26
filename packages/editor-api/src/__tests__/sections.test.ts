/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Presentation.sections`, an SDK extension over the section list: a `p14:sectionLst` in a
// `p:ext` (`CT_Extension`) of the presentation's `p:extLst` (`CT_ExtensionList`), ECMA-376
// Part 1, 19.2.1.11 and 19.2.1.12. Each section names its slides by `p:sldId/@id` (19.7.13).
// Another extension in the same list stays through every edit (ECMA-376 Part 3, clause 8).

import { describe, expect, test } from 'bun:test';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { createBrowserAutomationHost, createPptxEditor } from '@pptx-editor.dev/core/editor';
import { presentationInfo, readPptx, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  PptxEditor,
  type PptxEditorRuntime,
  type RequestContext,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTIONS_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const INTRO = '{6E1B33A1-4C0F-4B53-9C2B-0D4A4B3E8A11}';
const BODY = '{0B7F5C2D-1E4A-4F3B-8C6D-2A9E7B1C5D22}';
/** An extension that no part of the engine knows. It must survive every section edit. */
const UNKNOWN_EXT =
  '<p:ext uri="{A1B2C3D4-0000-4000-8000-00000000AB01}"><x:keep xmlns:x="urn:example:keep" v="1"/></p:ext>';

/** `count` slides (ids 256 and up), each with a text box that names it. */
function slides(count: number) {
  return Array.from({ length: count }, (_, i) => ({
    layout: 'blank' as const,
    shapes: shape({
      id: 2,
      name: 'Box',
      x: 914_400,
      y: 914_400,
      cx: 914_400,
      cy: 914_400,
      text: para(run(`Slide ${i + 1}`)),
    }),
  }));
}

/** Four slides in two sections, Intro (256, 257) and Body (258, 259), and an unknown extension. */
function sectioned(): Uint8Array {
  const section = (name: string, id: string, ids: number[]) =>
    `<p14:section name="${name}" id="${id}"><p14:sldIdLst>${ids
      .map((n) => `<p14:sldId id="${n}"/>`)
      .join('')}</p14:sldIdLst></p14:section>`;
  return buildPptx({
    slides: slides(4),
    presentationExt:
      `<p:extLst>${UNKNOWN_EXT}<p:ext uri="${SECTIONS_URI}"><p14:sectionLst xmlns:p14="${P14}">` +
      section('Intro', INTRO, [256, 257]) +
      section('Body', BODY, [258, 259]) +
      '</p14:sectionLst></p:ext></p:extLst>',
  });
}

/** Three slides, no sections, and an unknown extension. */
function plain(): Uint8Array {
  return buildPptx({ slides: slides(3), presentationExt: `<p:extLst>${UNKNOWN_EXT}</p:extLst>` });
}

interface SectionView {
  name: string;
  slideIds: string[];
}

/** The name and slides of each section, in order. */
async function sectionsOf(context: RequestContext): Promise<SectionView[]> {
  const sections = context.presentation.sections.load('items/name,items/slideIds');
  await context.sync();
  return sections.items.map((s) => ({ name: s.name, slideIds: s.slideIds }));
}

async function slideIdsOf(context: RequestContext): Promise<string[]> {
  const all = context.presentation.slides.load('items/id');
  await context.sync();
  return all.items.map((s) => s.id);
}

/** The section names and slide ids in the saved presentation part, and its extension URIs. */
function savedSections(bytes: Uint8Array): { uris: string[]; sections: SectionView[] } {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  const root = xmlPartOf(read.pkg, presentationInfo(read.pkg).presentationPart)!.root;
  const kids = (el: { children: readonly unknown[] }) =>
    el.children.filter(
      (
        c
      ): c is { local: string; ns: string; attrs: { local: string; value: string }[] } & {
        children: readonly unknown[];
      } => (c as { type?: string }).type === 'element'
    );
  const attrOf = (el: { attrs: { local: string; value: string }[] }, name: string) =>
    el.attrs.find((a) => a.local === name)?.value ?? '';
  const extLst = kids(root).find((e) => e.local === 'extLst');
  const exts = extLst ? kids(extLst) : [];
  const list = exts
    .filter((e) => attrOf(e, 'uri') === SECTIONS_URI)
    .flatMap((e) => kids(e))
    .find((e) => e.local === 'sectionLst');
  return {
    uris: exts.map((e) => attrOf(e, 'uri')),
    sections: (list ? kids(list) : []).map((s) => ({
      name: attrOf(s, 'name'),
      slideIds: kids(kids(s).find((e) => e.local === 'sldIdLst')!).map((e) => attrOf(e, 'id')),
    })),
  };
}

const UNKNOWN_URI = '{A1B2C3D4-0000-4000-8000-00000000AB01}';

describe('Presentation.sections: reads', () => {
  test('a load reads each section in order, with its id, name, index, and slides', async () => {
    const runtime = await open(sectioned());
    const read = await runtime.run(async (context) => {
      const sections = context.presentation.sections.load('items');
      await context.sync();
      return sections.items.map((s) => s.toJSON());
    });
    expect(read).toEqual([
      { id: INTRO, name: 'Intro', index: 0, slideIds: ['256', '257'] },
      { id: BODY, name: 'Body', index: 1, slideIds: ['258', '259'] },
    ]);
    runtime.dispose();
  });

  test('a presentation without sections has an empty collection', async () => {
    const runtime = await open(plain());
    const count = await runtime.run(async (context) => {
      const n = context.presentation.sections.getCount();
      await context.sync();
      return n.value;
    });
    expect(count).toBe(0);
    runtime.dispose();
  });

  test('getItem finds a section by id in either case, and getItemOrNullObject finds none', async () => {
    const runtime = await open(sectioned());
    const [name, isNull] = await runtime.run(async (context) => {
      const body = context.presentation.sections.getItem(BODY.toLowerCase()).load('name');
      const none = context.presentation.sections.getItemOrNullObject(
        '{00000000-0000-0000-0000-000000000000}'
      );
      await context.sync();
      return [body.name, none.isNullObject] as const;
    });
    expect(name).toBe('Body');
    expect(isNull).toBe(true);
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => context.presentation.sections.getItemAt(5).load('name'))
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    runtime.dispose();
  });

  test('a load of the presentation can name its sections', async () => {
    const runtime = await open(sectioned());
    const names = await runtime.run(async (context) => {
      context.presentation.load('sections/items/name');
      await context.sync();
      return context.presentation.sections.items.map((s) => s.name);
    });
    expect(names).toEqual(['Intro', 'Body']);
    runtime.dispose();
  });
});

describe('Presentation.sections: writes', () => {
  test('add in a presentation without sections adds a default section first, and survives a save', async () => {
    const runtime = await open(plain());
    const added = await runtime.run(async (context) => {
      const section = context.presentation.sections.add('Results', '257');
      section.load('id,name,index,slideIds');
      await context.sync();
      return section.toJSON();
    });
    expect(added).toMatchObject({ name: 'Results', index: 1, slideIds: ['257', '258'] });
    expect(added.id).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/
    );
    const saved = savedSections(await runtime.save());
    expect(saved.sections).toEqual([
      { name: 'Default Section', slideIds: ['256'] },
      { name: 'Results', slideIds: ['257', '258'] },
    ]);
    // The new section list goes first in the extension list, and the unknown extension stays.
    expect(saved.uris).toEqual([SECTIONS_URI, UNKNOWN_URI]);
    const again = await reopen(runtime);
    expect(await again.run(sectionsOf)).toEqual(saved.sections);
    runtime.dispose();
    again.dispose();
  });

  test('add splits a section, and the new section works in the same batch', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(3);
      const section = context.presentation.sections.add('Close', slide);
      section.name = 'Closing';
      await context.sync();
    });
    expect(await runtime.run(sectionsOf)).toEqual([
      { name: 'Intro', slideIds: ['256', '257'] },
      { name: 'Body', slideIds: ['258'] },
      { name: 'Closing', slideIds: ['259'] },
    ]);
    runtime.dispose();
  });

  test('add at the first slide of a section leaves that section empty', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.add('Main', '258');
      await context.sync();
    });
    expect(await runtime.run(sectionsOf)).toEqual([
      { name: 'Intro', slideIds: ['256', '257'] },
      { name: 'Body', slideIds: [] },
      { name: 'Main', slideIds: ['258', '259'] },
    ]);
    runtime.dispose();
  });

  test('a name write renames the section, and the current name writes nothing', async () => {
    const runtime = await open(sectioned());
    const before = await runtime.save();
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(INTRO).name = 'Intro';
      await context.sync();
    });
    expect(savedSections(await runtime.save())).toEqual(savedSections(before));
    await runtime.run(async (context) => {
      context.presentation.sections.getItemAt(1).name = 'Line one\nline two';
      await context.sync();
    });
    // A section name is one line: a line break becomes a space.
    expect((await runtime.run(sectionsOf))[1]!.name).toBe('Line one line two');
    runtime.dispose();
  });

  test('moveTo moves a section with its slides, and keeps their order', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.add('Close', '259');
      await context.sync();
      context.presentation.sections.getItem(BODY).moveTo(0);
      await context.sync();
    });
    await runtime.run(async (context) => {
      expect(await sectionsOf(context)).toEqual([
        { name: 'Body', slideIds: ['258'] },
        { name: 'Intro', slideIds: ['256', '257'] },
        { name: 'Close', slideIds: ['259'] },
      ]);
      expect(await slideIdsOf(context)).toEqual(['258', '256', '257', '259']);
      // A move of two places, down.
      context.presentation.sections.getItemAt(0).moveTo(2);
      await context.sync();
      expect(await slideIdsOf(context)).toEqual(['256', '257', '259', '258']);
    });
    const again = await reopen(runtime);
    expect((await again.run(sectionsOf)).map((s) => s.name)).toEqual(['Intro', 'Close', 'Body']);
    runtime.dispose();
    again.dispose();
  });

  test('delete keeps the slides in the section before it, or after it for the first section', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(INTRO).delete();
      await context.sync();
      expect(await sectionsOf(context)).toEqual([
        { name: 'Body', slideIds: ['256', '257', '258', '259'] },
      ]);
      expect(await slideIdsOf(context)).toEqual(['256', '257', '258', '259']);
    });
    runtime.dispose();
  });

  test('delete with deleteSlides deletes the slides too', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(BODY).delete({ deleteSlides: true });
      await context.sync();
      expect(await sectionsOf(context)).toEqual([{ name: 'Intro', slideIds: ['256', '257'] }]);
      expect(await slideIdsOf(context)).toEqual(['256', '257']);
    });
    const again = await reopen(runtime);
    expect(await again.run(slideIdsOf)).toEqual(['256', '257']);
    runtime.dispose();
    again.dispose();
  });

  test('deleting the last section removes the section list and keeps the other extension', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(INTRO).delete();
      context.presentation.sections.getItem(BODY).delete();
      await context.sync();
    });
    const saved = savedSections(await runtime.save());
    expect(saved).toEqual({ uris: [UNKNOWN_URI], sections: [] });
    runtime.dispose();
  });

  test('deleteAll removes every section and keeps the slides and the other extension', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.deleteAll();
      await context.sync();
      expect(await sectionsOf(context)).toEqual([]);
      expect(await slideIdsOf(context)).toEqual(['256', '257', '258', '259']);
      // With no sections, deleteAll changes nothing.
      context.presentation.sections.deleteAll();
      await context.sync();
    });
    expect(savedSections(await runtime.save()).uris).toEqual([UNKNOWN_URI]);
    runtime.dispose();
  });
});

describe('Presentation.sections: refusals', () => {
  /** The refusal of `call`, and the sections after it: a refused batch writes nothing. */
  async function refusal(call: (context: RequestContext) => void) {
    const runtime = await open(sectioned());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        // A valid write before the refused call: the batch is atomic, so it must not apply.
        context.presentation.sections.getItem(INTRO).name = 'Changed';
        call(context);
      })
    );
    const after = await runtime.run(sectionsOf);
    runtime.dispose();
    expect(after.map((s) => s.name)).toEqual(['Intro', 'Body']);
    if (!isPptxEditorError(error)) throw error;
    return { code: error.code, inner: error.debugInfo.innerError?.code };
  }

  test('an empty or long name refuses with InvalidArgument and edit.section-name', async () => {
    expect(await refusal((c) => (c.presentation.sections.getItem(BODY).name = ' '))).toEqual({
      code: 'InvalidArgument',
      inner: 'edit.section-name',
    });
    expect(await refusal((c) => c.presentation.sections.add('x'.repeat(256), '257'))).toEqual({
      code: 'InvalidArgument',
      inner: 'edit.section-name',
    });
  });

  test('arguments of the wrong type refuse with InvalidArgument', async () => {
    const sections = (c: RequestContext) => c.presentation.sections;
    for (const call of [
      (c: RequestContext) => ((sections(c).getItemAt(0) as { name: unknown }).name = 7),
      (c: RequestContext) => sections(c).add(7 as unknown as string, '257'),
      (c: RequestContext) => sections(c).add('A', 257 as unknown as string),
      (c: RequestContext) => sections(c).getItemAt(0).moveTo(-1),
      (c: RequestContext) => sections(c).getItemAt(0).moveTo(1.5),
      (c: RequestContext) =>
        sections(c)
          .getItemAt(0)
          .delete({ deleteSlide: true } as never),
      (c: RequestContext) =>
        sections(c)
          .getItemAt(0)
          .delete({ deleteSlides: 'yes' } as never),
    ])
      expect((await refusal(call)).code).toBe('InvalidArgument');
  });

  test('a position past the last section refuses with InvalidArgument', async () => {
    expect((await refusal((c) => c.presentation.sections.getItem(INTRO).moveTo(2))).code).toBe(
      'InvalidArgument'
    );
  });

  test('an unknown slide or section refuses with ItemNotFound', async () => {
    expect((await refusal((c) => c.presentation.sections.add('A', '9999'))).code).toBe(
      'ItemNotFound'
    );
    expect(
      (
        await refusal((c) =>
          c.presentation.sections.getItem('{00000000-0000-0000-0000-000000000000}').delete()
        )
      ).code
    ).toBe('ItemNotFound');
  });

  test('deleting the slides of a section that holds every slide refuses with NotSupported', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(INTRO).delete();
      await context.sync();
    });
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        context.presentation.sections.getItemAt(0).delete({ deleteSlides: true })
      )
    );
    expect(isPptxEditorError(error) && [error.code, error.debugInfo.innerError?.code]).toEqual([
      'NotSupported',
      'edit.last-slide',
    ]);
    expect(await runtime.run(slideIdsOf)).toHaveLength(4);
    runtime.dispose();
  });

  test('a section that an earlier sync deleted refuses a write with InvalidObjectPath', async () => {
    const runtime = await open(sectioned());
    const error = await runtime.run(async (context) => {
      const body = context.presentation.sections.getItem(BODY);
      body.load('id');
      await context.sync();
      body.delete();
      await context.sync();
      body.name = 'Gone';
      return context.sync().then(
        () => undefined,
        (e: unknown) => e
      );
    });
    expect(isPptxEditorError(error) && error.code).toBe('InvalidObjectPath');
    expect(await runtime.run(sectionsOf)).toEqual([
      { name: 'Intro', slideIds: ['256', '257', '258', '259'] },
    ]);
    runtime.dispose();
  });
});

describe('Presentation.sections: slide calls keep the sections in step', () => {
  test('SlideCollection.add puts the new slide at the end of the last section', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.slides.add();
      await context.sync();
      const ids = await slideIdsOf(context);
      expect((await sectionsOf(context))[1]!.slideIds).toEqual(ids.slice(2));
    });
    runtime.dispose();
  });

  test('Slide.moveTo puts the slide into the section of the slide before it, or the first section', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.slides.getItem('256').moveTo(2);
      await context.sync();
      expect(await sectionsOf(context)).toEqual([
        { name: 'Intro', slideIds: ['257'] },
        { name: 'Body', slideIds: ['258', '256', '259'] },
      ]);
      context.presentation.slides.getItem('259').moveTo(0);
      await context.sync();
      expect(await sectionsOf(context)).toEqual([
        { name: 'Intro', slideIds: ['259', '257'] },
        { name: 'Body', slideIds: ['258', '256'] },
      ]);
    });
    runtime.dispose();
  });

  test('Slide.duplicate puts the copy into the section of its source', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      const copy = context.presentation.slides.getItem('257').duplicate().load('id');
      await context.sync();
      expect((await sectionsOf(context))[0]).toEqual({
        name: 'Intro',
        slideIds: ['256', '257', copy.id],
      });
    });
    runtime.dispose();
  });

  test('Slide.delete takes the slide out of its section, and an empty section stays', async () => {
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.slides.getItem('258').delete();
      context.presentation.slides.getItem('259').delete();
      await context.sync();
      expect(await sectionsOf(context)).toEqual([
        { name: 'Intro', slideIds: ['256', '257'] },
        { name: 'Body', slideIds: [] },
      ]);
    });
    const again = await reopen(runtime);
    expect((await again.run(sectionsOf)).map((s) => s.name)).toEqual(['Intro', 'Body']);
    runtime.dispose();
    again.dispose();
  });

  test('insertSlidesFromBase64 puts the slides into the section of targetSlideId', async () => {
    const source = await PptxEditor.createServer(plain(), { substitutes: false });
    const base64 = await source.run(async (context) => {
      const bytes = context.presentation.slides.getItemAt(0).exportAsBase64();
      await context.sync();
      return bytes.value;
    });
    source.dispose();
    const runtime = await open(sectioned());
    await runtime.run(async (context) => {
      context.presentation.insertSlidesFromBase64(base64, { targetSlideId: '256' });
      await context.sync();
      const ids = await slideIdsOf(context);
      expect(ids).toHaveLength(5);
      expect((await sectionsOf(context))[0]).toEqual({
        name: 'Intro',
        slideIds: ['256', ids[1]!, '257'],
      });
      context.presentation.insertSlidesFromBase64(base64);
      await context.sync();
      const now = await slideIdsOf(context);
      expect((await sectionsOf(context))[0]!.slideIds[0]).toBe(now[0]!);
    });
    runtime.dispose();
  });
});

describe('Presentation.sections: runtimes', () => {
  test('the browser runtime gives the same result as the server runtime, as one undo step', async () => {
    const edit = async (runtime: PptxEditorRuntime) => {
      await runtime.run(async (context) => {
        const section = context.presentation.sections.add('Close', '259');
        section.name = 'Closing';
        section.moveTo(0);
        await context.sync();
      });
      return runtime.run(sectionsOf);
    };
    const server = await open(sectioned());
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(sectioned())).ok).toBe(true);
    const browser = BrowserEditor.createBrowser(live);
    const fromServer = await edit(server);
    expect(await edit(browser)).toEqual(fromServer);
    expect(fromServer.map((s) => s.name)).toEqual(['Closing', 'Intro', 'Body']);
    expect(savedSections(live.save()!)).toEqual(savedSections(await server.save()));
    expect(live.snapshot().sections!.map((s) => s.name)).toEqual(['Closing', 'Intro', 'Body']);
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(live.snapshot().sections!.map((s) => s.name)).toEqual(['Intro', 'Body']);
    expect(live.snapshot().canUndo).toBe(false);
    server.dispose();
  });

  test('the calls send the lookups and one section operation each', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-sections.test.ts`) sends these
    // operations to two linked editors; this test keeps them equal to what the calls send.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(sectioned())).ok).toBe(true);
    const inner = createBrowserAutomationHost(live);
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: false }).run(async (context) => {
      const sections = context.presentation.sections;
      sections.add('Close', '259');
      sections.getItemAt(1).name = 'Main';
      sections.getItemAt(1).moveTo(0);
      sections.getItemAt(0).delete({ deleteSlides: true });
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { id: '259' } },
        { op: 'addSection', name: 'Close', slide: { result: 1 } },
        { op: 'getSection', key: { index: 1 } },
        { op: 'renameSection', section: { result: 3 }, name: 'Main' },
        { op: 'getSection', key: { index: 1 } },
        { op: 'moveSection', section: { result: 5 }, index: 0 },
        { op: 'getSection', key: { index: 0 } },
        { op: 'deleteSection', section: { result: 7 }, deleteSlides: true },
      ],
    ]);
  });

  test('a name that trims to the current name writes nothing and adds no undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(sectioned())).ok).toBe(true);
    await BrowserEditor.createBrowser(live).run(async (context) => {
      context.presentation.sections.getItem(INTRO).name = '  Intro  ';
      await context.sync();
    });
    expect(live.snapshot().canUndo).toBe(false);
  });
});

describe('Presentation.sections: unusual section lists', () => {
  /** A deck of `count` slides whose section list is `sections` (raw `p14:section` XML). */
  function withList(count: number, sections: string): Uint8Array {
    return buildPptx({
      slides: slides(count),
      presentationExt: `<p:extLst><p:ext uri="${SECTIONS_URI}"><p14:sectionLst xmlns:p14="${P14}">${sections}</p14:sectionLst></p:ext></p:extLst>`,
    });
  }
  const section = (name: string, id: string, ids: number[]) =>
    `<p14:section name="${name}" id="${id}"><p14:sldIdLst>${ids
      .map((n) => `<p14:sldId id="${n}"/>`)
      .join('')}</p14:sldIdLst></p14:section>`;
  const ID = (n: number) => `{00000000-0000-4000-8000-00000000000${n}}`;

  test('moveTo steps past empty sections in both directions, and the order survives a save', async () => {
    const runtime = await open(
      withList(
        3,
        section('A', ID(1), [256]) +
          section('Empty1', ID(2), []) +
          section('B', ID(3), [257]) +
          section('Empty2', ID(4), []) +
          section('C', ID(5), [258])
      )
    );
    await runtime.run(async (context) => {
      context.presentation.sections.getItem(ID(1)).moveTo(4);
      await context.sync();
      expect(await sectionsOf(context)).toEqual([
        { name: 'Empty1', slideIds: [] },
        { name: 'B', slideIds: ['257'] },
        { name: 'Empty2', slideIds: [] },
        { name: 'C', slideIds: ['258'] },
        { name: 'A', slideIds: ['256'] },
      ]);
      context.presentation.sections.getItem(ID(4)).moveTo(0);
      await context.sync();
    });
    const expected = [
      { name: 'Empty2', slideIds: [] },
      { name: 'Empty1', slideIds: [] },
      { name: 'B', slideIds: ['257'] },
      { name: 'C', slideIds: ['258'] },
      { name: 'A', slideIds: ['256'] },
    ];
    expect(await runtime.run(sectionsOf)).toEqual(expected);
    expect(await runtime.run(slideIdsOf)).toEqual(['257', '258', '256']);
    const again = await reopen(runtime);
    expect(await again.run(sectionsOf)).toEqual(expected);
    runtime.dispose();
    again.dispose();
  });

  test('a repeated section id, an unknown slide id, and a slide in no section read as the writer saves them', async () => {
    const runtime = await open(
      withList(
        4,
        section('A', ID(1), [256, 999]) + section('Twin', ID(1), [257]) + section('B', ID(2), [258])
      )
    );
    const read = await runtime.run(sectionsOf);
    // Each slide is in one section once, and the read matches the saved section list.
    expect(read.flatMap((s) => s.slideIds).sort()).toEqual(['256', '257', '258', '259']);
    expect(read).toEqual(savedSections(await runtime.save()).sections);
    // A write repairs the list first, then edits it.
    await runtime.run(async (context) => {
      context.presentation.sections.getItemAt(0).name = 'First';
      await context.sync();
    });
    const after = await runtime.run(sectionsOf);
    expect(after[0]!.name).toBe('First');
    expect(after).toEqual(savedSections(await runtime.save()).sections);
    runtime.dispose();
  });
});
