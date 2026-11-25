/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Slide.speakerNotes`, an SDK extension over the body placeholder (`p:ph type="body"`, ECMA-376
// Part 1, 19.3.1.36) of the slide's notes slide (`p:notes`, `CT_NotesSlide`, 19.3.1.26) in its
// Notes Slide part (13.3.5), which relates to the Notes Master part (13.3.4).

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  notesText,
} from '@pptx-editor.dev/core/editor';
import {
  presentationInfo,
  readPptx,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import { isPptxEditorError, type PptxEditorRuntime, type RequestContext } from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

const NOTES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const MASTER_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster';

/**
 * Three slides. The first has notes with a bold run, a line break, a second paragraph, and an
 * extension list the engine does not know; the second has plain notes; the third has none.
 */
function withNotes(): Uint8Array {
  const box = (n: number) =>
    shape({
      id: 2,
      name: 'Box',
      x: 914_400,
      y: 914_400,
      cx: 914_400,
      cy: 914_400,
      text: para(run(`Slide ${n}`)),
    });
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: box(1),
        notesXml:
          '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Open</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> with the goal</a:t></a:r>' +
          '<a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US"/><a:t>then pause</a:t></a:r></a:p>' +
          '<a:p><a:pPr><a:extLst><a:ext uri="{00000000-0000-0000-0000-00000000A173}"><x:keep xmlns:x="urn:example:keep" v="1"/></a:ext></a:extLst></a:pPr>' +
          '<a:r><a:rPr lang="en-US"/><a:t>Source: survey</a:t></a:r></a:p>',
      },
      { layout: 'blank', shapes: box(2), notes: 'Plain notes' },
      { layout: 'blank', shapes: box(3) },
    ],
  });
}

/** Load `speakerNotes` of every slide. */
async function notesOf(context: RequestContext): Promise<string[]> {
  const slides = context.presentation.slides.load('items/speakerNotes');
  await context.sync();
  return slides.items.map((s) => s.speakerNotes);
}

/** Read saved bytes, or fail. */
function read(bytes: Uint8Array) {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('the saved presentation does not read back');
  return r.pkg;
}

/** The XML of the notes slide of slide `index`, or undefined. */
function notesXmlOf(bytes: Uint8Array, index: number): string | undefined {
  const pkg = read(bytes);
  const part = presentationInfo(pkg).slides[index]?.notesPart;
  const xml = part ? xmlPartOf(pkg, part) : undefined;
  return xml ? serializeXmlPart(xml) : undefined;
}

/** `bytes` with the notes master and every notes slide removed. */
function withoutNotesMaster(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes);
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(files)) {
    if (name.includes('notesMaster') || name.includes('notesSlide')) continue;
    let text = strFromU8(data);
    if (name === 'ppt/presentation.xml')
      text = text.replace(/<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/, '');
    if (name.endsWith('.rels') || name === '[Content_Types].xml')
      text = text.replace(/<(Relationship|Override)[^>]*(notesMaster|notesSlide)[^>]*\/>/g, '');
    out[name] = strToU8(text);
  }
  return zipSync(out);
}

describe('Slide.speakerNotes', () => {
  test('a load reads the notes text, and "" for a slide with no notes slide', async () => {
    const runtime = await open(withNotes());
    const notes = await runtime.run(notesOf);
    // Paragraphs join with `\n`, and a line break (`a:br`) reads as `\n` too.
    expect(notes).toEqual(['Open with the goal\nthen pause\nSource: survey', 'Plain notes', '']);
    runtime.dispose();
  });

  test('a load of every scalar includes speakerNotes', async () => {
    const runtime = await open(withNotes());
    const json = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(1).load();
      await context.sync();
      return slide.toJSON();
    });
    expect(json).toMatchObject({ speakerNotes: 'Plain notes', hidden: false });
    runtime.dispose();
  });

  test('a write changes only the text that differs, and keeps formatting and unknown content', async () => {
    const runtime = await open(withNotes());
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      slide.speakerNotes = 'Open with the goal\nthen pause\nSource: 2026 survey';
      await context.sync();
      slide.load('speakerNotes');
      await context.sync();
      expect(slide.speakerNotes).toBe('Open with the goal\nthen pause\nSource: 2026 survey');
    });
    const bytes = await runtime.save();
    const xml = notesXmlOf(bytes, 0)!;
    expect(xml).toContain('<a:rPr lang="en-US" b="1"/><a:t>Open</a:t>');
    expect(xml).toContain('<a:br>');
    expect(xml).toContain('<x:keep xmlns:x="urn:example:keep" v="1"/>');
    expect(xml).toContain('2026 survey');
    const again = await reopen(runtime);
    expect((await again.run(notesOf))[0]).toBe(
      'Open with the goal\nthen pause\nSource: 2026 survey'
    );
    runtime.dispose();
    again.dispose();
  });

  test('each line of a write is a paragraph; CRLF, CR, and vertical tab end a line too', async () => {
    const runtime = await open(withNotes());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).speakerNotes = 'one\ntwo\r\nthree\rfour\vfive';
      await context.sync();
    });
    const pkg = read(await runtime.save());
    expect(notesText(pkg, 1)).toBe('one\ntwo\nthree\nfour\nfive');
    expect(notesXmlOf(await runtime.save(), 1)!.match(/<a:p>/g)).toHaveLength(5);
    runtime.dispose();
  });

  test('a write to a slide with no notes slide creates one from the notes master', async () => {
    const runtime = await open(withNotes());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(2).speakerNotes = 'Talk track\nSecond point';
      await context.sync();
      expect(await notesOf(context)).toEqual([
        'Open with the goal\nthen pause\nSource: survey',
        'Plain notes',
        'Talk track\nSecond point',
      ]);
    });
    const bytes = await runtime.save();
    const pkg = read(bytes);
    const slide = presentationInfo(pkg).slides[2]!;
    const notesPart = slide.notesPart!;
    expect(relationshipsOf(pkg, slide.partName).filter((r) => r.type === NOTES_REL)).toHaveLength(
      1
    );
    const master = presentationInfo(pkg).notesMasterPart!;
    expect(relationshipsOf(pkg, notesPart).find((r) => r.type === MASTER_REL)?.target).toBe(master);
    expect(notesXmlOf(bytes, 2)).toContain('<p:ph type="body" idx="1"/>');
    const again = await reopen(runtime);
    expect((await again.run(notesOf))[2]).toBe('Talk track\nSecond point');
    runtime.dispose();
    again.dispose();
  });

  test('a write of "" to a slide with no notes slide creates nothing, and clears existing notes', async () => {
    const runtime = await open(withNotes());
    const parts = read(await runtime.save()).parts.size;
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(2).speakerNotes = '';
      context.presentation.slides.getItemAt(1).speakerNotes = '';
      await context.sync();
      expect(await notesOf(context)).toEqual([
        'Open with the goal\nthen pause\nSource: survey',
        '',
        '',
      ]);
    });
    const pkg = read(await runtime.save());
    expect(pkg.parts.size).toBe(parts);
    // A cleared notes slide stays, with its body placeholder.
    expect(presentationInfo(pkg).slides[1]!.notesPart).toBeDefined();
    expect(presentationInfo(pkg).slides[2]!.notesPart).toBeUndefined();
    runtime.dispose();
  });

  test('a write of the current text changes nothing', async () => {
    const runtime = await open(withNotes());
    const before = notesXmlOf(await runtime.save(), 0);
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0).load('speakerNotes');
      await context.sync();
      slide.speakerNotes = slide.speakerNotes;
      await context.sync();
    });
    expect(notesXmlOf(await runtime.save(), 0)).toBe(before);
    runtime.dispose();
  });

  test('a value that is not a string refuses with InvalidArgument at the sync', async () => {
    const runtime = await open(withNotes());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        (context.presentation.slides.getItemAt(0) as { speakerNotes: unknown }).speakerNotes = 42;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    expect((await runtime.run(notesOf))[0]).toBe('Open with the goal\nthen pause\nSource: survey');
    runtime.dispose();
  });

  test('text that XML cannot hold refuses the whole sync with InvalidArgument', async () => {
    const runtime = await open(withNotes());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItemAt(1).speakerNotes = 'Kept out';
        context.presentation.slides.getItemAt(2).speakerNotes = 'bell \u0007';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'argument.not-accepted'
    );
    // No partial write: the valid write of the same sync did not apply either.
    expect(await runtime.run(notesOf)).toEqual([
      'Open with the goal\nthen pause\nSource: survey',
      'Plain notes',
      '',
    ]);
    runtime.dispose();
  });

  test('a presentation with no notes master refuses a write with NotSupported', async () => {
    const runtime = await open(withoutNotesMaster(withNotes()));
    expect(await runtime.run(notesOf)).toEqual(['', '', '']);
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItemAt(0).speakerNotes = 'Talk track';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.no-notes-master'
    );
    // An empty text needs no notes slide, so it writes nothing and succeeds.
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).speakerNotes = '';
      await context.sync();
    });
    runtime.dispose();
  });

  test('a notes slide with no body placeholder refuses a write with NotSupported', async () => {
    const files = unzipSync(withNotes());
    const name = 'ppt/notesSlides/notesSlide2.xml';
    files[name] = strToU8(strFromU8(files[name]!).replace('type="body"', 'type="dt"'));
    const runtime = await open(zipSync(files));
    expect((await runtime.run(notesOf))[1]).toBe('');
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItemAt(1).speakerNotes = 'Talk track';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.no-notes-placeholder'
    );
    runtime.dispose();
  });

  test('replaceAll and speakerNotes read and write the same notes', async () => {
    const runtime = await open(withNotes());
    const notes = await runtime.run(async (context) => {
      const count = context.presentation.replaceAll('survey', 'study');
      await context.sync();
      expect(count.value).toBe(1);
      context.presentation.slides.getItemAt(1).speakerNotes = 'Plain notes about the study';
      const second = context.presentation.slides.getItemAt(1).replaceAll('study', 'trial');
      await context.sync();
      expect(second.value).toBe(1);
      return notesOf(context);
    });
    expect(notes).toEqual([
      'Open with the goal\nthen pause\nSource: study',
      'Plain notes about the trial',
      '',
    ]);
    runtime.dispose();
  });

  test('a duplicate has its own notes: a write to one slide never changes the other', async () => {
    const runtime = await open(withNotes());
    const notes = await runtime.run(async (context) => {
      const source = context.presentation.slides.getItemAt(1);
      const copy = source.duplicate();
      copy.load('speakerNotes');
      await context.sync();
      expect(copy.speakerNotes).toBe('Plain notes');
      copy.speakerNotes = 'Copy notes';
      const empty = context.presentation.slides.getItemAt(3).duplicate();
      empty.speakerNotes = 'New notes on a copy';
      await context.sync();
      return notesOf(context);
    });
    expect(notes).toEqual([
      'Open with the goal\nthen pause\nSource: survey',
      'Plain notes',
      'Copy notes',
      '',
      'New notes on a copy',
    ]);
    const pkg = read(await runtime.save());
    const parts = presentationInfo(pkg).slides.map((s) => s.notesPart);
    expect(new Set(parts.filter(Boolean)).size).toBe(4);
    runtime.dispose();
  });

  test('a write to a slide that does not exist refuses with ItemNotFound', async () => {
    const runtime = await open(withNotes());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItem('9999').speakerNotes = 'x';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    runtime.dispose();
  });

  test('the browser runtime reads and writes as the server runtime does, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(withNotes())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const server = await open(withNotes());
    expect(await runtime.run(notesOf)).toEqual(await server.run(notesOf));
    const write = async (context: RequestContext) => {
      context.presentation.slides.getItemAt(1).speakerNotes = 'Edited';
      context.presentation.slides.getItemAt(2).speakerNotes = 'Created';
      await context.sync();
      return notesOf(context);
    };
    expect(await runtime.run(write)).toEqual(await server.run(write));
    expect(notesXmlOf(live.save()!, 2)).toBe(notesXmlOf(await server.save(), 2));
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    const pkg = read(live.save()!);
    expect([notesText(pkg, 1), presentationInfo(pkg).slides[2]!.notesPart]).toEqual([
      'Plain notes',
      undefined,
    ]);
    expect(live.snapshot().canUndo).toBe(false);
    server.dispose();
  });

  test('a write sends the lookups and one setSlideNotes operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-slide-notes.test.ts`) sends these
    // operations to two linked editors; this test keeps them equal to what the property sends.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(withNotes())).ok).toBe(true);
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
      context.presentation.slides.getItemAt(2).speakerNotes = 'Talk track';
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 2 } },
        { op: 'setSlideNotes', slide: { result: 1 }, text: 'Talk track' },
      ],
    ]);
  });

  test('the example of the slides and shapes page runs as written', async () => {
    // docs/site/content/editor-api/slides-and-shapes.mdx, "Read and write speaker notes".
    async function addSource(runtime: PptxEditorRuntime, position: number, source: string) {
      return runtime.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(position).load('speakerNotes');
        await context.sync();
        const lines = slide.speakerNotes ? [slide.speakerNotes] : [];
        slide.speakerNotes = [...lines, `Source: ${source}`].join('\n');
        await context.sync();
        slide.load('speakerNotes');
        await context.sync();
        return slide.speakerNotes;
      });
    }
    const runtime = await open(withNotes());
    expect(await addSource(runtime, 1, 'field interviews')).toBe(
      'Plain notes\nSource: field interviews'
    );
    expect(await addSource(runtime, 2, 'field interviews')).toBe('Source: field interviews');
    runtime.dispose();
  });
});
