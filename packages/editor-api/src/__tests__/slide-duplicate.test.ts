/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Slide.duplicate on a server runtime and on a browser runtime over a live editor: the copy
// follows its source in `p:sldIdLst` (ECMA-376 Part 1, 19.2.1.34) and keeps the layout, the
// master, the notes, the pictures, the charts, and the tags of the source slide.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  presentationInfo,
  readPptx,
  relationshipsOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { chartsDeck } from '../../../../scripts/fixtures/drawingml.ts';
import { isPptxEditorError, PptxEditor, type PptxEditorRuntime, type Slide } from '../browser.ts';
import type { RequestContext } from '../runtime/public.ts';
import { IN, rejection } from './support.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A 1 × 1 PNG. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

const PICTURE =
  '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

/**
 * Two slides. Slide 1 is on the Title Only layout and holds a title, a chart, and a picture,
 * with notes. Slide 2 is blank.
 */
function deck(): Uint8Array {
  const charts = chartsDeck();
  const chart = charts.slides[0]!;
  return buildPptx({
    layouts: [{ name: 'Title Only', type: 'titleOnly', shapes: '' }],
    slides: [
      {
        layoutName: 'Title Only',
        // A title (shape 2), a chart frame (shape 3), and a picture (shape 4).
        shapes: chart.shapes + PICTURE,
        rels: [...(chart.rels ?? []), ['rIdImg', `${R}/image`, '../media/image1.png']],
        notes: 'Speaker notes',
      },
      { layout: 'blank', shapes: '' },
    ],
    extraParts: [charts.extraParts![0]!],
    media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
  });
}

/** The parts slide `index` points at, by relationship type (the last segment of its URI). */
function targets(pkg: PptxPackage, index: number): Record<string, string> {
  const slide = presentationInfo(pkg).slides[index]!;
  return Object.fromEntries(
    relationshipsOf(pkg, slide.partName).map((r) => [
      r.type.slice(r.type.lastIndexOf('/') + 1),
      r.target ?? '',
    ])
  );
}

function read(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('the saved presentation does not read back');
  return r.pkg;
}

/** Each slide's id, layout name, master id, shape names, and tags, in order. */
async function outline(context: RequestContext) {
  const slides = context.presentation.slides.load('items/id');
  await context.sync();
  const parts = slides.items.map((s) => ({
    layout: s.layout.load('name'),
    master: s.slideMaster.load('id'),
    shapes: s.shapes.load('items/name'),
    tags: s.tags.load('items/key,items/value'),
  }));
  await context.sync();
  return slides.items.map((s, i) => ({
    id: s.id,
    layout: parts[i]!.layout.name,
    master: parts[i]!.master.id,
    shapes: parts[i]!.shapes.items.map((x) => x.name),
    tags: parts[i]!.tags.items.map((t) => `${t.key}=${t.value}`),
  }));
}

/** Tag slide 1, then duplicate it and use the copy in the same batch, before and after a sync. */
async function duplicateFirst(runtime: PptxEditorRuntime) {
  await runtime.run(async (context) => {
    context.presentation.slides.getItemAt(0).tags.add('KIND', 'hero');
    await context.sync();
  });
  return runtime.run(async (context) => {
    const source = context.presentation.slides.getItemAt(0).load('id');
    const copy = source.duplicate();
    // Before the sync: a write and a load on the copy join the same batch.
    copy.tags.add('COPY', 'yes');
    copy.load('id,index');
    await context.sync();
    // After the sync: the proxy names the new slide.
    const title = copy.shapes.getItem('2').textFrame.textRange;
    title.text = 'Two';
    title.load('text');
    await context.sync();
    return { source: source.id, id: copy.id, index: copy.index, text: title.text };
  });
}

async function liveEditor(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

/** What both runtimes must answer after `duplicateFirst`. */
function checkCopy(
  answer: Awaited<ReturnType<typeof duplicateFirst>>,
  slides: Awaited<ReturnType<typeof outline>>,
  pkg: PptxPackage
): void {
  expect(answer.index).toBe(1);
  expect(answer.id).not.toBe(answer.source);
  expect(answer.text).toBe('Two');
  expect(slides.map((s) => s.id)).toEqual([answer.source, answer.id, slides[2]!.id]);
  const [source, copy] = slides;
  expect(copy!.layout).toBe('Title Only');
  expect(copy!.layout).toBe(source!.layout);
  expect(copy!.master).toBe(source!.master);
  expect(copy!.shapes).toEqual(source!.shapes);
  expect(source!.tags).toEqual(['KIND=hero']);
  expect(copy!.tags).toEqual(['KIND=hero', 'COPY=yes']);
  const from = targets(pkg, 0);
  const to = targets(pkg, 1);
  // The layout and the picture stay shared; the notes, the chart, and the tags are copies.
  expect(to.slideLayout).toBe(from.slideLayout!);
  expect(to.image).toBe(from.image!);
  for (const own of ['notesSlide', 'chart', 'tags']) {
    expect(to[own]).toBeDefined();
    expect(to[own]).not.toBe(from[own]!);
    expect(pkg.parts.has(to[own]!)).toBe(true);
  }
  expect(presentationInfo(pkg).slides[1]!.notesPart).toBe(to.notesSlide!);
}

describe('Slide.duplicate', () => {
  test('a server runtime copies the slide after its source, and the copy survives a save', async () => {
    const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
    try {
      const answer = await duplicateFirst(runtime);
      const slides = await runtime.run(outline);
      const saved = await runtime.save();
      checkCopy(answer, slides, read(saved));
      const reopened = await PptxEditor.createServer(saved, { substitutes: false });
      try {
        expect(await reopened.run(outline)).toEqual(slides);
      } finally {
        reopened.dispose();
      }
    } finally {
      runtime.dispose();
    }
  });

  test('a browser runtime copies the slide the same way, and one undo removes the copy', async () => {
    const live = await liveEditor();
    const runtime = PptxEditor.createBrowser(live);
    try {
      const answer = await duplicateFirst(runtime);
      const slides = await runtime.run(outline);
      checkCopy(answer, slides, read(live.save()!));
      // Each sync that writes is one undo step: the first undo reverts the title, and the
      // second removes the copy with the tag written to it in the same sync.
      expect(live.exec({ type: 'undo' }).ok).toBe(true);
      expect((await runtime.run(outline)).map((s) => s.id)).toEqual(slides.map((s) => s.id));
      expect(live.exec({ type: 'undo' }).ok).toBe(true);
      const after = await runtime.run(outline);
      expect(after.map((s) => s.id)).toEqual([answer.source, slides[2]!.id]);
      expect(after[0]!.tags).toEqual(['KIND=hero']);
      expect(live.exec({ type: 'redo' }).ok).toBe(true);
      expect((await runtime.run(outline)).map((s) => s.id)).toEqual([
        answer.source,
        answer.id,
        slides[2]!.id,
      ]);
    } finally {
      runtime.dispose();
    }
  });

  test('both runtimes save the same copy', async () => {
    const server = await PptxEditor.createServer(deck(), { substitutes: false });
    const live = await liveEditor();
    const browser = PptxEditor.createBrowser(live);
    try {
      const a = await duplicateFirst(server);
      const b = await duplicateFirst(browser);
      expect(b).toEqual(a);
      expect(await browser.run(outline)).toEqual(await server.run(outline));
    } finally {
      server.dispose();
      browser.dispose();
    }
  });

  test('a slide with an embedded control refuses with NotSupported and changes nothing', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: '',
          rels: [['rIdCtl', `${R}/control`, '../activeX/activeX1.xml']],
        },
      ],
      extraParts: [
        [
          'ppt/activeX/activeX1.xml',
          'application/vnd.ms-office.activeX+xml',
          '<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"/>',
        ],
      ],
    });
    const runtime = await PptxEditor.createServer(bytes, { substitutes: false });
    try {
      const error = await runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).duplicate();
        return rejection(context.sync());
      });
      expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
      expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
        'edit.slide-has-controls'
      );
      expect((await runtime.run(outline)).length).toBe(1);
    } finally {
      runtime.dispose();
    }
  });

  test('a copy of a deleted slide refuses as a move of it does, and changes nothing', async () => {
    const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
    try {
      const refusal = (act: (slide: Slide) => void) =>
        runtime.run(async (context) => {
          const slide = context.presentation.slides.getItemAt(1);
          slide.delete();
          act(slide);
          return rejection(context.sync());
        });
      const code = (error: unknown) => (isPptxEditorError(error) ? error.code : String(error));
      const copied = code(await refusal((slide) => slide.duplicate()));
      expect(copied).toBe(code(await refusal((slide) => slide.moveTo(0))));
      expect(copied).toBe('InvalidObjectPath');
      // The refused batches changed nothing.
      expect((await runtime.run(outline)).length).toBe(2);
    } finally {
      runtime.dispose();
    }
  });
});
