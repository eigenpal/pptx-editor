/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Slide.headerFooter` and `Presentation.applyHeaderFooter`, SDK extensions over the slide's own
// `sldNum`, `ftr`, and `dt` placeholders (`p:ph`, `CT_Placeholder`, ECMA-376 Part 1, 19.3.1.36;
// `ST_PlaceholderType`, 19.7.10). The number is a `slidenum` field and an automatic date a
// `datetime1` field (`a:fld`, `CT_TextField`, 21.1.2.2.4).

import { describe, expect, test } from 'bun:test';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  presentationInfo,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { headerFooterDeck } from '../../../../scripts/fixtures/round-trip.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  type PptxEditorRuntime,
  type RequestContext,
  type SlideHeaderFooter,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

/** Today's date as the engine writes an automatic date: month/day/year. */
const today = (): string => {
  const now = new Date();
  return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
};

/** The header-footer deck: slide 1 has a footer and a number, slide 2 has neither. */
const twoSlides = (): Uint8Array => buildPptx(headerFooterDeck());

/** Slide 1 on a title slide layout (`p:sldLayout type="title"`), then two content slides. */
function withTitleSlide(): Uint8Array {
  const deck = headerFooterDeck();
  return buildPptx({
    ...deck,
    layouts: [{ name: 'Title Slide', type: 'title', shapes: '' }],
    slides: [{ ...deck.slides[0]!, layoutName: 'Title Slide' }, deck.slides[1]!, deck.slides[1]!],
  });
}

const FOOTER_PH =
  '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Footer Placeholder 8"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';

/**
 * A master with no footer, date, or number placeholder. Slide 1 is on a layout that has a footer
 * placeholder; slide 2 is on a layout with none.
 */
function oneLayoutWithFooter(): Uint8Array {
  const body = shape({
    id: 2,
    name: 'Box',
    x: 0,
    y: 0,
    cx: 914_400,
    cy: 914_400,
    text: para(run('S')),
  });
  return buildPptx({
    layouts: [{ name: 'With Footer', shapes: FOOTER_PH }],
    slides: [
      { layoutName: 'With Footer', shapes: body },
      { layout: 'blank', shapes: body },
    ],
  });
}

const PROPS = 'slideNumberVisible,footerVisible,footerText,dateVisible,dateText,dateFixed';

/** The settings of every slide. */
async function settingsOf(context: RequestContext) {
  const slides = context.presentation.slides.load('items');
  await context.sync();
  const all = slides.items.map((s) => s.headerFooter.load(PROPS));
  await context.sync();
  return all.map((hf: SlideHeaderFooter) => hf.toJSON());
}

/** The XML of slide `index` of `bytes`. */
function slideXml(bytes: Uint8Array, index: number): string {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  const slide = presentationInfo(read.pkg).slides[index]!;
  return serializeXmlPart(xmlPartOf(read.pkg, slide.partName)!);
}

/** The texts that the live editor lays out on slide `index`, line by line. */
const painted = (e: PptxEditorInstance, index: number): string[] =>
  e
    .session!.layout(index)!
    .items.flatMap((item) =>
      item.kind === 'shape' && item.text
        ? item.text.lines.map((l) => l.runs.map((r) => r.text).join(''))
        : []
    );

const NONE = {
  slideNumberVisible: false,
  footerVisible: false,
  footerText: '',
  dateVisible: false,
  dateText: '',
  dateFixed: false,
};

describe('Slide.headerFooter: read', () => {
  test('a load reads what each slide shows', async () => {
    const runtime = await open(twoSlides());
    expect(await runtime.run(settingsOf)).toEqual([
      { ...NONE, slideNumberVisible: true, footerVisible: true, footerText: 'Footer text & more' },
      NONE,
    ]);
    runtime.dispose();
  });

  test('a load of the slide can name headerFooter', async () => {
    const runtime = await open(twoSlides());
    const json = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0).load('headerFooter');
      await context.sync();
      return slide.toJSON();
    });
    expect(json).toMatchObject({ headerFooter: { slideNumberVisible: true, footerVisible: true } });
    runtime.dispose();
  });
});

describe('Slide.headerFooter: write one slide', () => {
  test('shows the number, a footer, and a fixed date on one slide only, through save and reopen', async () => {
    const runtime = await open(twoSlides());
    await runtime.run(async (context) => {
      const hf = context.presentation.slides.getItemAt(1).headerFooter;
      hf.slideNumberVisible = true;
      hf.footerText = 'Only two';
      hf.dateText = 'Q3 2026';
      await context.sync();
    });
    const again = await reopen(runtime);
    expect(await again.run(settingsOf)).toEqual([
      { ...NONE, slideNumberVisible: true, footerVisible: true, footerText: 'Footer text & more' },
      {
        slideNumberVisible: true,
        footerVisible: true,
        footerText: 'Only two',
        dateVisible: true,
        dateText: 'Q3 2026',
        dateFixed: true,
      },
    ]);
    const xml = slideXml(await runtime.save(), 1);
    expect(xml).toMatch(/<a:fld id="\{[0-9A-F-]{36}\}" type="slidenum">/);
    expect(xml).toContain('<p:ph type="ftr" sz="quarter" idx="11"/>');
    expect(xml).not.toContain('datetime');
    runtime.dispose();
    again.dispose();
  });

  test('hides the number and the footer of one slide', async () => {
    const runtime = await open(twoSlides());
    await runtime.run(async (context) => {
      const hf = context.presentation.slides.getItemAt(0).headerFooter;
      hf.slideNumberVisible = false;
      hf.footerVisible = false;
      await context.sync();
    });
    expect(await runtime.run(settingsOf)).toEqual([NONE, NONE]);
    expect(slideXml(await runtime.save(), 0)).not.toMatch(/type="(sldNum|ftr)"/);
    runtime.dispose();
  });

  test('footerVisible true spreads the footer text of another slide', async () => {
    const runtime = await open(twoSlides());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).headerFooter.footerVisible = true;
      await context.sync();
    });
    const [, second] = await runtime.run(settingsOf);
    expect(second).toMatchObject({ footerVisible: true, footerText: 'Footer text & more' });
    runtime.dispose();
  });

  test('dateFixed false gives an automatic date of today, and dateText then fixes it', async () => {
    const runtime = await open(twoSlides());
    const read = (context: RequestContext) => {
      const hf = context.presentation.slides.getItemAt(0).headerFooter.load(PROPS);
      return context.sync().then(() => hf.toJSON());
    };
    const automatic = await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).headerFooter.dateFixed = false;
      await context.sync();
      return read(context);
    });
    expect(automatic).toMatchObject({ dateVisible: true, dateText: today(), dateFixed: false });
    expect(slideXml(await runtime.save(), 0)).toMatch(
      /type="datetime1"><a:rPr lang="en-US"\/><a:t>/
    );
    const fixed = await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).headerFooter.dateText = '7 November';
      await context.sync();
      return read(context);
    });
    expect(fixed).toMatchObject({ dateVisible: true, dateText: '7 November', dateFixed: true });
    // An automatic date again: fixed text becomes today's date as a field.
    const back = await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).headerFooter.dateFixed = false;
      await context.sync();
      return read(context);
    });
    expect(back).toMatchObject({ dateText: today(), dateFixed: false });
    runtime.dispose();
  });
});

describe('Presentation.applyHeaderFooter', () => {
  test('applies to all slides and skips the title slide', async () => {
    const runtime = await open(withTitleSlide());
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({
        slideNumberVisible: true,
        footerText: 'Board review',
        dateText: 'June 2026',
        hideOnTitleSlides: true,
      });
      await context.sync();
    });
    const shown = {
      slideNumberVisible: true,
      footerVisible: true,
      footerText: 'Board review',
      dateVisible: true,
      dateText: 'June 2026',
      dateFixed: true,
    };
    expect(await runtime.run(settingsOf)).toEqual([NONE, shown, shown]);
    runtime.dispose();
  });

  test('a setting left out stays as each slide has it', async () => {
    const runtime = await open(twoSlides());
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({ dateVisible: true });
      await context.sync();
    });
    const [first, second] = await runtime.run(settingsOf);
    expect(first).toMatchObject({ slideNumberVisible: true, footerText: 'Footer text & more' });
    expect(second).toMatchObject({ slideNumberVisible: false, footerVisible: false });
    expect([first!.dateText, second!.dateText]).toEqual([today(), today()]);
    runtime.dispose();
  });

  test('skips a slide whose layout and master have no placeholder', async () => {
    const runtime = await open(oneLayoutWithFooter());
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({ footerText: 'Kept' });
      await context.sync();
    });
    const [first, second] = await runtime.run(settingsOf);
    expect(first).toMatchObject({ footerVisible: true, footerText: 'Kept' });
    expect(second).toEqual(NONE);
    runtime.dispose();
  });

  test('the read of one slide applied to all slides copies its settings', async () => {
    const runtime = await open(twoSlides());
    for (const index of [0, 1]) {
      await runtime.run(async (context) => {
        const hf = context.presentation.slides.getItemAt(index).headerFooter.load(PROPS);
        await context.sync();
        context.presentation.applyHeaderFooter(hf.toJSON());
        await context.sync();
      });
      const [first, second] = await runtime.run(settingsOf);
      expect(second).toEqual(first!);
    }
    expect((await runtime.run(settingsOf))[1]).toMatchObject({ footerText: 'Footer text & more' });
    runtime.dispose();
  });

  test('dateText with dateFixed false is an automatic date field that holds the text', async () => {
    const runtime = await open(twoSlides());
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({ dateText: '6/1/2026', dateFixed: false });
      await context.sync();
    });
    const [first] = await runtime.run(settingsOf);
    expect(first).toMatchObject({ dateVisible: true, dateText: '6/1/2026', dateFixed: false });
    expect(slideXml(await runtime.save(), 0)).toContain('type="datetime1"');
    runtime.dispose();
  });

  test('hiding what no slide can show changes nothing', async () => {
    const runtime = await open(oneLayoutWithFooter());
    const before = await runtime.save();
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({ slideNumberVisible: false, dateVisible: false });
      context.presentation.slides.getItemAt(1).headerFooter.footerVisible = false;
      await context.sync();
    });
    for (const i of [0, 1]) expect(slideXml(await runtime.save(), i)).toBe(slideXml(before, i));
    runtime.dispose();
  });
});

describe('refusals', () => {
  test('a slide that cannot show the item refuses with NotSupported, and nothing changes', async () => {
    const runtime = await open(oneLayoutWithFooter());
    const before = await runtime.save();
    for (const write of [
      (hf: SlideHeaderFooter) => (hf.footerText = 'No'),
      (hf: SlideHeaderFooter) => (hf.slideNumberVisible = true),
      (hf: SlideHeaderFooter) => (hf.dateVisible = true),
    ]) {
      const error = await runtime.run(async (context) => {
        write(context.presentation.slides.getItemAt(1).headerFooter);
        return context.sync().then(
          () => undefined,
          (e: unknown) => e
        );
      });
      expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
    }
    const footer = await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).headerFooter.footerText = 'No';
      return context.sync().then(
        () => undefined,
        (e: unknown) => e
      );
    });
    expect(isPptxEditorError(footer) && footer.debugInfo.innerError?.code).toBe(
      'edit.no-footer-placeholder'
    );
    expect(slideXml(await runtime.save(), 1)).toBe(slideXml(before, 1));
    runtime.dispose();
  });

  test('a call refuses as a whole when no slide can show one item', async () => {
    const runtime = await open(oneLayoutWithFooter());
    const error = await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({ footerText: 'Kept', slideNumberVisible: true });
      return context.sync().then(
        () => undefined,
        (e: unknown) => e
      );
    });
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.no-slide-number-placeholder'
    );
    // The footer of slide 1 did not apply either: one call is one write.
    expect((await runtime.run(settingsOf))[0]).toEqual(NONE);
    runtime.dispose();
  });

  test('invalid settings refuse with InvalidArgument at the sync, and write nothing', async () => {
    const runtime = await open(twoSlides());
    const cases: ((context: RequestContext) => void)[] = [
      (c) => c.presentation.applyHeaderFooter({}),
      (c) => c.presentation.applyHeaderFooter({ footerVisible: false, footerText: 'X' }),
      (c) => c.presentation.applyHeaderFooter({ dateVisible: false, dateText: 'X' }),
      (c) => c.presentation.applyHeaderFooter({ dateText: '' }),
      (c) => c.presentation.applyHeaderFooter({ footerText: ' ' }),
      (c) => c.presentation.applyHeaderFooter({ footerText: 'two\nlines' }),
      (c) => c.presentation.applyHeaderFooter({ footerText: 'x'.repeat(1025) }),
      (c) => c.presentation.applyHeaderFooter({ footerText: 'bad \u0001' }),
      (c) => c.presentation.applyHeaderFooter({ slideNumber: true } as never),
      (c) => c.presentation.applyHeaderFooter({ slideNumberVisible: 'yes' } as never),
      (c) => c.presentation.applyHeaderFooter(null as never),
      (c) => {
        (c.presentation.slides.getItemAt(0).headerFooter as { footerText: unknown }).footerText = 3;
      },
      (c) => {
        c.presentation.slides.getItemAt(0).headerFooter.dateText = '';
      },
    ];
    for (const call of cases) {
      const error = await runtime.run((context) => refusedAtSync(context, () => call(context)));
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
    // A valid write before an invalid one in the same sync does not apply.
    await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItemAt(1).headerFooter.slideNumberVisible = true;
        context.presentation.applyHeaderFooter({ footerText: '' });
      })
    );
    expect((await runtime.run(settingsOf))[1]).toEqual(NONE);
    runtime.dispose();
  });

  test('a write to a slide that does not exist refuses with ItemNotFound', async () => {
    const runtime = await open(twoSlides());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItem('9999').headerFooter.slideNumberVisible = true;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    runtime.dispose();
  });
});

describe('browser runtime', () => {
  async function live(bytes: Uint8Array): Promise<PptxEditorInstance> {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(bytes)).ok).toBe(true);
    return editor;
  }

  test('writes as the server runtime does, paints the fields, and undoes in one step', async () => {
    const editor = await live(withTitleSlide());
    const runtime = BrowserEditor.createBrowser(editor);
    await runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({
        slideNumberVisible: true,
        footerText: 'Board review',
        dateVisible: true,
        hideOnTitleSlides: true,
      });
      await context.sync();
    });
    const server = await open(withTitleSlide());
    await server.run(async (context) => {
      context.presentation.applyHeaderFooter({
        slideNumberVisible: true,
        footerText: 'Board review',
        dateVisible: true,
        hideOnTitleSlides: true,
      });
      await context.sync();
    });
    expect(await runtime.run(settingsOf)).toEqual(await server.run(settingsOf));
    // The slide number field paints the slide's position, and the date field its date.
    expect(painted(editor, 1)).toEqual(expect.arrayContaining(['2', 'Board review', today()]));
    expect(painted(editor, 2)).toEqual(expect.arrayContaining(['3', 'Board review', today()]));
    expect(painted(editor, 0)).not.toContain('Board review');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.snapshot().canUndo).toBe(false);
    expect(await runtime.run(settingsOf)).toEqual(
      await (await open(withTitleSlide())).run(settingsOf)
    );
    server.dispose();
  });

  test('a write sends the lookups and one setHeaderFooter operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-header-footer.test.ts`) sends
    // these operations to two linked editors; this test keeps them equal to what the API sends.
    const editor = await live(twoSlides());
    const inner = createBrowserAutomationHost(editor);
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: false }).run(async (context) => {
      context.presentation.slides.getItemAt(1).headerFooter.footerText = 'Two';
      context.presentation.applyHeaderFooter({ slideNumberVisible: true });
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 1 } },
        { op: 'setHeaderFooter', slides: [{ result: 1 }], settings: { footerText: 'Two' } },
        { op: 'setHeaderFooter', settings: { slideNumberVisible: true } },
      ],
    ]);
  });
});

test('the example of the slides and shapes page runs as written', async () => {
  // docs/site/content/editor-api/slides-and-shapes.mdx, "Show slide numbers, the date, and a footer".
  async function numberSlides(runtime: PptxEditorRuntime, footer: string) {
    return runtime.run(async (context) => {
      context.presentation.applyHeaderFooter({
        slideNumberVisible: true,
        footerText: footer,
        dateVisible: true,
        dateFixed: false,
        hideOnTitleSlides: true,
      });
      const third = context.presentation.slides.getItemAt(2).headerFooter;
      third.footerText = `${footer} (appendix)`;
      third.load('slideNumberVisible,footerText,dateFixed');
      await context.sync();
      return third.toJSON();
    });
  }
  const runtime = await open(withTitleSlide());
  expect(await numberSlides(runtime, 'Board review')).toEqual({
    slideNumberVisible: true,
    footerText: 'Board review (appendix)',
    dateFixed: false,
  });
  expect((await runtime.run(settingsOf))[0]).toEqual(NONE);
  runtime.dispose();
});
