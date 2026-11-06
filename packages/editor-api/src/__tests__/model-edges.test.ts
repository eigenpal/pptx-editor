/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Edge cases of the object model: links found by a lookup, empty ranges, odd tag and property
// values, and the reasons a refusal gives.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { readZip, writeZip } from '@pptx-editor.dev/core/store';
import { isPptxEditorError, type RequestContext, type Slide } from '../index.ts';
import { nearest } from '../runtime/errors.ts';
import { deck, IN, open, rejection, reopen } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);
const inner = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.debugInfo.innerError?.code : undefined;

/** One box whose text has three runs: `Hel` (bold), `lo ` (italic), and `World`. */
const runs = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'T',
            x: 0,
            y: 0,
            cx: 3 * IN,
            cy: IN,
            text: para(run('Hel', 'b="1"') + run('lo ', 'i="1"') + run('World')),
          }) + shape({ id: 3, name: 'Plain', x: 4 * IN, y: 0, cx: IN, cy: IN }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

const slide = (context: RequestContext) => context.presentation.slides.getItemAt(0);

describe('hyperlinks a lookup finds', () => {
  test('getItemAt gives the linked range and shape in the same sync', async () => {
    const runtime = await open(runs());
    await runtime.run(async (context) => {
      slide(context)
        .shapes.getItem('2')
        .textFrame.textRange.getSubstring(1, 7)
        .setHyperlink({ address: 'https://example.com/t', screenTip: 'T' });
      slide(context).shapes.getItem('3').setHyperlink({ address: 'https://example.com/s' });
      await context.sync();
      const links = slide(context).hyperlinks;
      const range = links.getItemAt(0).getLinkedTextRangeOrNullObject().load('text,start,length');
      const noShape = links.getItemAt(0).getLinkedShapeOrNullObject().load('id');
      const linkedShape = links.getItemAt(1).getLinkedShapeOrNullObject().load('id');
      const noRange = links.getItemAt(1).getLinkedTextRangeOrNullObject().load('text');
      await context.sync();
      expect([range.isNullObject, range.text, range.start, range.length]).toEqual([
        false,
        'ello Wo',
        1,
        7,
      ]);
      expect(noShape.isNullObject).toBe(true);
      expect([linkedShape.isNullObject, linkedShape.id]).toEqual([false, '3']);
      expect(noRange.isNullObject).toBe(true);
    });
  });

  test('a write to one property of a found link keeps the other', async () => {
    const runtime = await open(runs());
    await runtime.run(async (context) => {
      slide(context)
        .shapes.getItem('2')
        .textFrame.textRange.getSubstring(0, 5)
        .setHyperlink({ address: 'https://example.com/a', screenTip: 'Old' });
      await context.sync();
      // No sync between the lookup and the write.
      slide(context).hyperlinks.getItemAt(0).screenTip = 'New';
      await context.sync();
      const links = slide(context).hyperlinks.load('items');
      await context.sync();
      expect(links.items.map((l) => [l.address, l.screenTip])).toEqual([
        ['https://example.com/a', 'New'],
      ]);
    });
  });

  test('a link on an empty range refuses and writes nothing', async () => {
    const runtime = await open(runs());
    const error = await rejection(
      runtime.run(async (context) => {
        slide(context)
          .shapes.getItem('2')
          .textFrame.textRange.getSubstring(2, 0)
          .setHyperlink({ address: 'https://example.com' });
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
    await runtime.run(async (context) => {
      const count = slide(context).hyperlinks.getCount();
      await context.sync();
      expect(count.value).toBe(0);
    });
  });
});

describe('characters XML cannot carry', () => {
  test('a tag key or value, or a property, refuses as an argument', async () => {
    const runtime = await open(runs());
    const attempts: ((context: RequestContext) => void)[] = [
      (context) => context.presentation.tags.add('key\u0000', 'v'),
      (context) => slide(context).tags.add('k', 'v￿'),
      (context) => {
        context.presentation.properties.title = 'a\u0001b';
      },
      (context) => {
        slide(context).shapes.getItem('2').altTextTitle = 'a\u0000';
      },
      (context) => {
        slide(context).shapes.getItem('2').name = 'a\u0007b';
      },
      (context) => {
        slide(context).shapes.getItem('3').setHyperlink({
          address: 'https://example.com',
          screenTip: '\uFFFE',
        });
      },
    ];
    for (const attempt of attempts) {
      const error = await rejection(
        runtime.run(async (context) => {
          attempt(context);
          await context.sync();
        })
      );
      expect(code(error)).toBe('InvalidArgument');
    }
    // Odd but valid keys round-trip, stored in upper case.
    await runtime.run(async (context) => {
      context.presentation.tags.add('__proto__', '<&>"');
      context.presentation.tags.add('ä b', 'x');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const tags = context.presentation.tags.load('items');
      await context.sync();
      expect(tags.items.map((t) => [t.key, t.value])).toEqual([
        ['__PROTO__', '<&>"'],
        ['Ä B', 'x'],
      ]);
    });
  });
});

describe('refusal reasons', () => {
  test('groups need two different shapes of one parent on a slide', async () => {
    const runtime = await open(runs());
    const group = (values: string[]) =>
      rejection(
        runtime.run(async (context) => {
          slide(context).shapes.addGroup(values);
          await context.sync();
        })
      );
    expect(inner(await group(['2']))).toBe('edit.group-count');
    expect(inner(await group(['2', '2']))).toBe('edit.group-count');
    const master = await rejection(
      runtime.run(async (context) => {
        const shapes = context.presentation.slideMasters.getItemAt(0).shapes;
        shapes.addGroup([shapes.addTextBox('a'), shapes.addTextBox('b')]);
        await context.sync();
      })
    );
    expect([code(master), inner(master)]).toEqual(['NotSupported', 'edit.group-off-slide']);
  });

  test('a placeholder cannot be grouped, and the refusal says why', async () => {
    const runtime = await open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              shape({
                id: 2,
                name: 'Title',
                ph: '<p:ph type="title"/>',
                x: 0,
                y: 0,
                cx: IN,
                cy: IN,
              }) + shape({ id: 3, name: 'Box', x: 2 * IN, y: 0, cx: IN, cy: IN }),
          },
        ],
      })
    );
    const error = await rejection(
      runtime.run(async (context) => {
        slide(context).shapes.addGroup(['2', '3']);
        await context.sync();
      })
    );
    expect([code(error), inner(error)]).toEqual(['NotSupported', 'edit.group-placeholder']);
  });

  test('the last slide cannot be deleted', async () => {
    const runtime = await open(runs());
    const error = await rejection(
      runtime.run(async (context) => {
        const slides = context.presentation.slides.load('items');
        await context.sync();
        for (const item of slides.items) item.delete();
        await context.sync();
      })
    );
    expect([code(error), inner(error)]).toEqual(['NotSupported', 'edit.last-slide']);
  });

  test('a group has no outline of its own', async () => {
    const runtime = await open(runs());
    const error = await rejection(
      runtime.run(async (context) => {
        const group = slide(context).shapes.addGroup(['2', '3']);
        group.lineFormat.color = '#FF0000';
        await context.sync();
      })
    );
    expect([code(error), inner(error)]).toEqual(['NotSupported', 'edit.group-line']);
    // A group fill is valid (CT_GroupShapeProperties holds EG_FillProperties).
    await runtime.run(async (context) => {
      slide(context).shapes.addGroup(['2', '3']).fill.setSolidColor('#00FF00');
      await context.sync();
    });
    await reopen(runtime);
  });
});

describe('slide ids', () => {
  test('a deleted slide id does not come back, so an old proxy cannot reach a new slide', async () => {
    const runtime = await open(runs());
    let old!: Slide;
    await runtime.run(async (context) => {
      old = context.presentation.slides.getItemAt(0).load('id');
      context.trackedObjects.add(old);
      await context.sync();
    });
    const ids = await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).delete();
      await context.sync();
      context.presentation.slides.add();
      const slides = context.presentation.slides.load('items/id');
      await context.sync();
      return slides.items.map((s) => s.id);
    });
    expect(ids).not.toContain(old.id);
    const error = await rejection(
      runtime.run(old, async (context) => {
        old.load('index');
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidObjectPath');
  });
});

describe('document properties', () => {
  test('creationDate is always a Date, invalid when the file has none', async () => {
    const runtime = await open();
    const dated = await runtime.run(async (context) => {
      const props = context.presentation.properties.load('creationDate');
      await context.sync();
      return props.creationDate;
    });
    expect(dated).toBeInstanceOf(Date);
    expect(dated.toISOString()).toBe('2026-01-01T00:00:00.000Z');

    // The same deck with no dcterms:created in docProps/core.xml.
    const zip = readZip(deck());
    if (!zip.ok) throw new Error('fixture');
    const entries = new Map(zip.entries);
    const name = [...entries.keys()].find((key) => key.endsWith('docProps/core.xml'))!;
    const core = new TextDecoder().decode(entries.get(name));
    const stripped = core.replace(/<dcterms:created[^>]*>[^<]*<\/dcterms:created>/, '');
    expect(stripped).not.toBe(core);
    entries.set(name, new TextEncoder().encode(stripped));
    const undated = await open(writeZip(entries));
    const missing = await undated.run(async (context) => {
      const props = context.presentation.properties.load('creationDate');
      await context.sync();
      return props.creationDate;
    });
    expect(missing).toBeInstanceOf(Date);
    expect(Number.isNaN(missing.getTime())).toBe(true);
  });
});

describe('argument refusals name what is accepted', () => {
  /** The message of the refusal that `edit` makes, after checking its code and diagnostic. */
  async function refused(edit: (context: RequestContext) => void): Promise<string> {
    const runtime = await open(runs());
    const error = await rejection(
      runtime.run(async (context) => {
        edit(context);
        await context.sync();
      })
    );
    if (!isPptxEditorError(error)) throw error;
    expect([error.code, inner(error)]).toEqual(['InvalidArgument', 'argument.not-accepted']);
    expect(error.message).toContain(error.debugInfo.innerError?.message ?? '');
    return error.message;
  }

  test('an enum value lists the values of the enum', async () => {
    expect(
      await refused((context) => {
        context.presentation.insertSlidesFromBase64('UEsDBA==', {
          formatting: 'keepSource' as never,
        });
      })
    ).toContain("Use one of 'KeepSourceFormatting' or 'UseDestinationTheme'.");
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').textFrame.autoSizeSetting = 'Shrink' as never;
      })
    ).toContain(
      "Use one of 'AutoSizeNone', 'AutoSizeTextToFitShape', or 'AutoSizeShapeToFitText'."
    );
    // A long enum names its size and a few values.
    expect(
      await refused((context) => {
        slide(context).shapes.addGeometricShape('Blob' as never);
      })
    ).toMatch(/Use one of the \d+ values of the enum, such as /);
  });

  test('an enum typo suggests the nearest value, except where the host may refuse it', async () => {
    expect(
      await refused((context) => {
        slide(context).shapes.addGeometricShape('Rectange' as never);
      })
    ).toContain("Did you mean 'Rectangle'?");
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').textFrame.autoSizeSetting = 'AutoSizeNon' as never;
      })
    ).toContain("Did you mean 'AutoSizeNone'?");
    // A table style that is not defined in the presentation refuses at the sync, so a typo of
    // one gets no suggestion.
    const style = await refused((context) => {
      const table = slide(context).shapes.addTable(1, 1);
      table.getTable().styleSettings.style = 'MediumStyle2Acent1' as never;
    });
    expect(style).not.toContain('Did you mean');
  });

  test('a value far from every typo suggests the value that holds its words', async () => {
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').textFrame.textRange.paragraphFormat.bulletFormat.style =
          'ArabicPeriod' as never;
      })
    ).toContain("Did you mean 'ArabicNumeralPeriod'?");
    expect(
      await refused((context) => {
        slide(context).shapes.addGeometricShape('Rect' as never);
      })
    ).toContain("Did you mean 'Rectangle'?");
    expect(
      await refused((context) => {
        context.presentation.insertSlidesFromBase64('UEs=', { formatting: 'KeepSource' as never });
      })
    ).toContain("Did you mean 'KeepSourceFormatting'?");
    expect(
      await refused((context) => {
        slide(context).shapes.addTable(1, 1, { colums: [] } as never);
      })
    ).toContain("Did you mean 'columns'?");
  });

  test('a suggestion is always one of the accepted values', () => {
    const accepted = ['ArabicNumeralPeriod', 'RomanUppercasePeriod', 'Rectangle', 'AutoSizeNone'];
    const values = [
      'ArabicPeriod',
      'Period',
      'rect',
      'auto',
      'zzz',
      'Arabic Numeral',
      'x'.repeat(80),
    ];
    for (let seed = 0; seed < 200; seed += 1)
      values.push(
        Math.random()
          .toString(36)
          .slice(2, 2 + (seed % 12))
      );
    for (const value of values) {
      const guess = nearest(value, accepted);
      if (guess !== undefined) expect(accepted).toContain(guess);
    }
    expect(nearest('zzz', accepted)).toBeUndefined();
    expect(nearest('somethingElse', accepted)).toBeUndefined();
  });

  test('a link address lists the allowed schemes', async () => {
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').setHyperlink({ address: 'javascript:alert(1)' });
      })
    ).toContain('Use an http, https, mailto, tel, or ftp address');
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').textFrame.textRange.setHyperlink({ address: 'data:x' });
      })
    ).toContain('Use an http, https, mailto, tel, or ftp address');
  });

  test('a custom property value names the accepted types', async () => {
    expect(
      await refused((context) => {
        context.presentation.properties.customProperties.add('X', null as never);
      })
    ).toContain('Use a string, a finite number, a boolean, or a valid Date.');
  });

  test('a scalar of the wrong type names the type', async () => {
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').visible = 'yes' as never;
      })
    ).toContain('Use true or false.');
    expect(
      await refused((context) => {
        slide(context).shapes.getItem('2').left = Number.NaN;
      })
    ).toContain('Use a finite number.');
  });
});
