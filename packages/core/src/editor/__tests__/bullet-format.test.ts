// The Bullets and Numbering galleries (`setBulletFormat`): the gallery characters, fonts, and
// schemes, written in schema order, with margins where nothing is inherited, as one undo step.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  readPptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type TextMeasurer } from '../../layout/index.ts';
import {
  colorChange,
  colorChoiceOf,
  colorPickOf,
  colorPicks,
  createPptxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';

const IN = 914_400;
const BODY = '<p:ph type="body" idx="1"/>';

function deck(text: string, ph?: string): Uint8Array {
  return buildPptx({
    slides: [
      { shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text, ph }) },
    ],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const pPr = (e: PptxEditorInstance, i = 0): Element | undefined =>
  child(elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p')[i]!, 'a:pPr');
const names = (el: Element | undefined): string[] => elements(el).map((c) => c.local);
const select = (e: PptxEditorInstance, from = 0, to = from) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: from, offset: 0 },
    focus: { paragraph: to, offset: 1 },
  });
const marker = (e: PptxEditorInstance, line = 0): string | undefined => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
  const runs = item.kind === 'shape' ? item.text!.lines[line]!.runs : [];
  return runs.find((r) => r.kind === 'bullet')?.text;
};

describe('bullets gallery', () => {
  test('a size alone that a list already has is no refusal', async () => {
    const e = await open(
      deck(
        para(
          run('one'),
          '<a:pPr marL="342900" indent="-342900"><a:buSzPct val="100000"/><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>'
        )
      )
    );
    select(e, 0, 3);
    expect(e.exec({ type: 'setBulletFormat', format: { sizePercent: 100 } }).ok).toBe(true);
  });

  test('a size or a color alone on text with no list is refused, with no undo step', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e, 0, 3);
    const before = canonicalPackageFingerprint(e.session!.store.package);
    const command = { type: 'setBulletFormat', format: { sizePercent: 150 } } as const;
    expect(e.can(command)).toEqual({ ok: false, reason: 'disabled.noList' });
    expect(e.exec(command)).toEqual({ ok: false, reason: 'disabled.noList' });
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(before);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('a Wingdings bullet with its font attributes and a hanging indent', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e);
    expect(runToolbarCommand(e, 'list.bulletStyle', 'filledSquare').ok).toBe(true);
    const p = pPr(e)!;
    expect([attr(p, 'marL'), attr(p, 'indent')]).toEqual(['342900', '-342900']);
    expect(names(p)).toEqual(['buFont', 'buChar']);
    const font = find(p, 'a:buFont')!;
    expect([attr(font, 'typeface'), attr(font, 'charset'), attr(font, 'pitchFamily')]).toEqual([
      'Wingdings',
      '2',
      '2',
    ]);
    expect(attr(find(p, 'a:buChar'), 'char')).toBe('§');
    expect(pPr(e, 1)).toBeUndefined();
    expect(marker(e)).toBe('§');
    expect(toolbarCommandState(e, 'list.bulletStyle').value).toBe('filledSquare');
    expect(toolbarCommandState(e, 'list.bullet').active).toBe(true);
    e.exec({ type: 'undo' });
    expect(pPr(e)).toBeUndefined();
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('on a placeholder, the gallery overrides the inherited bullet and keeps its margins', async () => {
    const e = await open(deck(para(run('item')), BODY));
    select(e);
    expect(e.snapshot().format?.bulletFormat?.kind).toBe('bullet');
    expect(runToolbarCommand(e, 'list.bulletStyle', 'checkmark').ok).toBe(true);
    const p = pPr(e)!;
    expect(attr(p, 'marL')).toBeUndefined();
    expect(names(p)).toEqual(['buFont', 'buChar']);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ kind: 'bullet', char: 'ü' });
    // None turns the list off.
    expect(runToolbarCommand(e, 'list.bulletStyle', 'none').ok).toBe(true);
    expect(e.snapshot().format?.bulletFormat).toBeNull();
    expect(toolbarCommandState(e, 'list.bulletStyle').value).toBe('none');
  });

  test('size and color write a:buClr and a:buSzPct in schema order; null follows the text', async () => {
    const e = await open(
      deck(para(run('x'), '<a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:tabLst/></a:pPr>'))
    );
    select(e);
    expect(
      e.exec({
        type: 'setBulletFormat',
        format: {
          kind: 'bullet',
          char: '–',
          font: 'Arial',
          sizePercent: 150,
          color: { theme: 'accent2' },
        },
      }).ok
    ).toBe(true);
    expect(names(pPr(e))).toEqual(['lnSpc', 'buClr', 'buSzPct', 'buFont', 'buChar', 'tabLst']);
    expect(attr(find(pPr(e)!, 'a:buSzPct'), 'val')).toBe('150000');
    // A gallery pick keeps the size and color.
    runToolbarCommand(e, 'list.bulletStyle', 'star');
    expect(names(pPr(e))).toEqual(['lnSpc', 'buClr', 'buSzPct', 'buFont', 'buChar', 'tabLst']);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ char: 'v', sizePercent: 150 });
    e.exec({
      type: 'setBulletFormat',
      format: { kind: 'bullet', char: 'v', sizePercent: null, color: null },
    });
    // Null follows the text: the color, the size, and (a character with no font) the font.
    expect(names(pPr(e))).toEqual(['lnSpc', 'buClrTx', 'buSzTx', 'buFontTx', 'buChar', 'tabLst']);
  });

  test('a restyle keeps the margins of a list the paragraph already shows', async () => {
    const own =
      '<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
    const e = await open(deck(para(run('x'), own)));
    select(e);
    expect(runToolbarCommand(e, 'list.bulletStyle', 'star').ok).toBe(true);
    expect([attr(pPr(e), 'marL'), attr(pPr(e), 'indent')]).toEqual(['285750', '-285750']);
    e.exec({ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 120 } });
    expect([attr(pPr(e), 'marL'), attr(pPr(e), 'indent')]).toEqual(['285750', '-285750']);
    expect(attr(find(pPr(e)!, 'a:buChar'), 'char')).toBe('v');
  });

  test('None puts a:buNone after the spacing, color, and size, in schema order', async () => {
    const e = await open(deck(para(run('x')), BODY));
    select(e);
    e.exec({ type: 'setParagraphFormat', spaceBeforePt: 6 });
    e.exec({
      type: 'setBulletFormat',
      format: { kind: 'bullet', color: { theme: 'accent2' }, sizePercent: 120 },
    });
    expect(runToolbarCommand(e, 'list.bulletStyle', 'none').ok).toBe(true);
    expect(names(pPr(e))).toEqual(['spcBef', 'buClr', 'buSzPct', 'buNone']);
  });

  test('null follows the text also over a list style, and a character alone follows its font', async () => {
    const lst =
      '<a:lvl1pPr marL="342900" indent="-342900"><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="150000"/><a:buFont typeface="Wingdings"/><a:buChar char="v"/></a:lvl1pPr>';
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: 3 * IN,
              lstStyle: lst,
              text: para(run('x')),
            }),
          },
        ],
      })
    );
    select(e);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ color: '#FF0000', sizePercent: 150 });
    e.exec({
      type: 'setBulletFormat',
      format: { kind: 'bullet', char: '–', color: null, sizePercent: null },
    });
    const read = e.snapshot().format?.bulletFormat;
    expect(read).toMatchObject({ kind: 'bullet', char: '–', sizePercent: 100 });
    expect(read && 'color' in read ? read.color : undefined).toBeUndefined();
    expect(read && 'font' in read ? read.font : undefined).not.toBe('Wingdings');
  });

  test('a size or color change keeps each paragraph its own bullet, and the read is absent when they differ', async () => {
    const e = await open(
      deck(
        para(
          run('a'),
          '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="v"/></a:pPr>'
        ) +
          para(
            run('b'),
            '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:pPr>'
          )
      )
    );
    select(e, 0, 1);
    expect(e.snapshot().format?.bulletFormat).toBeUndefined();
    expect(toolbarCommandState(e, 'list.bulletStyle').value).toBeUndefined();
    e.exec({ type: 'setBulletFormat', format: { kind: 'bullet', sizePercent: 150 } });
    expect([
      attr(find(pPr(e, 0)!, 'a:buChar'), 'char'),
      attr(find(pPr(e, 1)!, 'a:buChar'), 'char'),
    ]).toEqual(['v', 'ü']);
    expect(attr(find(pPr(e, 1)!, 'a:buSzPct'), 'val')).toBe('150000');
  });

  test('a point-sized bullet reads as a percent of the text size', async () => {
    const e = await open(
      deck(
        para(
          run('a', 'sz="2000"'),
          '<a:pPr marL="342900" indent="-342900"><a:buSzPts val="3000"/><a:buChar char="•"/></a:pPr>'
        )
      )
    );
    select(e);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ kind: 'bullet', sizePercent: 150 });
  });

  test('a start number alone keeps the scheme, also one the editor does not set', async () => {
    const e = await open(
      deck(
        para(
          run('a'),
          '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="circleNumDbPlain"/></a:pPr>'
        )
      )
    );
    select(e);
    expect(e.exec({ type: 'setBulletFormat', format: { kind: 'number', startAt: 3 } }).ok).toBe(
      true
    );
    const num = find(pPr(e)!, 'a:buAutoNum')!;
    expect([attr(num, 'type'), attr(num, 'startAt')]).toEqual(['circleNumDbPlain', '3']);
  });
});

describe('numbering gallery', () => {
  test('writes the scheme with startAt, and numbers the paragraphs', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e, 0, 1);
    expect(runToolbarCommand(e, 'list.numberStyle', 'alphaLcParenR').ok).toBe(true);
    expect(attr(find(pPr(e)!, 'a:buAutoNum'), 'type')).toBe('alphaLcParenR');
    expect([marker(e, 0), marker(e, 1)]).toEqual(['a)', 'b)']);
    expect(toolbarCommandState(e, 'list.numbered').active).toBe(true);
    expect(toolbarCommandState(e, 'list.numberStyle').value).toBe('alphaLcParenR');
    e.exec({
      type: 'setBulletFormat',
      format: { kind: 'number', scheme: 'romanUcPeriod', startAt: 4 },
    });
    expect(attr(find(pPr(e)!, 'a:buAutoNum'), 'startAt')).toBe('4');
    expect([marker(e, 0), marker(e, 1)]).toEqual(['IV.', 'V.']);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({
      scheme: 'romanUcPeriod',
      startAt: 4,
    });
    // Another scheme with no start keeps the paragraphs' start, as size and color do.
    expect(runToolbarCommand(e, 'list.numberStyle', 'arabicPeriod').ok).toBe(true);
    expect([marker(e, 0), marker(e, 1)]).toEqual(['4.', '5.']);
  });

  test('a start number on an inherited scheme writes the scheme on the paragraph', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: 3 * IN,
              lstStyle:
                '<a:lvl1pPr marL="342900" indent="-342900"><a:buAutoNum type="romanLcParenR"/></a:lvl1pPr>',
              text: para(run('one')),
            }),
          },
        ],
      })
    );
    select(e);
    expect(e.exec({ type: 'setBulletFormat', format: { kind: 'number', startAt: 3 } }).ok).toBe(
      true
    );
    expect(names(pPr(e))).toEqual(['buAutoNum']);
    expect(marker(e)).toBe('iii)');
  });
});

describe('setBulletFormat', () => {
  test('size or color with no kind keeps each paragraph its own bullet or number', async () => {
    const e = await open(
      deck(
        para(
          run('dot'),
          '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>'
        ) +
          para(
            run('roman'),
            '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr>'
          ) +
          para(run('plain'))
      )
    );
    select(e, 0, 2);
    expect(e.snapshot().format?.bulletFormat).toBeUndefined();
    expect(
      e.exec({ type: 'setBulletFormat', format: { sizePercent: 150, color: '#C00000' } }).ok
    ).toBe(true);
    expect(names(pPr(e, 0))).toEqual(['buClr', 'buSzPct', 'buFont', 'buChar']);
    expect(names(pPr(e, 1))).toEqual(['buClr', 'buSzPct', 'buFont', 'buAutoNum']);
    expect(attr(find(pPr(e, 1)!, 'a:buAutoNum'), 'startAt')).toBe('3');
    expect(pPr(e, 2)).toBeUndefined();
    expect([marker(e, 0), marker(e, 1), marker(e, 2)]).toEqual(['•', 'III.', undefined]);
    e.exec({ type: 'undo' });
    expect(names(pPr(e, 0))).toEqual(['buFont', 'buChar']);
  });

  test('a start number over a:buNone writes the inherited scheme with it', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 6 * IN,
              cy: 3 * IN,
              lstStyle:
                '<a:lvl1pPr marL="342900" indent="-342900"><a:buAutoNum type="romanLcParenR"/></a:lvl1pPr>',
              text: para(run('one'), '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>'),
            }),
          },
        ],
      })
    );
    select(e);
    expect(e.exec({ type: 'setBulletFormat', format: { kind: 'number', startAt: 3 } }).ok).toBe(
      true
    );
    expect(marker(e)).toBe('iii)');
    expect(e.snapshot().format?.bulletFormat).toMatchObject({
      scheme: 'romanLcParenR',
      startAt: 3,
    });
  });

  test('the read covers every selected shape, in any order', async () => {
    const two = (id: number, name: string, pPrXml: string, y: number) =>
      shape({ id, name, x: IN, y, cx: 4 * IN, cy: IN, text: para(run(name), pPrXml) });
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes:
              two(
                2,
                'A',
                '<a:pPr marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr>',
                IN
              ) +
              two(
                3,
                'B',
                '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="romanUcPeriod"/></a:pPr>',
                3 * IN
              ),
          },
        ],
      })
    );
    const id = (name: string) =>
      elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === name)!
        .id;
    for (const order of [
      ['A', 'B'],
      ['B', 'A'],
    ]) {
      e.setSelection({ kind: 'shapes', nodeIds: order.map(id) });
      expect(e.snapshot().format?.bulletFormat).toBeUndefined();
      expect(toolbarCommandState(e, 'list.bulletStyle').value).toBeUndefined();
    }
  });

  test('refuses styles the schema or layout does not support', async () => {
    const e = await open(deck(para(run('x'))));
    select(e);
    for (const style of [
      { kind: 'bullet', char: 'ab' },
      { kind: 'bullet', char: '' },
      { kind: 'bullet', char: '\u0007' },
      { kind: 'bullet', char: '•', font: 'x;y{}' },
      { kind: 'number', scheme: 'circleNumDbPlain' },
      { kind: 'number', scheme: 'romanUcPlain' },
      { kind: 'bullet', char: '\uD800' },
      { kind: 'number', scheme: 'arabicPeriod', startAt: 0 },
      { kind: 'number', scheme: 'arabicPeriod', startAt: 40_000 },
      { kind: 'bullet', char: '•', sizePercent: 10 },
      { kind: 'bullet', char: '•', color: 'red' },
      { kind: 'picture' },
    ]) {
      const r = e.exec({ type: 'setBulletFormat', format: style as never });
      expect({ style, ok: r.ok }).toEqual({ style, ok: false });
    }
    // A field the kind does not take, or a format with no kind and nothing to change, is refused
    // with the field's name, never dropped.
    for (const [style, field] of [
      [{ kind: 'bullet', scheme: 'romanUcPeriod' }, 'scheme'],
      [{ kind: 'number', char: 'x' }, 'char'],
      [{ startAt: 3 }, 'startAt'],
      [{ char: 'x', sizePercent: 120 }, 'char'],
      [{}, 'kind'],
      [{ kind: 'bullet', glyph: '•' }, 'glyph'],
      [{ kind: 'none', sizePercent: 120 }, 'kind'],
      [{ kind: 'constructor' }, 'kind'],
      [{ kind: 'toString', sizePercent: 120 }, 'kind'],
    ] as const) {
      const r = e.exec({ type: 'setBulletFormat', format: style as never });
      expect(r.ok ? null : r.detail).toBe(field);
    }
    expect(e.snapshot().canUndo).toBe(false);
    // An emoji is one character.
    expect(e.exec({ type: 'setBulletFormat', format: { kind: 'bullet', char: '👉' } }).ok).toBe(
      true
    );
  });

  test('the style survives save and reopen', async () => {
    const e = await open(deck(para(run('saved'))));
    select(e);
    e.exec({
      type: 'setBulletFormat',
      format: { kind: 'number', scheme: 'arabicParenBoth', startAt: 7, sizePercent: 80 },
    });
    const again = await open(e.save()!);
    select(again);
    expect(again.snapshot().format?.bulletFormat).toEqual({
      kind: 'number',
      scheme: 'arabicParenBoth',
      startAt: 7,
      sizePercent: 80,
    });
  });
});

describe('symbol font bullets', () => {
  test('a Wingdings bullet draws its Unicode twin when the font is not loaded, and keeps its byte', () => {
    const bytes = deck(
      para(
        run('done'),
        '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:pPr>'
      )
    );
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('open failed');
    const base = createFixedMeasurer();
    const bullet = (measurer: TextMeasurer) => {
      const item = layoutSlide(resolveSlide(r.pkg, 0)!, measurer).items.find(
        (x) => x.kind === 'shape' && x.text
      )!;
      const runs = item.kind === 'shape' ? item.text!.lines[0]!.runs : [];
      return runs.find((x) => x.kind === 'bullet')?.text;
    };
    // Loaded: the font's own byte. Not loaded: the check mark.
    expect(bullet(base)).toBe('ü');
    const missing: TextMeasurer = Object.assign(Object.create(base) as TextMeasurer, {
      resolve: (q: Parameters<TextMeasurer['resolve']>[0]) => ({
        ...base.resolve(q),
        substituted: /wingdings/i.test(q.family),
      }),
    });
    expect(bullet(missing)).toBe('✔');
  });
});

describe('color field picks', () => {
  test('a tint, a shade, and a standard color round-trip through a pick', () => {
    expect(colorPickOf({ theme: 'accent1', shift: 40 })).toBe('accent1:40');
    expect(colorChoiceOf('accent1:-25')).toEqual({ theme: 'accent1', shift: -25 });
    expect(colorPickOf('#c00000')).toBe('#C00000');
    expect(colorChange('auto', '#C00000')).toBe('#C00000');
    expect(colorChange('keep', 'accent2')).toEqual({ theme: 'accent2' });
    expect(colorPicks('auto', 'accent1:40').at(-1)).toBe('accent1:40');
    expect(colorPicks('auto', 'accent1')).not.toContain('accent1:40');
  });
});
