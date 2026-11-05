// The theme color grid the color menus show (audit Q18): the deck's own theme colors, through
// the slide's color map, with the lighter and darker variants of the color menus.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { variantsFor } from '../theme-palette.ts';

/** A one-slide deck whose theme is edited by `patch` (the builder writes the default theme). */
function deck(patch: (themeXml: string) => string = (x) => x, slideXml = ''): Uint8Array {
  const files = unzipSync(buildPptx({ slides: [{ shapes: slideXml }] }));
  const theme = 'ppt/theme/theme1.xml';
  files[theme] = strToU8(patch(strFromU8(files[theme]!)));
  return zipSync(files);
}

async function grid(bytes: Uint8Array) {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return { editor, rows: editor.snapshot().themeColors };
}

/** Channel distance between two `#RRGGBB` colors. */
const distance = (a: string, b: string): number =>
  Math.max(
    ...[1, 3, 5].map((i) =>
      Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16))
    )
  );

describe('theme color grid', () => {
  test('is empty before a document is open', () => {
    expect(createPptxEditor().snapshot().themeColors).toEqual([]);
  });

  test('lists the ten theme colors in menu order, then five variant rows', async () => {
    const { rows } = await grid(deck());
    expect(rows.length).toBe(6);
    expect(rows[0]!.map((s) => s.slot)).toEqual([
      'bg1',
      'tx1',
      'bg2',
      'tx2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
    ]);
    expect(rows[0]!.map((s) => s.color)).toEqual([
      '#FFFFFF',
      '#000000',
      '#E7E6E6',
      '#44546A',
      '#4472C4',
      '#ED7D31',
      '#A5A5A5',
      '#FFC000',
      '#5B9BD5',
      '#70AD47',
    ]);
    for (const row of rows) expect(row.length).toBe(10);
  });

  test('variants match the reference colors for the default theme', async () => {
    const { rows } = await grid(deck());
    const column = (i: number) => rows.map((r) => r[i]!);
    // Background 1 (white): Darker 5, 15, 25, 35, 50%.
    expect(column(0).map((s) => s.shift)).toEqual([0, -5, -15, -25, -35, -50]);
    // Accent 1: Lighter 80, 60, 40%, Darker 25, 50%.
    const accent = column(4);
    expect(accent.map((s) => s.shift)).toEqual([0, 80, 60, 40, -25, -50]);
    const defaultAccents = ['#4472C4', '#DAE3F3', '#B4C7E7', '#8FAADC', '#2F5597', '#203864'];
    accent.forEach((s, i) => expect(distance(s.color, defaultAccents[i]!)).toBeLessThanOrEqual(1));
    // Background 2 (light gray): Darker 10, 25, 50, 75, 90%.
    expect(column(2).map((s) => s.shift)).toEqual([0, -10, -25, -50, -75, -90]);
    expect(distance(column(2)[1]!.color, '#D0CECE')).toBeLessThanOrEqual(1);
  });

  test('uses the deck theme, not a fixed palette', async () => {
    const { rows } = await grid(deck((x) => x.replace('4472C4', '0F9ED5')));
    expect(rows[0]![4]!.color).toBe('#0F9ED5');
  });

  test('follows the slide color map override', async () => {
    const override =
      '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    const files = unzipSync(deck());
    const slide = 'ppt/slides/slide1.xml';
    files[slide] = strToU8(strFromU8(files[slide]!).replace('<a:masterClrMapping/>', override));
    const { rows } = await grid(zipSync(files));
    expect(rows[0]![0]!.color).toBe('#000000');
    expect(rows[0]![1]!.color).toBe('#FFFFFF');
  });

  test('is reference-stable across edits that keep the colors', async () => {
    const { editor, rows } = await grid(deck());
    editor.exec({ type: 'addSlide' });
    expect(editor.snapshot().themeColors).toBe(rows);
  });

  test('variants by lightness', () => {
    expect(variantsFor({ r: 0, g: 0, b: 0, a: 1 })).toEqual([50, 35, 25, 15, 5]);
    expect(variantsFor({ r: 20, g: 30, b: 40, a: 1 })).toEqual([90, 75, 50, 25, 10]);
    expect(variantsFor({ r: 240, g: 240, b: 240, a: 1 })).toEqual([-10, -25, -50, -75, -90]);
  });
});

const IN = 914_400;

describe('theme color picks', () => {
  const shapes = shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 2 * IN,
    cy: IN,
    fill: solid('FFC000'),
    text: para(run('Hi')),
  });
  const slideXml = (e: PptxEditorInstance): string =>
    serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);
  const selectShape = (e: PptxEditorInstance): void => {
    const id = e.session!.layout(0)!.items[0]!.element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
  };

  test('a theme swatch writes a:schemeClr with its lumMod and lumOff, not a hex color', async () => {
    const { editor, rows } = await grid(deck(undefined, shapes));
    selectShape(editor);
    const swatch = rows[2]![4]!; // Accent 1, Lighter 60%
    expect(swatch).toMatchObject({ slot: 'accent1', shift: 60 });
    expect(editor.exec({ type: 'setShapeFill', color: { theme: 'accent1', shift: 60 } }).ok).toBe(
      true
    );
    expect(slideXml(editor)).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:solidFill>'
    );
    // The shape renders the color the swatch shows.
    expect(editor.snapshot().geometry?.fill).toBe(swatch.color);

    expect(editor.exec({ type: 'setLineColor', color: { theme: 'tx1', shift: -25 } }).ok).toBe(
      true
    );
    expect(slideXml(editor)).toContain(
      '<a:ln><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln>'
    );
    expect(editor.exec({ type: 'setTextColor', color: { theme: 'accent2' } }).ok).toBe(true);
    expect(slideXml(editor)).toContain('<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>');
    expect(
      editor.exec({
        type: 'setSlideBackground',
        background: { kind: 'solid', color: { theme: 'bg2' } },
      }).ok
    ).toBe(true);
    expect(slideXml(editor)).toContain('<a:schemeClr val="bg2"/>');
    expect(readPptx(editor.save()!).ok).toBe(true);
  });

  test('a standard color still writes a:srgbClr', async () => {
    const { editor } = await grid(deck(undefined, shapes));
    selectShape(editor);
    expect(editor.exec({ type: 'setShapeFill', color: '#C00000' }).ok).toBe(true);
    expect(slideXml(editor)).toContain('<a:srgbClr val="C00000"/>');
  });

  test('refuses unknown theme names and out-of-range variants', async () => {
    const { editor } = await grid(deck(undefined, shapes));
    selectShape(editor);
    for (const color of [
      { theme: 'phClr' },
      { theme: 'accent1', shift: 100 },
      { theme: 'accent1', shift: 2.5 },
      { theme: 'accent1" injected="1' },
    ])
      expect(editor.can({ type: 'setShapeFill', color: color as never }).ok).toBe(false);
  });
});
