// The format painter (format-painter.ts): copy text and shape formatting, paint it as one undo
// step, the once and sticky modes of the button, and typed refusals.

import { afterEach, describe, expect, setSystemTime, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const STYLE =
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
const EFFECTS =
  '<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text:
              para(
                run('alpha beta gamma'),
                '<a:pPr algn="ctr"><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:pPr>'
              ) + para(run('delta  epsilon')),
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          }) +
          shape({
            id: 3,
            x: IN,
            y: 3 * IN,
            cx: 2 * IN,
            cy: IN,
            prst: 'ellipse',
            fill: solid('C00000'),
            line: '<a:ln w="38100"><a:solidFill><a:srgbClr val="002060"/></a:solidFill></a:ln>',
            effects: EFFECTS,
            style: STYLE,
          }) +
          shape({
            id: 4,
            x: 4 * IN,
            y: 3 * IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('92D050'),
            text: para(run('target', 'sz="1400"')),
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

const ids = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);
const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
/** The XML of the n-th shape on the slide (0-based). */
const shapeXml = (e: PptxEditorInstance, n: number): string =>
  slideXml(e).split('<p:sp>')[n + 1]!.split('</p:sp>')[0]!;
const text = (
  e: PptxEditorInstance,
  id: string,
  paragraph: number,
  from: number,
  to = from
): void =>
  e.setSelection({
    kind: 'text',
    nodeId: id,
    anchor: { paragraph, offset: from },
    focus: { paragraph, offset: to },
  });
/** The `a:r` of the run holding `word`. */
const runOf = (xml: string, word: string): string =>
  xml.split('<a:r>').find((r) => r.includes(`<a:t>${word}`)) ?? '';

afterEach(() => setSystemTime());

describe('format painter: text', () => {
  test('copies the run at the selection start and paints a range and a word as one undo step', async () => {
    const e = await open();
    const [box] = ids(e);
    text(e, box, 0, 0, 5);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(e.exec({ type: 'setTextColor', color: '#C00000' }).ok).toBe(true);

    // Arm from a caret inside the bold word, then paint "gamma" (a range).
    text(e, box, 0, 3);
    expect(e.exec({ type: 'toggleFormatPainter' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('once');
    text(e, box, 0, 11, 16);
    const before = slideXml(e);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    const gamma = runOf(shapeXml(e, 0), 'gamma');
    expect(gamma).toContain('b="1"');
    expect(gamma).toContain('<a:srgbClr val="C00000"/>');
    // The language stays the target's; "beta" keeps its own formatting.
    expect(gamma).toContain('lang="en-US"');
    expect(runOf(shapeXml(e, 0), 'beta')).not.toContain('b="1"');
    // One paint, then the painter is off.
    expect(e.snapshot().formatPainter).toBe('off');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(e)).toBe(before);

    // Ctrl+Shift+C / Ctrl+Shift+V: a caret paints the word under it.
    text(e, box, 0, 2);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('off');
    text(e, box, 1, 9);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(runOf(shapeXml(e, 0), 'epsilon')).toContain('b="1"');
    expect(runOf(shapeXml(e, 0), 'delta')).not.toContain('b="1"');
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  // The painter takes the double-click word (UAX #29), so a hyphen ends a word and an
  // underscore joins one (WB13a, WB13b): `snake_case` paints whole.
  test('a caret paints only its double-click word: a hyphen ends a word, an underscore does not', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text: para(run('Bold', 'b="1"') + run(' well-known snake_case')),
          }),
        },
      ],
    });
    expect((await e.load(bytes)).ok).toBe(true);
    const [box] = ids(e);
    text(e, box, 0, 1);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    text(e, box, 0, 12);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    text(e, box, 0, 1);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    text(e, box, 0, 18);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    const xml = shapeXml(e, 0);
    expect(runOf(xml, 'known')).toContain('b="1"');
    expect(runOf(xml, 'snake_case')).toContain('b="1"');
    expect(runOf(xml, ' well-')).not.toContain('b="1"');
  });

  test('a whole paragraph also copies its alignment and spacing', async () => {
    const e = await open();
    const [box] = ids(e);
    text(e, box, 0, 0, 16);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    text(e, box, 1, 0, 14);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    const second = shapeXml(e, 0).split('<a:p>')[2]!;
    expect(second).toContain('algn="ctr"');
    expect(second).toContain('<a:spcBef><a:spcPts val="1200"/></a:spcBef>');
  });

  test('text formatting on a shape paints all of its text; a shape without text is refused', async () => {
    const e = await open();
    const [box, ellipse, target] = ids(e);
    text(e, box, 0, 0, 5);
    e.exec({ type: 'toggleItalic' });
    e.exec({ type: 'copyFormat' });
    e.setSelection({ kind: 'shapes', nodeIds: [target] });
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(runOf(shapeXml(e, 2), 'target')).toContain('i="1"');
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.can({ type: 'pasteFormat' })).toEqual({ ok: false, reason: 'disabled.noText' });
    expect(e.exec({ type: 'pasteFormat' })).toMatchObject({ ok: false, reason: 'disabled.noText' });
    expect(e.snapshot().lastRejection?.reason).toBe('disabled.noText');
    // A caret between two spaces has no word to paint.
    text(e, box, 1, 6);
    expect(e.exec({ type: 'pasteFormat' })).toMatchObject({ ok: false, reason: 'disabled.noWord' });
  });
});

describe('format painter: shapes', () => {
  test('copies fill, outline, effects, and p:style, not geometry or position', async () => {
    const e = await open();
    const [, ellipse, target] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.exec({ type: 'toggleFormatPainter' }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [target] });
    const geometry = e.snapshot().geometry!;
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    const xml = shapeXml(e, 2);
    expect(xml).toContain('<a:prstGeom prst="rect">');
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="002060"/></a:solidFill></a:ln><a:effectLst>'
    );
    expect(xml).not.toContain('92D050');
    expect(xml).toContain('</p:spPr><p:style><a:lnRef idx="2">');
    const after = e.snapshot().geometry!;
    expect(after.fill).toBe('#C00000');
    expect([after.x, after.y, after.cx, after.cy]).toEqual([
      geometry.x,
      geometry.y,
      geometry.cx,
      geometry.cy,
    ]);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().geometry?.fill).toBe('#92D050');
    expect(readPptx(e.save()!).ok).toBe(true);
  });
});

describe('format painter: the button', () => {
  test('one press paints once, a double press or `sticky` keeps it on, a press turns it off', async () => {
    const e = await open();
    const [, ellipse, target] = ids(e);
    setSystemTime(new Date('2026-10-29T09:00:00Z'));
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.exec({ type: 'toggleFormatPainter' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('once');
    // A second press later is a single press: it turns the painter off.
    setSystemTime(new Date('2026-10-29T09:00:02Z'));
    expect(e.exec({ type: 'toggleFormatPainter' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('off');

    // Two presses inside the double-click time: sticky, through any number of paints.
    setSystemTime(new Date('2026-10-29T09:00:04Z'));
    e.exec({ type: 'toggleFormatPainter' });
    e.exec({ type: 'toggleFormatPainter' });
    expect(e.snapshot().formatPainter).toBe('sticky');
    e.setSelection({ kind: 'shapes', nodeIds: [target] });
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('sticky');
    // The button stays enabled while the painter is on, so a press can put it away.
    e.setSelection({ kind: 'none' });
    expect(e.can({ type: 'toggleFormatPainter' }).ok).toBe(true);
    setSystemTime(new Date('2026-10-29T09:00:06Z'));
    e.exec({ type: 'toggleFormatPainter' });
    expect(e.snapshot().formatPainter).toBe('off');
    expect(e.can({ type: 'toggleFormatPainter' })).toEqual({
      ok: false,
      reason: 'disabled.noSelection',
    });

    // `sticky` arms it at once.
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.exec({ type: 'toggleFormatPainter', sticky: true }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('sticky');
  });

  test('a slide change, viewing, and a new document turn it off', async () => {
    const e = await open();
    const [, ellipse] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    e.exec({ type: 'toggleFormatPainter', sticky: true });
    e.goToSlide(1);
    expect(e.snapshot().formatPainter).toBe('off');
    e.goToSlide(0);
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    e.exec({ type: 'toggleFormatPainter' });
    e.setMode('view');
    expect(e.snapshot().formatPainter).toBe('off');
    expect(e.can({ type: 'toggleFormatPainter' })).toEqual({
      ok: false,
      reason: 'disabled.viewing',
    });
    e.setMode('edit');
    expect(e.can({ type: 'pasteFormat' }).ok).toBe(true);
    expect((await e.load(deck())).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [ellipse] });
    expect(e.exec({ type: 'pasteFormat' })).toMatchObject({
      ok: false,
      reason: 'disabled.noFormatCopied',
    });
  });

  test('nothing selected and groups are refused with a reason', async () => {
    const e = await open();
    expect(e.exec({ type: 'toggleFormatPainter' })).toMatchObject({
      ok: false,
      reason: 'disabled.noSelection',
    });
    expect(e.snapshot().formatPainter).toBe('off');
    e.selectAll();
    expect(e.exec({ type: 'group' }).ok).toBe(true);
    expect(e.exec({ type: 'toggleFormatPainter' })).toMatchObject({
      ok: false,
      reason: 'disabled.unsupportedElement',
    });
  });
});
