import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { find, readPptx, relationshipsOf, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, fill: solid('4472C4') }) +
          shape({
            id: 3,
            x: IN,
            y: 3 * IN,
            cx: 2 * IN,
            cy: 0.5 * IN,
            text: para(run('A long line of text that wraps in a narrow box')),
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
          }),
      },
    ],
  });

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const ids = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);

describe('shape style commands', () => {
  test('fill, outline color, weight, and dash write valid spPr and show in the snapshot', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    expect(e.exec({ type: 'setShapeFill', color: '#FFC000' }).ok).toBe(true);
    expect(e.snapshot().geometry?.fill).toBe('#FFC000');
    expect(e.exec({ type: 'setLineWidth', points: 6 }).ok).toBe(true);
    // A shape without a visible outline gets one, so the weight shows.
    expect(e.snapshot().geometry?.lineWidth).toBe(6);
    expect(e.snapshot().geometry?.lineColor).not.toBe('none');
    expect(e.exec({ type: 'setLineColor', color: '#C00000' }).ok).toBe(true);
    expect(e.exec({ type: 'setLineDash', dash: 'dash' }).ok).toBe(true);
    const g = e.snapshot().geometry!;
    expect([g.lineColor, g.lineWidth, g.lineDash]).toEqual(['#C00000', 6, 'dash']);
    // Schema order inside spPr and ln: geometry, fill, ln; ln fill before prstDash.
    expect(slideXml(e)).toContain(
      '<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill><a:ln w="76200"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="dash"/></a:ln>'
    );
    expect(e.exec({ type: 'setShapeFill', color: null }).ok).toBe(true);
    expect(e.snapshot().geometry?.fill).toBe('none');
    for (let i = 0; i < 5; i += 1) e.exec({ type: 'undo' });
    expect(e.snapshot().geometry?.fill).toBe('#4472C4');
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('text fitting: resize the shape to fit, shrink, or neither', async () => {
    const e = await open(deck());
    const [, text] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [text] });
    expect(e.snapshot().geometry?.textFitting).toBe('none');
    const before = e.snapshot().geometry!.cy;
    expect(e.exec({ type: 'setAutofit', mode: 'shape' }).ok).toBe(true);
    expect(e.snapshot().geometry?.textFitting).toBe('shape');
    expect(e.snapshot().geometry!.cy).toBeGreaterThan(before);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.exec({ type: 'setAutofit', mode: 'shrink' }).ok).toBe(true);
    const body = find(
      xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root,
      'a:normAutofit'
    );
    expect(body?.attrs.some((a) => a.local === 'fontScale')).toBe(true);
    expect(e.exec({ type: 'setAutofit', mode: 'none' }).ok).toBe(true);
    expect(slideXml(e)).toContain('<a:noAutofit/>');
  });

  test('pictures and groups have no fill; text fitting needs text', async () => {
    const bytes = new Uint8Array(readFileSync('fixtures/pictures.pptx'));
    const e = await open(bytes);
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]] });
    expect(e.can({ type: 'setShapeFill', color: '#000000' })).toEqual({
      ok: false,
      reason: 'disabled.noShapeFill',
    });
    expect(e.can({ type: 'setLineColor', color: '#000000' }).ok).toBe(true);
    expect(e.can({ type: 'setAutofit', mode: 'shape' }).ok).toBe(false);
  });

  test('the Format panel geometry writes back unchanged', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    const g = e.snapshot().geometry!;
    expect(g.nodeId).toBe(rect);
    expect(e.exec({ type: 'setTransform', ...g, cx: 3 * IN }).ok).toBe(true);
    expect(e.snapshot().geometry).toMatchObject({ x: IN, y: IN, cx: 3 * IN, cy: IN });
  });
});

describe('gradients, vertical alignment, lines, and the gallery', () => {
  test('a custom gradient writes a:gradFill and reads back for the dialog', async () => {
    const e = await open(deck());
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    const gradient = {
      kind: 'linear' as const,
      angle: 45,
      stops: [
        { pos: 1, color: '#7F3F00' },
        { pos: 0, color: '#FFC000' },
        { pos: 0.5, color: '#ED7D31' },
      ],
      rotWithShape: false,
    };
    expect(e.exec({ type: 'setShapeGradient', gradient }).ok).toBe(true);
    const g = e.snapshot().geometry!;
    expect(g.fill).toBe('gradient');
    expect(g.gradient).toEqual({
      kind: 'linear',
      angle: 45,
      stops: [
        { pos: 0, color: '#FFC000' },
        { pos: 0.5, color: '#ED7D31' },
        { pos: 1, color: '#7F3F00' },
      ],
      rotWithShape: false,
    });
    expect(slideXml(e)).toContain('<a:lin ang="2700000" scaled="0"/>');
    expect(e.exec({ type: 'setShapeGradient', gradient: { ...gradient, kind: 'radial' } }).ok).toBe(
      true
    );
    expect(e.snapshot().geometry?.gradient?.kind).toBe('radial');
    expect(
      e.can({ type: 'setShapeGradient', gradient: { ...gradient, stops: [gradient.stops[0]] } }).ok
    ).toBe(false);
  });

  test('vertical alignment sets bodyPr/@anchor and shows in the format', async () => {
    const e = await open(deck());
    const [, text] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [text] });
    expect(e.snapshot().format?.verticalAlign).toBe('t');
    expect(e.exec({ type: 'setVerticalAlign', anchor: 'b' }).ok).toBe(true);
    expect(e.snapshot().format?.verticalAlign).toBe('b');
    const [rect] = ids(e);
    e.setSelection({ kind: 'shapes', nodeIds: [rect] });
    expect(e.can({ type: 'setVerticalAlign', anchor: 'ctr' }).ok).toBe(false);
  });

  test('inserted lines and arrows, and their editable ends', async () => {
    const e = await open(deck());
    for (const kind of ['line', 'arrow', 'elbow', 'curved'] as const) {
      expect(e.exec({ type: 'insertLine', kind }).ok).toBe(true);
      const g = e.snapshot().geometry!;
      expect(g.kind).toBe('connector');
      expect(g.lineTail).toBe(kind === 'arrow' ? 'triangle' : 'none');
    }
    expect(e.exec({ type: 'setLineEnds', head: 'oval' }).ok).toBe(true);
    expect(e.exec({ type: 'setLineEnds', tail: 'stealth' }).ok).toBe(true);
    expect(e.snapshot().geometry).toMatchObject({ lineHead: 'oval', lineTail: 'stealth' });
    // headEnd comes before tailEnd in a:ln.
    expect(slideXml(e)).toContain('<a:headEnd type="oval"/><a:tailEnd type="stealth"/>');
    e.setSelection({ kind: 'shapes', nodeIds: [ids(e)[0]] });
    expect(e.can({ type: 'setLineEnds', tail: 'triangle' })).toEqual({
      ok: false,
      reason: 'disabled.notALine',
    });
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('every gallery preset has a preview and inserts', async () => {
    const { SHAPE_GALLERY, shapePreview } = await import('../index.ts');
    const e = await open(deck());
    for (const category of SHAPE_GALLERY)
      for (const preset of category.presets) {
        expect(shapePreview(preset)?.paths.length).toBeGreaterThan(0);
        expect(e.can({ type: 'insertShape', preset }).ok).toBe(true);
      }
    expect(shapePreview('notAShape')).toBeNull();
    expect(e.can({ type: 'insertShape', preset: 'notAShape' }).ok).toBe(false);
    // upArrow is drawn as the mirror of downArrow (the reference file lacks it).
    expect(e.exec({ type: 'insertShape', preset: 'upArrow' }).ok).toBe(true);
    expect(e.session!.layout(0)!.diagnostics.some((d) => d.includes('unknown preset'))).toBe(false);
  });
});

describe('links, clear formatting, backgrounds, and line end points', () => {
  const textSel = (e: PptxEditorInstance, from: number, to: number) => {
    const [, text] = ids(e);
    e.setSelection({
      kind: 'text',
      nodeId: text,
      anchor: { paragraph: 0, offset: from },
      focus: { paragraph: 0, offset: to },
    });
  };

  test('a link is an external relationship with link styling; unsafe schemes are refused', async () => {
    const e = await open(deck());
    textSel(e, 0, 6);
    expect(e.exec({ type: 'setLink', url: 'example.com/a' }).ok).toBe(true);
    textSel(e, 1, 1);
    expect(e.snapshot().format?.link).toBe('https://example.com/a');
    const pkg = e.session!.store.package;
    const part = e.session!.info.slides[0].partName;
    const rels = relationshipsOf(pkg, part).filter((r) => r.external);
    expect(rels.map((r) => r.externalTarget)).toEqual(['https://example.com/a']);
    // Link text takes the theme hyperlink color and an underline.
    const run = e.session!.layout(0)!.items.find((i) => i.kind === 'shape' && i.text)!;
    const props = (
      run as unknown as {
        element: { text: { paragraphs: { runs: { props: { underline: string } }[] }[] } };
      }
    ).element.text.paragraphs[0].runs[0].props;
    expect(props.underline).toBe('sng');
    textSel(e, 0, 6);
    expect(e.can({ type: 'setLink', url: 'javascript:alert(1)' }).ok).toBe(false);
    expect(e.can({ type: 'setLink', url: 'data:text/html,x' }).ok).toBe(false);
    // A caret outside any link links the word it touches: at the end of "long", that word.
    textSel(e, 6, 6);
    expect(e.exec({ type: 'setLink', url: 'https://x.test' }).ok).toBe(true);
    textSel(e, 3, 3);
    expect(e.snapshot().format?.link).toBe('https://x.test');
    textSel(e, 8, 8);
    expect(e.snapshot().format?.link).toBeUndefined();
    e.exec({ type: 'undo' });
    textSel(e, 0, 6);
    expect(e.exec({ type: 'setLink', url: null }).ok).toBe(true);
    textSel(e, 1, 1);
    expect(e.snapshot().format?.link).toBeUndefined();
  });

  test('clear formatting removes character formatting and keeps links', async () => {
    const e = await open(deck());
    textSel(e, 0, 6);
    e.exec({ type: 'toggleBold' });
    e.exec({ type: 'setTextColor', color: '#C00000' });
    e.exec({ type: 'setLink', url: 'https://example.com' });
    expect(e.exec({ type: 'clearFormatting' }).ok).toBe(true);
    textSel(e, 1, 1);
    expect(e.snapshot().format?.bold).toBe(false);
    expect(e.snapshot().format?.link).toBe('https://example.com');
    expect(slideXml(e)).not.toContain('C00000');
  });

  test('slide backgrounds: color, gradient, picture, all slides, and reset', async () => {
    const e = await open(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: '' },
          { layout: 'blank', shapes: '' },
        ],
      })
    );
    expect(e.snapshot().slideBackground.inherited).toBe(true);
    expect(
      e.exec({ type: 'setSlideBackground', background: { kind: 'solid', color: '#002060' } }).ok
    ).toBe(true);
    expect(e.snapshot().slideBackground).toEqual({
      kind: 'solid',
      color: '#002060',
      inherited: false,
    });
    const gradient = {
      kind: 'linear' as const,
      angle: 90,
      stops: [
        { pos: 0, color: '#FFFFFF' },
        { pos: 1, color: '#4472C4' },
      ],
      rotWithShape: true,
    };
    expect(
      e.exec({
        type: 'setSlideBackground',
        background: { kind: 'gradient', gradient },
        allSlides: true,
      }).ok
    ).toBe(true);
    e.goToSlide(1);
    expect(e.snapshot().slideBackground.kind).toBe('gradient');
    const png = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0,
      1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0, 0, 0, 0,
    ]);
    expect(
      e.exec({
        type: 'setSlideBackground',
        background: { kind: 'picture', bytes: png, mime: 'image/png' },
      }).ok
    ).toBe(true);
    expect(e.snapshot().slideBackground.kind).toBe('picture');
    expect(e.exec({ type: 'setSlideBackground', background: null }).ok).toBe(true);
    expect(e.snapshot().slideBackground.inherited).toBe(true);
    expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('moving a line end keeps the other end and sets the direction by flips', async () => {
    const {
      lineEnds,
      lineTransform,
      findItem: find,
    } = await import('../../automation/commands/placement.ts');
    const e = await open(deck());
    e.exec({ type: 'insertLine', kind: 'arrow' });
    const id = (e.snapshot().selection as unknown as { nodeIds: string[] }).nodeIds[0];
    const item = find(e.session!.layout(0)!.items, id)!;
    const ends = lineEnds(item)!;
    // Drag the end to the left of and above the start.
    const target = { x: ends.start.x - IN, y: ends.start.y - IN };
    expect(
      e.exec({ type: 'setTransforms', items: [lineTransform(item, ends.start, target)] }).ok
    ).toBe(true);
    const moved = lineEnds(find(e.session!.layout(0)!.items, id)!)!;
    expect(Math.round(moved.start.x)).toBe(Math.round(ends.start.x));
    expect(Math.round(moved.start.y)).toBe(Math.round(ends.start.y));
    expect(Math.round(moved.end.x)).toBe(Math.round(target.x));
    expect(Math.round(moved.end.y)).toBe(Math.round(target.y));
    expect(slideXml(e)).toContain('flipH="1" flipV="1"');
  });
});
