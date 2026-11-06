// A diagram frame draws its saved drawing (`dsp:drawing`; ECMA-376 Part 1, 21.4) through the
// normal shape cascade and layout, or shows its placeholder when the drawing is missing,
// hostile, or outside the supported subset.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  binaryPartOf,
  elements,
  path,
  PptxDocumentStore,
  xmlPartOf,
  readPptx,
  writePptx,
  type PptxPackage,
} from '../../store/index.ts';
import { resolveSlide, type ResolvedGraphicFrame } from '../../cascade/index.ts';
import {
  apply,
  createFixedMeasurer,
  layoutSlide,
  layoutText,
  type LayoutItem,
  type SlideLayout,
} from '../index.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DRAWING_REL = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
const DATA_CT = 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml';
const DRAWING_CT = 'application/vnd.ms-office.drawingml.diagramDrawing+xml';

const FRAME = { x: 1_000_000, y: 2_000_000, cx: 4_000_000, cy: 2_000_000 };

const frame =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Diagram 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off x="${FRAME.x}" y="${FRAME.y}"/><a:ext cx="${FRAME.cx}" cy="${FRAME.cy}"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdData"/></a:graphicData></a:graphic></p:graphicFrame>`;

const dataModel = (relId = 'rIdDrawing', bg = ''): string =>
  `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst/><dgm:cxnLst/>${bg || '<dgm:bg/><dgm:whole/>'}` +
  `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">` +
  `<dsp:dataModelExt xmlns:dsp="${DSP}" relId="${relId}" minVer="${DGM}"/></a:ext></dgm:extLst></dgm:dataModel>`;

interface SpOptions {
  readonly off?: [number, number];
  readonly ext?: [number, number];
  readonly rot?: number;
  readonly preset?: string;
  readonly text?: string;
  readonly txXfrm?: [number, number, number, number];
  readonly hidden?: boolean;
  readonly flipH?: boolean;
  /** The `dsp:spPr` paint; the default sets a solid fill and a 2 pt white line. */
  readonly paint?: string;
  /** The `dsp:style` references; the default names fill 1, line 2, and the minor font. */
  readonly style?: string;
  readonly rPr?: string;
  /** The `a:pPr` of each paragraph. */
  readonly pPr?: string;
}

const sp = (o: SpOptions = {}): string => {
  const [x, y] = o.off ?? [100_000, 200_000];
  const [cx, cy] = o.ext ?? [1_000_000, 500_000];
  const tx = o.txXfrm;
  return (
    `<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""${o.hidden ? ' hidden="1"' : ''}/><dsp:cNvSpPr/></dsp:nvSpPr>` +
    `<dsp:spPr><a:xfrm${o.rot ? ` rot="${o.rot}"` : ''}${o.flipH ? ' flipH="1"' : ''}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.preset ?? 'roundRect'}"><a:avLst/></a:prstGeom>` +
    (o.paint ??
      '<a:solidFill><a:srgbClr val="9BBB59"><a:hueOff val="0"/></a:srgbClr></a:solidFill>' +
        '<a:ln w="25400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>') +
    '</dsp:spPr><dsp:style>' +
    (o.style ??
      '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef>') +
    '</dsp:style>' +
    (o.text
      ? '<dsp:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
        o.text
          .split('|')
          .map(
            (t) =>
              `<a:p>${o.pPr ?? ''}<a:r>${o.rPr ?? '<a:rPr lang="en-US" sz="1800"/>'}<a:t>${t}</a:t></a:r></a:p>`
          )
          .join('') +
        '</dsp:txBody>'
      : '') +
    (tx
      ? `<dsp:txXfrm><a:off x="${tx[0]}" y="${tx[1]}"/><a:ext cx="${tx[2]}" cy="${tx[3]}"/></dsp:txXfrm>`
      : '') +
    '</dsp:sp>'
  );
};

const drawing = (shapes: string, root = ''): string =>
  `<dsp:drawing xmlns:dgm="${DGM}" xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree>` +
  `<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr>${root}</dsp:grpSpPr>` +
  `${shapes}</dsp:spTree></dsp:drawing>`;

interface DeckOptions {
  /** How many frames on the slide name the one diagram. */
  readonly frames?: number;
  readonly data?: string;
  readonly drawingXml?: string;
  readonly drawingRel?:
    readonly [string, string, string] | readonly [string, string, string, 'External'] | null;
}

function deck(shapes: string, o: DeckOptions = {}): Uint8Array {
  const drawingRel =
    o.drawingRel === undefined
      ? (['rIdDrawing', DRAWING_REL, '../diagrams/drawing1.xml'] as const)
      : o.drawingRel;
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: frame.repeat(o.frames ?? 1),
        rels: [
          ['rIdData', `${R}/diagramData`, '../diagrams/data1.xml'],
          ...(drawingRel ? [drawingRel] : []),
        ],
      },
    ],
    extraParts: [
      ['ppt/diagrams/data1.xml', DATA_CT, o.data ?? dataModel()],
      ['ppt/diagrams/drawing1.xml', DRAWING_CT, o.drawingXml ?? drawing(shapes)],
    ],
  });
}

function open(bytes: Uint8Array): { pkg: PptxPackage; layout: SlideLayout } {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('open failed');
  return { pkg: r.pkg, layout: layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()) };
}

function frameOf(layout: SlideLayout): Extract<LayoutItem, { kind: 'frame' }> {
  const item = layout.items.find((i) => i.kind === 'frame');
  if (item?.kind !== 'frame') throw new Error('no frame');
  return item;
}

const PLACEHOLDER = 'graphic frame rendered as placeholder: diagram';

function expectPlaceholder(bytes: Uint8Array): void {
  const { layout } = open(bytes);
  const f = frameOf(layout);
  expect(f.drawing).toBeUndefined();
  expect(layout.diagnostics).toContain(PLACEHOLDER);
}

describe('diagram saved drawing: shapes', () => {
  test('a shape lays out in the frame space, with the geometry of its own size', () => {
    const { layout } = open(deck(sp()));
    const f = frameOf(layout);
    expect(layout.diagnostics).not.toContain(PLACEHOLDER);
    expect(layout.diagnostics).toContain('approximation: diagram-drawing');
    expect(f.drawing).toHaveLength(1);
    const shape = f.drawing![0]!;
    if (shape.kind !== 'shape') throw new Error('not a shape');
    expect(shape.element.partName).toBe('/ppt/diagrams/drawing1.xml');
    expect(shape.element.fill).toEqual({
      type: 'solid',
      color: { r: 0x9b, g: 0xbb, b: 0x59, a: 1 },
    });
    expect(shape.element.line?.width).toBe(25_400);
    const topLeft = apply(shape.matrix, 0, 0);
    expect(topLeft.x).toBeCloseTo(FRAME.x + 100_000, 3);
    expect(topLeft.y).toBeCloseTo(FRAME.y + 200_000, 3);
    const bottomRight = apply(shape.matrix, 1_000_000, 500_000);
    expect(bottomRight.x).toBeCloseTo(FRAME.x + 1_100_000, 3);
    expect(shape.geometry.paths.length).toBeGreaterThan(0);
  });

  test('a nested group places its children through its child space', () => {
    const group =
      '<dsp:grpSp><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
      '<dsp:grpSpPr><a:xfrm><a:off x="1000000" y="0"/><a:ext cx="2000000" cy="1000000"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="500000"/></a:xfrm></dsp:grpSpPr>' +
      `${sp({ off: [0, 0], ext: [500_000, 500_000] })}</dsp:grpSp>`;
    const f = frameOf(open(deck(group)).layout);
    const g = f.drawing![0]!;
    if (g.kind !== 'group') throw new Error('not a group');
    const child = g.children[0]!;
    // The group doubles its child space: the 500 000 child is 1 000 000 wide on the slide.
    expect(apply(child.matrix, 0, 0).x).toBeCloseTo(FRAME.x + 1_000_000, 3);
    expect(child.element.xfrm.cx).toBeCloseTo(1_000_000, 3);
  });

  test('a hidden shape is not drawn, and the rest of the diagram is', () => {
    const f = frameOf(open(deck(sp({ hidden: true }) + sp({ off: [0, 0] }))).layout);
    expect(f.drawing).toHaveLength(1);
  });

  test('the data model background paints under the shapes over the whole frame', () => {
    const bg =
      '<dgm:bg><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></dgm:bg>' +
      '<dgm:whole><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></dgm:whole>';
    const f = frameOf(open(deck(sp(), { data: dataModel('rIdDrawing', bg) })).layout);
    expect(f.drawing).toHaveLength(2);
    const back = f.drawing![0]!;
    if (back.kind !== 'shape') throw new Error('not a shape');
    expect(back.element.fill).toEqual({ type: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } });
    expect(back.element.line?.width).toBe(12_700);
    expect(apply(back.matrix, FRAME.cx, FRAME.cy)).toEqual({
      x: FRAME.x + FRAME.cx,
      y: FRAME.y + FRAME.cy,
    });
  });
});

describe('diagram saved drawing: style references', () => {
  test('a shape with no own paint takes its fill, line, and text color from the theme', () => {
    const style =
      '<a:lnRef idx="2"><a:srgbClr val="00FF00"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="0000FF"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="FF0000"/></a:fontRef>';
    const shape = sp({ paint: '', style, text: 'Plan' });
    const item = frameOf(open(deck(shape)).layout).drawing![0]!;
    if (item.kind !== 'shape') throw new Error('not a shape');
    const el = item.element;
    // Fill 1 and line 2 of the theme, painted in the reference colors (`phClr`).
    expect(el.fill).toEqual({ type: 'solid', color: { r: 0, g: 0, b: 0xff, a: 1 } });
    expect(el.line?.width).toBe(12_700);
    expect(el.line?.fill).toEqual({ type: 'solid', color: { r: 0, g: 0xff, b: 0, a: 1 } });
    const run = el.text!.paragraphs[0]!.runs[0]!;
    expect(run.props.fill).toEqual({ type: 'solid', color: { r: 0xff, g: 0, b: 0, a: 1 } });
  });

  test('a line with no width takes the width of the theme line it references', () => {
    const paint = '<a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>';
    const item = frameOf(open(deck(sp({ paint }))).layout).drawing![0]!;
    if (item.kind !== 'shape') throw new Error('not a shape');
    expect(item.element.line?.width).toBe(12_700);
    expect(item.element.line?.fill).toEqual({
      type: 'solid',
      color: { r: 255, g: 255, b: 255, a: 1 },
    });
  });
});

describe('diagram saved drawing: turns and flips', () => {
  test('a turned, flipped shape turns and mirrors about its own center', () => {
    const shape = sp({ off: [0, 0], ext: [1_000_000, 500_000], rot: 5_400_000, flipH: true });
    const item = frameOf(open(deck(shape)).layout).drawing![0]!;
    // The center stays where the file puts it.
    const c = apply(item.matrix, 500_000, 250_000);
    expect(c.x).toBeCloseTo(FRAME.x + 500_000, 3);
    expect(c.y).toBeCloseTo(FRAME.y + 250_000, 3);
    // A quarter turn and a mirror: the local x axis points down, the y axis points left.
    const [a, b, cc, d] = item.matrix;
    expect([a, b, cc, d].map((v) => Math.round(v) + 0)).toEqual([0, -1, -1, 0]);
  });

  test('text in a flipped group stays readable', () => {
    const group =
      '<dsp:grpSp><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
      '<dsp:grpSpPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></dsp:grpSpPr>' +
      `${sp({ off: [0, 0], ext: [1_000_000, 500_000], text: 'Plan', txXfrm: [0, 0, 1_000_000, 500_000] })}</dsp:grpSp>`;
    const g = frameOf(open(deck(group)).layout).drawing![0]!;
    if (g.kind !== 'group') throw new Error('not a group');
    const item = g.children[0]!;
    if (item.kind !== 'shape' || !item.textMatrix || !item.text) throw new Error('no text');
    const m = item.textMatrix;
    // Not mirrored: the matrix keeps its orientation.
    expect(m[0] * m[3] - m[1] * m[2]).toBeGreaterThan(0);
    // The flipped group moves the text rectangle to its mirrored place on the right.
    const left = apply(m, item.text.box.x, 0).x;
    const right = apply(m, item.text.box.x + item.text.box.width, 0).x;
    expect(Math.min(left, right)).toBeCloseTo(FRAME.x + 1_000_000, 3);
    expect(Math.max(left, right)).toBeCloseTo(FRAME.x + 2_000_000, 3);
  });
});

describe('diagram saved drawing: text rectangle', () => {
  test('text lays out in the dsp:txXfrm rectangle, not in the geometry text rectangle', () => {
    const shape = sp({ text: 'Plan', txXfrm: [150_000, 250_000, 400_000, 300_000] });
    const item = frameOf(open(deck(shape)).layout).drawing![0]!;
    if (item.kind !== 'shape' || !item.text || !item.textMatrix) throw new Error('no text');
    expect(item.text.box.width).toBeCloseTo(400_000, 3);
    expect(item.text.box.height).toBeCloseTo(300_000, 3);
    const origin = apply(item.textMatrix, item.text.box.x, item.text.box.y);
    expect(origin.x).toBeCloseTo(FRAME.x + 150_000, 3);
    expect(origin.y).toBeCloseTo(FRAME.y + 250_000, 3);
  });

  test("the text rectangle's turn adds to the shape's turn", () => {
    const shape = sp({ rot: 5_400_000, text: 'Plan', txXfrm: [0, 0, 400_000, 200_000] });
    const item = frameOf(open(deck(shape)).layout).drawing![0]!;
    if (item.kind !== 'shape' || !item.textMatrix) throw new Error('no text');
    const m = item.textMatrix;
    expect(m[0]).toBeCloseTo(0, 6);
    expect(m[1]).toBeCloseTo(1, 6);
    // Turned about its own center, which stays where the file puts it.
    const c = apply(m, 200_000, 100_000);
    expect(c.x).toBeCloseTo(FRAME.x + 200_000, 3);
    expect(c.y).toBeCloseTo(FRAME.y + 100_000, 3);
  });

  test('percent line spacing scales the face line, not 1.2 × the size', () => {
    // The fixed measurer's face: ascent 0.9 and descent 0.25, so its line is 1.15 × the size.
    const shape = sp({
      text: 'Plan|Write',
      txXfrm: [0, 0, 1_000_000, 1_000_000],
      rPr: '<a:rPr lang="en-US" sz="2000"/>',
      pPr: '<a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr>',
    });
    const item = frameOf(open(deck(shape)).layout).drawing![0]!;
    if (item.kind !== 'shape' || !item.text?.lines[1]) throw new Error('no text');
    const [first, second] = item.text.lines;
    expect(first!.height).toBeCloseTo(0.9 * 1.15 * 20 * 12_700, 3);
    expect(second!.top - first!.top).toBeCloseTo(0.9 * 1.15 * 20 * 12_700, 3);
    // The same body in a slide shape spaces its lines on 1.2 × the size.
    const box = { x: 0, y: 0, width: 1_000_000, height: 1_000_000 };
    const plain = layoutText(item.element.text!, box, createFixedMeasurer());
    expect(plain.lines[0]!.height).toBeCloseTo(0.9 * 1.2 * 20 * 12_700, 3);
  });

  test('a shape with no dsp:txXfrm uses its geometry text rectangle', () => {
    const item = frameOf(open(deck(sp({ text: 'Plan' }))).layout).drawing![0]!;
    if (item.kind !== 'shape') throw new Error('not a shape');
    expect(item.element.textXfrm).toBeUndefined();
    expect(item.text).toBeDefined();
  });
});

describe('diagram saved drawing: whole or not at all', () => {
  test('one shape with an unknown preset geometry gives the placeholder', () => {
    expectPlaceholder(deck(sp() + sp({ preset: 'notAPreset' })));
    const { layout } = open(deck(sp({ preset: 'notAPreset' })));
    expect(layout.diagnostics.some((d) => d.startsWith('unknown preset'))).toBe(false);
  });

  test('an element outside the drawing schema gives the placeholder', () => {
    expectPlaceholder(deck(sp() + '<dsp:pic/>'));
    expectPlaceholder(
      deck(`${sp()}<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`)
    );
  });

  test('a shape with no box gives the placeholder', () => {
    expectPlaceholder(deck(sp().replace(/<a:xfrm>.*?<\/a:xfrm>/, '')));
  });
});

describe('diagram saved drawing: missing and hostile parts', () => {
  test('no drawing relationship gives the placeholder', () => {
    expectPlaceholder(deck(sp(), { drawingRel: null }));
  });

  test('a data model with no dsp:dataModelExt gives the placeholder', () => {
    expectPlaceholder(
      deck(sp(), { data: `<dgm:dataModel xmlns:dgm="${DGM}"><dgm:ptLst/></dgm:dataModel>` })
    );
  });

  test('a relationship of another type gives the placeholder', () => {
    expectPlaceholder(
      deck(sp(), { drawingRel: ['rIdDrawing', `${R}/image`, '../diagrams/drawing1.xml'] })
    );
  });

  test('an external drawing target is never followed', () => {
    expectPlaceholder(
      deck(sp(), {
        drawingRel: ['rIdDrawing', DRAWING_REL, 'https://example.com/drawing1.xml', 'External'],
      })
    );
  });

  test('a drawing part with another root gives the placeholder', () => {
    expectPlaceholder(deck(sp(), { drawingXml: `<dgm:dataModel xmlns:dgm="${DGM}"/>` }));
  });

  test('a drawing with a document type declaration is refused by the bounded reader', () => {
    // The same drawing draws without the declaration, so the declaration alone refuses it.
    expect(frameOf(open(deck(sp())).layout).drawing).toHaveLength(1);
    expectPlaceholder(deck(sp(), { drawingXml: `<!DOCTYPE d [<!ENTITY e "x">]>${drawing(sp())}` }));
  });

  test('a drawing past the shape limit gives the placeholder', () => {
    expectPlaceholder(deck(sp().repeat(513)));
    expect(frameOf(open(deck(sp().repeat(512))).layout).drawing).toHaveLength(512);
  });

  test('frames past the slide budget give the placeholder, however many name one drawing', () => {
    // Eight frames of 512 shapes fill the slide budget of 4096; the ninth frame does not draw.
    const { layout } = open(deck(sp().repeat(512), { frames: 9 }));
    const frames = layout.items.filter((i) => i.kind === 'frame');
    expect(frames).toHaveLength(9);
    expect(frames.slice(0, 8).every((f) => f.kind === 'frame' && f.drawing?.length === 512)).toBe(
      true
    );
    const last = frames[8]!;
    expect(last.kind === 'frame' && last.drawing).toBeFalsy();
    expect(layout.diagnostics).toContain(PLACEHOLDER);
  });

  test('groups nested past the depth limit give the placeholder', () => {
    const nest = (depth: number): string =>
      depth === 0
        ? sp()
        : `<dsp:grpSp><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${nest(depth - 1)}</dsp:grpSp>`;
    expect(frameOf(open(deck(nest(8))).layout).drawing).toHaveLength(1);
    expectPlaceholder(deck(nest(9)));
  });

  test('a placeholder names why: no drawing, unsupported content, or a limit', () => {
    const reason = (bytes: Uint8Array): unknown =>
      frameOf(open(bytes).layout).element.drawingFallback;
    expect(reason(deck(sp()))).toBeUndefined();
    expect(reason(deck(sp(), { drawingRel: null }))).toBe('no-drawing');
    expect(reason(deck(sp() + '<dsp:pic/>'))).toBe('unsupported');
    expect(reason(deck(sp({ preset: 'notAPreset' })))).toBe('unsupported');
    expect(reason(deck(sp().repeat(513)))).toBe('limit');
  });
});

describe('diagram saved drawing: preservation', () => {
  test('an edit and a save keep the diagram parts byte-identical', () => {
    const bytes = deck(sp({ text: 'Plan', txXfrm: [0, 0, 400_000, 200_000] }));
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('open failed');
    const before = binaryPartOf(r.pkg, '/ppt/diagrams/drawing1.xml')!;
    const data = binaryPartOf(r.pkg, '/ppt/diagrams/data1.xml')!;
    const store = new PptxDocumentStore(r.pkg);
    const resolved = resolveSlide(store.package, 0)!;
    const f = resolved.elements.find((e) => e.kind === 'graphicFrame') as ResolvedGraphicFrame;
    expect(f.drawing).toHaveLength(1);
    const cNvPr = path(
      elements(xmlPartOf(store.package, f.partName)!.root.children[0] as never).find(
        (e) => e.local === 'spTree'
      ),
      'p:graphicFrame',
      'p:nvGraphicFramePr',
      'p:cNvPr'
    )!;
    const result = store.transact([
      {
        op: 'setAttribute',
        part: f.partName,
        node: cNvPr.id,
        name: { ns: '', local: 'name' },
        value: 'Renamed',
      },
    ]);
    expect(result.ok).toBe(true);
    const again = readPptx(writePptx(store.package));
    if (!again.ok) throw new Error('reopen failed');
    expect(binaryPartOf(again.pkg, '/ppt/diagrams/drawing1.xml')).toEqual(before);
    expect(binaryPartOf(again.pkg, '/ppt/diagrams/data1.xml')).toEqual(data);
    const reopened = frameOf(layoutSlide(resolveSlide(again.pkg, 0)!, createFixedMeasurer()));
    expect(reopened.drawing).toHaveLength(1);
  });
});
