// Shapes from a DrawingML clipboard package (`clipboard-drawing.ts`): the package converts to
// this editor's shape payload, pastes as editable shapes, and a host's reader turns a picture
// paste into those shapes.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  elements,
  find,
  relationshipsOf,
  textOf,
  writeZip,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { drawingClipboardPayload } from '../clipboard-drawing.ts';
import { createClipboardHost, type ClipboardEditor } from '../clipboard-host.ts';
import { parsePayload } from '../clipboard.ts';
import type { EditorCommand } from '../../contracts/index.ts';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
/** A 1x1 PNG. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

const xfrm = (x: number, y: number, cx: number, cy: number): string =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;

/** A shape style, as shapes drawn in presentation applications have. */
const STYLE =
  `<a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1">` +
  `<a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/>` +
  `</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>`;

/**
 * A text shape as the package holds it: DrawingML, in its schema order (`nvSpPr`, `spPr`,
 * `txSp`, `style`), with its text in `a:txSp`.
 */
const textShape = (id: number, text: string, x = 457200, y = 457200): string =>
  `<a:sp><a:nvSpPr><a:cNvPr id="${id}" name="Shape ${id}"/><a:cNvSpPr/></a:nvSpPr>` +
  `<a:spPr>${xfrm(x, y, 2743200, 1828800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:spPr>` +
  `<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/>` +
  `<a:t>${text}</a:t></a:r></a:p></a:txBody><a:useSpRect/></a:txSp>${STYLE}` +
  `<a:extLst><a:ext uri="{00000000-0000-0000-0000-000000000000}"/></a:extLst></a:sp>`;

/**
 * A DrawingML clipboard package: the copied shapes in a locked canvas whose child space is the
 * source slide, as presentation applications write it.
 */
function drawingPackage(
  shapes: string,
  opts: { ext?: [number, number]; chExt?: [number, number]; image?: boolean } = {}
): Uint8Array {
  const [cx, cy] = opts.ext ?? [2743200, 1828800];
  const [chx, chy] = opts.chExt ?? [cx, cy];
  const drawing =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:graphic ${A}>` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
    `<lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
    `<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr><a:xfrm>` +
    `<a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="457200" y="457200"/>` +
    `<a:chExt cx="${chx}" cy="${chy}"/></a:xfrm></a:grpSpPr>${shapes}</lc:lockedCanvas>` +
    `</a:graphicData></a:graphic>`;
  const rel = (id: string, type: string, target: string, external = false): string =>
    `<Relationship Id="${id}" Type="${RELS}/${type}" Target="${target}"` +
    `${external ? ' TargetMode="External"' : ''}/>`;
  const rels = (body: string): Uint8Array =>
    enc(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships ` +
        `xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
    );
  const entries = new Map<string, Uint8Array>([
    [
      '/[Content_Types].xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8"?><Types ` +
          `xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="png" ContentType="image/png"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/clipboard/drawings/drawing1.xml" ` +
          `ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
      ),
    ],
    ['/_rels/.rels', rels(rel('rId1', 'drawing', 'clipboard/drawings/drawing1.xml'))],
    ['/clipboard/drawings/drawing1.xml', enc(drawing)],
    [
      '/clipboard/drawings/_rels/drawing1.xml.rels',
      rels(
        rel('rId1', 'image', '../media/image1.png') +
          rel('rId2', 'hyperlink', 'https://example.com/', true) +
          rel('rId3', 'chart', '../charts/chart1.xml')
      ),
    ],
  ]);
  if (opts.image) entries.set('/clipboard/media/image1.png', PNG);
  return writeZip(entries);
}

const picture =
  `<a:pic><a:nvPicPr><a:cNvPr id="5" name="Picture 5"><a:hlinkClick ${R} r:id="rId2"/>` +
  `</a:cNvPr><a:cNvPicPr><a:picLocks noChangeAspect="1"/></a:cNvPicPr></a:nvPicPr>` +
  `<a:blipFill><a:blip ${R} r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
  `<a:spPr>${xfrm(457200, 457200, 914400, 914400)}<a:prstGeom prst="rect"><a:avLst/>` +
  `</a:prstGeom></a:spPr></a:pic>`;
const group =
  `<a:grpSp><a:nvGrpSpPr><a:cNvPr id="2" name="Group 1"/><a:cNvGrpSpPr/></a:nvGrpSpPr>` +
  `<a:grpSpPr><a:xfrm rot="900000"><a:off x="457200" y="457200"/><a:ext cx="2743200" ` +
  `cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="1828800"/></a:xfrm>` +
  `</a:grpSpPr><a:txSp><a:txBody><a:bodyPr/><a:p/></a:txBody><a:useSpRect/></a:txSp>` +
  `${textShape(3, 'in group', 0, 0)}${textShape(4, 'second', 914400, 914400)}</a:grpSp>`;
const connector =
  `<a:cxnSp><a:nvCxnSpPr><a:cNvPr id="6" name="Connector 6"/><a:cNvCxnSpPr>` +
  `<a:stCxn id="3" idx="1"/></a:cNvCxnSpPr></a:nvCxnSpPr><a:spPr>${xfrm(0, 0, 914400, 0)}` +
  `<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom></a:spPr></a:cxnSp>`;
const chart =
  `<a:graphicFrame><a:nvGraphicFramePr><a:cNvPr id="7" name="Chart 7"/><a:cNvGraphicFramePr/>` +
  `</a:nvGraphicFramePr>${xfrm(0, 0, 914400, 914400)}<a:graphic><a:graphicData ` +
  `uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart ` +
  `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${R} r:id="rId3"/>` +
  `</a:graphicData></a:graphic></a:graphicFrame>`;
const table =
  `<a:graphicFrame><a:nvGraphicFramePr><a:cNvPr id="8" name="Table 8"/><a:cNvGraphicFramePr/>` +
  `</a:nvGraphicFramePr><a:graphic><a:graphicData ` +
  `uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>` +
  `<a:gridCol w="914400"/></a:tblGrid><a:tr h="370840"><a:tc><a:txBody><a:bodyPr/>` +
  `<a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>cell</a:t></a:r></a:p></a:txBody>` +
  `<a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic>${xfrm(0, 0, 914400, 370840)}` +
  `</a:graphicFrame>`;

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildPptx({ slides: [{ shapes: '' }] }))).ok).toBe(true);
  return e;
}
const slidePart = (e: PptxEditorInstance): string => e.session!.info.slides[0]!.partName;
const tree = (e: PptxEditorInstance): Element =>
  find(xmlPartOf(e.session!.store.package, slidePart(e))!.root, 'p:spTree')!;
const top = (e: PptxEditorInstance): Element[] =>
  elements(tree(e)).filter((c) => !c.local.endsWith('SpPr'));
const offOf = (el: Element): [string | undefined, string | undefined] => {
  const props = child(el, 'p:spPr') ?? child(el, 'p:grpSpPr');
  const off = find(props ?? el, 'a:off');
  return [off && attr(off, 'x'), off && attr(off, 'y')];
};

async function paste(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = await open();
  const payload = drawingClipboardPayload(bytes);
  expect(payload).not.toBeNull();
  expect((await e.exec({ type: 'paste', payload: payload! })).ok).toBe(true);
  return e;
}

describe('a DrawingML clipboard package', () => {
  test('a shape pastes as an editable shape at its place, with its text and fill', async () => {
    const e = await paste(drawingPackage(textShape(2, 'Hello')));
    const [sp] = top(e);
    expect(sp?.local).toBe('sp');
    expect(textOf(child(sp!, 'p:txBody')!)).toBe('Hello');
    expect(find(sp!, 'a:schemeClr') && attr(find(sp!, 'a:schemeClr')!, 'val')).toBe('accent1');
    expect(offOf(sp!)).toEqual(['457200', '457200']);
    // The PresentationML shape has its application data, and no element of the canvas.
    expect(child(child(sp!, 'p:nvSpPr')!, 'p:nvPr')).toBeDefined();
    expect(find(sp!, 'a:txSp')).toBeUndefined();
    expect(child(sp!, 'p:extLst')).toBeUndefined();
    // Schema order: the style goes before the text, unlike in the package.
    expect(elements(sp!).map((c) => c.local)).toEqual(['nvSpPr', 'spPr', 'style', 'txBody']);
    expect(e.snapshot().selection).toMatchObject({ kind: 'shapes' });
  });

  test('a group keeps its members, its transform, and its children space', async () => {
    const e = await paste(drawingPackage(group));
    const [grp] = top(e);
    expect(grp?.local).toBe('grpSp');
    // The group's own text has no place in a PresentationML group.
    expect(elements(grp!).map((c) => c.local)).toEqual(['nvGrpSpPr', 'grpSpPr', 'sp', 'sp']);
    const members = elements(grp!).filter((c) => c.local === 'sp');
    expect(members.map((m) => textOf(child(m, 'p:txBody')!))).toEqual(['in group', 'second']);
    expect(attr(find(child(grp!, 'p:grpSpPr')!, 'a:xfrm')!, 'rot')).toBe('900000');
  });

  test('a picture carries its image; a link to outside the package does not travel', async () => {
    const bytes = drawingPackage(picture, { image: true });
    const payload = drawingClipboardPayload(bytes)!;
    expect(payload.relationships).toHaveLength(1);
    expect(payload.relationships[0]).toMatchObject({ id: 'rId1', contentType: 'image/png' });
    const e = await paste(bytes);
    const [pic] = top(e);
    expect(pic?.local).toBe('pic');
    const embed = find(pic!, 'a:blip')!.attrs.find((a) => a.local === 'embed')!.value;
    const rel = relationshipsOf(e.session!.store.package, slidePart(e)).find((r) => r.id === embed);
    expect(rel?.target).toMatch(/^\/ppt\/media\/image\d*\.png$/);
    expect(find(pic!, 'a:hlinkClick')).toBeUndefined();
  });

  test('a picture whose image is missing pastes no broken relationship', () => {
    const payload = drawingClipboardPayload(drawingPackage(picture));
    expect(payload?.relationships).toEqual([]);
  });

  test('connectors and tables convert; a chart needs parts that do not travel', () => {
    const payload = drawingClipboardPayload(
      drawingPackage(textShape(3, 'a') + connector + chart + table)
    );
    expect(payload?.shapes.map((s) => s.local)).toEqual(['sp', 'cxnSp', 'graphicFrame']);
    const parsed = parsePayload(payload);
    expect(parsed?.shapes).toHaveLength(3);
    const frame = payload!.shapes[2]!;
    expect(frame.children?.map((c) => ('text' in c ? '' : c.local))).toEqual([
      'nvGraphicFramePr',
      'xfrm',
      'graphic',
    ]);
  });

  test('a canvas that scales its children places the shapes at the canvas scale', () => {
    const payload = drawingClipboardPayload(
      drawingPackage(textShape(2, 'x', 457200 + 914400, 457200), {
        ext: [1371600, 914400],
        chExt: [2743200, 1828800],
      })
    )!;
    const json = JSON.stringify(payload.shapes[0]);
    // Half size: the offset from the canvas origin halves, and so does the extent.
    expect(json).toContain('{"ns":"","local":"x","value":"914400"}');
    expect(json).toContain('{"ns":"","local":"cx","value":"1371600"}');
  });

  test('scaled positions and sizes stay within the coordinate range', () => {
    const payload = drawingClipboardPayload(
      drawingPackage(textShape(2, 'x'), { ext: [1e14, 1828800], chExt: [1, 1828800] })
    )!;
    const json = JSON.stringify(payload.shapes[0]);
    expect(json).toContain('{"ns":"","local":"cx","value":"27273042316900"}');
  });

  test('a package past the limits of a paste gives nothing, so the picture pastes', () => {
    expect(drawingClipboardPayload(drawingPackage(textShape(2, 'x').repeat(501)))).toBeNull();
  });

  test('a view of the bytes from another buffer reads the same', () => {
    const bytes = drawingPackage(textShape(2, 'x'));
    const padded = new Uint8Array(bytes.length + 8);
    padded.set(bytes, 4);
    expect(drawingClipboardPayload(padded.subarray(4, 4 + bytes.length))).not.toBeNull();
  });

  test('bytes that are not such a package give nothing', () => {
    expect(drawingClipboardPayload(new Uint8Array())).toBeNull();
    expect(drawingClipboardPayload(PNG)).toBeNull();
    expect(drawingClipboardPayload(buildPptx({ slides: [{ shapes: '' }] }))).toBeNull();
    // A canvas with nothing that converts.
    expect(drawingClipboardPayload(drawingPackage(chart))).toBeNull();
    // A shape without its drawing properties.
    expect(drawingClipboardPayload(drawingPackage('<a:sp><a:nvSpPr/><a:spPr/></a:sp>'))).toBeNull();
  });
});

/** A clipboard event's data, as a browser gives it. */
function transfer(data: Record<string, string>, files: File[] = []): DataTransfer {
  return {
    getData: (t: string) => data[t] ?? '',
    setData: (t: string, v: string) => void (data[t] = v),
    get types() {
      return Object.keys(data);
    },
    items: files.map((f) => ({ kind: 'file', type: f.type, getAsFile: () => f })),
    files,
  } as unknown as DataTransfer;
}

async function host(readDrawing?: ClipboardEditor['readDrawing']) {
  const e = await open();
  const at = { slide: 0 };
  const ran: EditorCommand[] = [];
  const editor: ClipboardEditor = {
    session: () => e.session,
    slide: () => at.slide,
    selection: () => e.snapshot().selection,
    layout: () => e.session!.layout(0),
    readOnly: () => false,
    emit: () => undefined,
    deleteText: () => undefined,
    cutShapes: () => undefined,
    insertText: () => undefined,
    refuse: () => undefined,
    exec: (command) => void ran.push(command),
    selectedSlides: () => [],
    ...(readDrawing ? { readDrawing } : {}),
  };
  return { e, at, ran, clip: createClipboardHost(editor) };
}

const until = async (done: () => boolean): Promise<void> => {
  for (let i = 0; i < 100 && !done(); i += 1) await new Promise((r) => setTimeout(r, 5));
};

describe('the host reader', () => {
  const file = (): File => new File([PNG], 'image.png', { type: 'image/png' });

  test('a picture paste pastes the shapes the host reads', async () => {
    const { ran, clip } = await host(async () => drawingPackage(textShape(2, 'Hello')));
    expect(clip.pasteFrom(transfer({}, [file()]))).toBe(true);
    await until(() => ran.length > 0);
    expect(ran[0]).toMatchObject({ type: 'paste' });
  });

  test('without a package, or with a failing reader, the picture pastes', async () => {
    for (const read of [
      () => null,
      () => new Uint8Array([1, 2]),
      () => Promise.reject(new Error('x')),
    ]) {
      const { ran, clip } = await host(read);
      expect(clip.pasteFrom(transfer({}, [file()]))).toBe(true);
      await until(() => ran.length > 0);
      expect(ran[0]).toMatchObject({ type: 'insertImage' });
    }
  });

  test('HTML text still wins over the picture, and the reader is not asked', async () => {
    let asked = 0;
    const { ran, clip } = await host(() => (asked++, null));
    const html = '<meta charset="utf-8"><p>Some text</p>';
    expect(clip.pasteFrom(transfer({ 'text/html': html }, [file()]))).toBe(true);
    expect(ran[0]).toMatchObject({ type: 'pasteText' });
    expect(asked).toBe(0);
  });

  test('a paste into edited text does not ask the reader', async () => {
    let asked = 0;
    const { e, clip } = await host(() => (asked++, null));
    expect((await e.exec({ type: 'insertTextBox' })).ok).toBe(true);
    expect(e.snapshot().selection.kind).toBe('text');
    clip.pasteFrom(transfer({}, [file()]));
    expect(asked).toBe(0);
  });

  test('a paste that ends on another slide pastes nothing', async () => {
    let release = (): void => undefined;
    const { at, ran, clip } = await host(
      () => new Promise((done) => (release = () => done(drawingPackage(textShape(2, 'Hello')))))
    );
    expect(clip.pasteFrom(transfer({}, [file()]))).toBe(true);
    await new Promise((r) => setTimeout(r, 5));
    at.slide = 1;
    release();
    await new Promise((r) => setTimeout(r, 20));
    expect(ran).toEqual([]);
  });

  test('a later paste wins over a paste still reading', async () => {
    const releases: (() => void)[] = [];
    const { ran, clip } = await host(
      () =>
        new Promise((done) =>
          releases.push(() => done(drawingPackage(textShape(2, `n${releases.length}`))))
        )
    );
    clip.pasteFrom(transfer({}, [file()]));
    clip.pasteFrom(transfer({}, [file()]));
    await until(() => releases.length === 2);
    releases[1]!();
    releases[0]!();
    await until(() => ran.length > 0);
    await new Promise((r) => setTimeout(r, 20));
    expect(ran).toHaveLength(1);
  });

  test('a reader that never answers pastes the picture after 2 seconds', async () => {
    const { ran, clip } = await host(() => new Promise<null>(() => undefined));
    expect(clip.pasteFrom(transfer({}, [file()]))).toBe(true);
    await new Promise((r) => setTimeout(r, 2100));
    await until(() => ran.length > 0);
    expect(ran[0]).toMatchObject({ type: 'insertImage' });
  }, 5000);
});
