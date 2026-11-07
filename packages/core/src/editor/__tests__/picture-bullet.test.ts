// Picture bullets (`a:buBlip`): the cascade finds the picture through the relationships of the
// part that holds the level, layout sizes it from the bullet size, a picture that cannot be read
// draws no bullet, and `setBulletFormat` sets one from image bytes as one undo step.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { PNG } from 'pngjs';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { bulletPictureProblem } from '../bullet-format.ts';
import { readBulletPicture } from '../bullets-dialog-fields.ts';
import { attr, child, elements, find, readPptx, writePptx, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import type { GlyphRun } from '../../layout/index.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const EMU_PER_PT = 12_700;

/** A `w` × `h` opaque PNG. */
function png(w: number, h: number): Uint8Array<ArrayBuffer> {
  const image = new PNG({ width: w, height: h });
  image.data.fill(200);
  return new Uint8Array(PNG.sync.write(image));
}

const BLIP = '<a:buBlip><a:blip r:embed="rIdBu"/></a:buBlip>';

function deck(pPr: string, opts: { image?: boolean; sz?: string } = {}): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 3 * IN,
          text: para(run('one', opts.sz ?? 'sz="1800"'), pPr) + para(run('two')),
        }),
        ...(opts.image === false ? {} : { rels: [['rIdBu', R_IMAGE, '../media/image1.png']] }),
      },
    ],
    media:
      opts.image === false
        ? []
        : [{ name: 'image1.png', bytes: png(40, 20), contentType: 'image/png' }],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const bulletOf = (e: PptxEditorInstance, line = 0): GlyphRun | undefined => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text)!;
  const runs = item.kind === 'shape' ? item.text!.lines[line]!.runs : [];
  return runs.find((r) => r.kind === 'bullet');
};

const shapeId = (e: PptxEditorInstance): string => {
  const root = xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
  return elements(find(root, 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!
    .id;
};

const select = (e: PptxEditorInstance, from = 0, to = from) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: from, offset: 0 },
    focus: { paragraph: to, offset: 1 },
  });

describe('picture bullets', () => {
  test('a slide picture bullet is 0.7 times the bullet size high, with the picture aspect', async () => {
    const e = await open(
      deck(`<a:pPr marL="342900" indent="-342900"><a:buSzPct val="150000"/>${BLIP}</a:pPr>`)
    );
    const b = bulletOf(e)!;
    expect(b.picture).toBeDefined();
    expect(b.text).toBe('');
    const height = 18 * 1.5 * 0.7 * EMU_PER_PT;
    expect(b.picture!.height).toBeCloseTo(height, 3);
    // 40 × 20 pixels: twice as wide as high.
    expect(b.x[1] - b.x[0]).toBeCloseTo(height * 2, 3);
    select(e);
    expect(e.snapshot().format?.bulletFormat).toEqual({ kind: 'picture', sizePercent: 150 });
    e.destroy();
  });

  test('a point size sets the picture height', async () => {
    const e = await open(deck(`<a:pPr><a:buSzPts val="2400"/>${BLIP}</a:pPr>`));
    expect(bulletOf(e)!.picture!.height).toBeCloseTo(24 * 0.7 * EMU_PER_PT, 3);
    e.destroy();
  });

  test('a missing relationship or picture draws no bullet and keeps the markup', async () => {
    const e = await open(deck(`<a:pPr>${BLIP}</a:pPr>`, { image: false }));
    expect(bulletOf(e)).toBeUndefined();
    const saved = await writePptx(e.session!.store.package);
    expect(strFromU8(unzipSync(saved)['ppt/slides/slide1.xml']!)).toContain('<a:buBlip>');
    e.destroy();
  });

  test('a master level finds its picture through the master relationships', async () => {
    // The master's body style level 1 takes a picture bullet related to the master part.
    const files = unzipSync(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 2,
              name: 'T',
              x: 0,
              y: 0,
              cx: 0,
              cy: 0,
              ph: '<p:ph idx="1"/>',
              text: para(run('body')),
            }),
          },
        ],
      })
    );
    const masterXml = strFromU8(files['ppt/slideMasters/slideMaster1.xml']!);
    const lvl1 = /<a:lvl1pPr[^>]*>/.exec(masterXml.slice(masterXml.indexOf('<p:bodyStyle>')))![0];
    files['ppt/slideMasters/slideMaster1.xml'] = strToU8(
      masterXml.replace(
        `<p:bodyStyle>${lvl1}`,
        `<p:bodyStyle>${lvl1}<a:buSzPct val="100000"/><a:buBlip><a:blip r:embed="rIdMasterBu"/></a:buBlip>`
      )
    );
    const relsName = 'ppt/slideMasters/_rels/slideMaster1.xml.rels';
    files[relsName] = strToU8(
      strFromU8(files[relsName]!).replace(
        '</Relationships>',
        `<Relationship Id="rIdMasterBu" Type="${R_IMAGE}" Target="../media/bullet.png"/></Relationships>`
      )
    );
    files['ppt/media/bullet.png'] = png(10, 10);
    files['[Content_Types].xml'] = strToU8(
      strFromU8(files['[Content_Types].xml']!).replace(
        '<Default ',
        '<Default Extension="png" ContentType="image/png"/><Default '
      )
    );
    const e = await open(zipSync(files));
    const b = bulletOf(e);
    expect(b?.picture?.part).toBe('/ppt/slideMasters/slideMaster1.xml');
    expect(b?.picture?.rId).toBe('rIdMasterBu');
    expect(b!.x[1] - b!.x[0]).toBeCloseTo(b!.picture!.height, 3);
    e.destroy();
  });
});

describe('setBulletFormat with a picture', () => {
  test('adds one media part and a:buBlip on every selected paragraph, as one undo step', async () => {
    const e = await open(deck('', { image: false }));
    select(e, 0, 1);
    const before = e.session!.store.package.parts.size;
    const r = e.exec({
      type: 'setBulletFormat',
      format: { kind: 'picture', picture: png(16, 16), sizePercent: 120 },
    });
    expect(r.ok).toBe(true);
    const pkg = e.session!.store.package;
    expect(pkg.parts.size).toBe(before + 1);
    const root = xmlPartOf(pkg, '/ppt/slides/slide1.xml')!.root;
    const ps = elements(find(root, 'p:txBody')!).filter((c) => c.local === 'p');
    const ids = ps.map((p) =>
      attr(child(child(child(p, 'a:pPr'), 'a:buBlip'), 'a:blip'), 'r:embed')
    );
    expect(ids[0]).toBeTruthy();
    expect(ids[1]).toBe(ids[0]);
    expect(bulletOf(e)?.picture?.height).toBeCloseTo(18 * 1.2 * 0.7 * EMU_PER_PT, 3);
    expect(e.snapshot().format?.bullets).toBe(true);
    // One undo removes the picture bullets and the part.
    e.exec({ type: 'undo' });
    expect(e.session!.store.package.parts.size).toBe(before);
    expect(bulletOf(e)).toBeUndefined();
    e.destroy();
  });

  test('the picture bullet survives save and reopen', async () => {
    const e = await open(deck('', { image: false }));
    select(e);
    expect(
      e.exec({ type: 'setBulletFormat', format: { kind: 'picture', picture: png(8, 8) } }).ok
    ).toBe(true);
    const saved = await writePptx(e.session!.store.package);
    e.destroy();
    const again = await open(saved);
    expect(bulletOf(again)?.picture).toBeDefined();
    const back = readPptx(saved);
    expect(back.ok).toBe(true);
    again.destroy();
  });

  test('refuses bytes that are no PNG, JPEG, or GIF image', async () => {
    const e = await open(deck(''));
    select(e);
    const svg = strToU8('<svg xmlns="http://www.w3.org/2000/svg"/>');
    for (const picture of [new Uint8Array(0), svg, new Uint8Array([0x89, 0x50, 0x4e, 0x47])]) {
      const r = e.exec({ type: 'setBulletFormat', format: { kind: 'picture', picture } });
      expect(r.ok).toBe(false);
    }
    const bad = { kind: 'picture', picture: png(4, 4), color: '#FF0000' } as never;
    expect(e.exec({ type: 'setBulletFormat', format: bad }).ok).toBe(false);
    e.destroy();
  });

  test('the Bullets button turns a picture bullet off', async () => {
    const e = await open(deck(`<a:pPr marL="342900" indent="-342900">${BLIP}</a:pPr>`));
    select(e);
    expect(e.snapshot().format?.bullets).toBe(true);
    expect(e.exec({ type: 'toggleBullets' }).ok).toBe(true);
    expect(bulletOf(e)).toBeUndefined();
    e.destroy();
  });
});

/** A PNG with only its signature and header: a size, and no pixels to decode. */
function pngHeader(w: number, h: number): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(33);
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(out.buffer);
  view.setUint32(16, w);
  view.setUint32(20, h);
  out.set([8, 6, 0, 0, 0], 24);
  return out;
}

/** Every glyph run with a picture in a slide layout, at any depth (tables hold cell layouts). */
function picturesIn(value: unknown, out: GlyphRun[] = [], depth = 0): GlyphRun[] {
  if (!value || typeof value !== 'object' || depth > 12) return out;
  const v = value as Partial<GlyphRun>;
  if (v.kind === 'bullet' && v.picture) out.push(v as GlyphRun);
  for (const child of Object.values(value)) picturesIn(child, out, depth + 1);
  return out;
}

describe('picture bullet limits', () => {
  test('a picture over 25 million pixels is refused, since it would never draw', async () => {
    const e = await open(deck(''));
    select(e);
    const huge = { kind: 'picture', picture: pngHeader(6000, 6000) } as const;
    expect(e.exec({ type: 'setBulletFormat', format: huge }).ok).toBe(false);
    expect(bulletPictureProblem(pngHeader(6000, 6000))).toBe('size');
    expect(bulletPictureProblem(png(4, 4))).toBeNull();
    e.destroy();
  });

  test('the Picture field gives a reason for a file it refuses', async () => {
    const file = (bytes: Uint8Array<ArrayBuffer>, type: string) => new Blob([bytes], { type });
    expect(await readBulletPicture(null)).toBeNull();
    expect(await readBulletPicture(file(png(4, 4), 'image/svg+xml'))).toEqual({
      ok: false,
      reason: 'bulletsAndNumberingDialog.pictureType',
    });
    expect(await readBulletPicture(file(new Uint8Array(0), 'image/png'))).toEqual({
      ok: false,
      reason: 'bulletsAndNumberingDialog.pictureType',
    });
    expect(await readBulletPicture(file(pngHeader(6000, 6000), 'image/png'))).toEqual({
      ok: false,
      reason: 'bulletsAndNumberingDialog.pictureTooLarge',
    });
    // A browser may give no type: the bytes decide.
    expect((await readBulletPicture(file(png(4, 4), '')))?.ok).toBe(true);
  });

  test('a table cell takes a picture bullet from the default text style', async () => {
    const files = unzipSync(
      buildPptx({
        slides: [
          {
            shapes: table({
              id: 2,
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: IN,
              columns: [4 * IN],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="${IN}">${cell('cell')}</a:tr>`,
            }),
          },
        ],
        presentationRels: [['rIdPresBu', R_IMAGE, 'media/bullet.png']],
        media: [{ name: 'bullet.png', bytes: png(10, 10), contentType: 'image/png' }],
      })
    );
    const name = 'ppt/presentation.xml';
    const xml = strFromU8(files[name]!);
    const style = xml.indexOf('<p:defaultTextStyle>');
    const lvl1 = /<a:lvl1pPr[^>]*>/.exec(xml.slice(style))![0];
    files[name] = strToU8(
      xml.slice(0, style) +
        xml.slice(style).replace(lvl1, `${lvl1}<a:buBlip><a:blip r:embed="rIdPresBu"/></a:buBlip>`)
    );
    const e = await open(zipSync(files));
    const [bullet] = picturesIn(e.session!.layout(0));
    expect(bullet?.picture?.part).toBe('/ppt/presentation.xml');
    expect(bullet?.picture?.rId).toBe('rIdPresBu');
    e.destroy();
  });
});
