// Copy and paste of charts, diagrams, and embedded objects: each travels with its part graph,
// and a paste gives it new parts, in the same presentation or in another one. A link to a
// slide stays only in the presentation it came from. The payload is untrusted input.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  contentTypeOf,
  find,
  NS,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type NodeSpec,
  type PptxPackage,
} from '../../store/index.ts';
import type { ClipboardPayload, ClipboardRelationship } from '../../contracts/index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { freePartName, renumberShapes } from '../../automation/commands/package-edits.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(import.meta.dir, `../../../../../fixtures/${name}.pptx`)));

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const partOf = (e: PptxEditorInstance, i: number): string => e.session!.info.slides[i]!.partName;
const nodeOf = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;

/** Copy the shape named `name` on slide 1, through JSON, as a system clipboard holds it. */
async function copied(deck: string, name: string): Promise<ClipboardPayload> {
  const src = await open(fixture(deck));
  src.setSelection({ kind: 'shapes', nodeIds: [nodeOf(src, name)] });
  const payload = src.copy();
  expect(payload).not.toBeNull();
  return JSON.parse(JSON.stringify(payload)) as ClipboardPayload;
}

/** The target of relationship `id` of `source`. */
const targetOf = (pkg: PptxPackage, source: string, id: string | undefined): string | undefined =>
  relationshipsOf(pkg, source).find((r) => r.id === id)?.target;

/** Every internal relationship of the package names a part that exists. */
function expectResolved(pkg: PptxPackage): void {
  for (const name of pkg.parts.keys()) {
    const m = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(name);
    if (!m) continue;
    for (const r of relationshipsOf(pkg, `${m[1]}/${m[2]}`))
      if (!r.external) expect(pkg.parts.has(r.target!)).toBe(true);
  }
}

/** Save and read the deck again. */
function reread(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('the saved deck does not read');
  return r.pkg;
}

describe('a chart', () => {
  test('pastes as a new chart part on another slide of the same deck', async () => {
    const payload = await copied('charts', 'Chart 1');
    const e = await open(fixture('charts'));
    const original = targetOf(
      pkgOf(e),
      partOf(e, 0),
      attr(find(xmlPartOf(pkgOf(e), partOf(e, 0))!.root, 'c:chart'), 'r:id')
    );
    e.goToSlide(1);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const pkg = reread(e);
    const frames: Element[] = [];
    const walk = (el: Element): void => {
      if (el.local === 'chart') frames.push(el);
      for (const c of el.children) if (c.type === 'element') walk(c);
    };
    walk(xmlPartOf(pkg, partOf(e, 1))!.root);
    const pasted = targetOf(pkg, partOf(e, 1), attr(frames.at(-1), 'r:id'));
    expect(pasted).toBeDefined();
    expect(pasted).not.toBe(original);
    expect(contentTypeOf(pkg, pasted!)).toBe(
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    );
    // The new part holds the same chart.
    const bytes = (n: string) => {
      const p = pkg.parts.get(n)!;
      return p.kind === 'binary' ? Buffer.from(p.bytes).toString() : '';
    };
    expect(bytes(pasted!)).toBe(bytes(original!));
    expectResolved(pkg);
  });

  test('pastes into another deck with the parts that it owns', async () => {
    const payload = await copied('charts', 'Chart 1');
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
    const image: ClipboardRelationship = {
      id: 'rId7',
      type: `${R}image`,
      part: '/ppt/media/image9.png',
      contentType: 'image/png',
      base64: png,
    };
    const chart = { ...payload.relationships[0]!, relationships: [image] };
    const e = await open(fixture('slide-structure'));
    expect(e.exec({ type: 'paste', payload: { ...payload, relationships: [chart] } }).ok).toBe(
      true
    );
    const pkg = reread(e);
    // The deck has a chart of its own: the pasted one is the new part.
    const rel = relationshipsOf(pkg, partOf(e, 0))
      .filter((r) => r.type === `${R}chart`)
      .find((r) => r.target !== '/ppt/charts/chart1.xml')!;
    // The chart's own relationship keeps its id, and its picture is a media part of the deck.
    const own = relationshipsOf(pkg, rel.target!);
    expect(own.map((r) => r.id)).toEqual(['rId7']);
    expect(own[0]!.target!.startsWith('/ppt/media/')).toBe(true);
    expectResolved(pkg);
  });
});

describe('a diagram', () => {
  test('pastes into another deck with its data, layout, style, colors, and drawing', async () => {
    const payload = await copied('diagrams', 'Diagram 1');
    const e = await open(fixture('slide-structure'));
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const pkg = reread(e);
    const rels = relationshipsOf(pkg, partOf(e, 0));
    const of = (type: string) => rels.find((r) => r.type === type);
    for (const t of ['diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors'])
      expect(of(`${R}${t}`)).toBeDefined();
    const drawing = of('http://schemas.microsoft.com/office/2007/relationships/diagramDrawing')!;
    expect(pkg.parts.has(drawing.target!)).toBe(true);
    // The data part names the new slide relationship of the drawing.
    const data = pkg.parts.get(of(`${R}diagramData`)!.target!)!;
    const text = data.kind === 'binary' ? new TextDecoder().decode(data.bytes) : '';
    expect(text).toContain(`relId="${drawing.id}"`);
    expectResolved(pkg);
  });
});

describe('an embedded object', () => {
  test('pastes into another deck with its package and its picture', async () => {
    const payload = await copied('preserved-content', 'Embedded workbook');
    const e = await open(fixture('slide-structure'));
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const pkg = reread(e);
    const rels = relationshipsOf(pkg, partOf(e, 0));
    const workbook = rels.find((r) => r.type === `${R}package`)!;
    expect(workbook.target).toMatch(/^\/ppt\/embeddings\/Embedded_Worksheet\d+\.xlsx$/);
    expect(contentTypeOf(pkg, workbook.target!)).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    expect(rels.some((r) => r.type === `${R}image`)).toBe(true);
    expectResolved(pkg);
  });
});

describe('an untrusted payload', () => {
  test('is refused when a carried part claims another type', async () => {
    const payload = await copied('charts', 'Chart 1');
    const bad = {
      ...payload,
      relationships: [
        {
          ...payload.relationships[0]!,
          contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
        },
      ],
    };
    const e = await open(fixture('slide-structure'));
    expect(e.exec({ type: 'paste', payload: bad }).ok).toBe(false);
  });

  test('is refused when a carried part names a relationship that cannot be written', async () => {
    const payload = await copied('charts', 'Chart 1');
    for (const nested of [
      {
        id: 'r" Target="x',
        type: `${R}image`,
        part: '/a.png',
        contentType: 'image/png',
        base64: '',
      },
      {
        id: 'rId2',
        type: `${R}slide`,
        part: '/ppt/slides/slide1.xml',
        contentType: '',
        base64: '',
      },
      {
        id: 'rId2',
        type: `${R}hyperlink`,
        part: '',
        contentType: '',
        base64: '',
        target: 'javascript:x',
      },
    ]) {
      const chart = { ...payload.relationships[0]!, relationships: [nested] };
      const e = await open(fixture('slide-structure'));
      expect(e.exec({ type: 'paste', payload: { ...payload, relationships: [chart] } }).ok).toBe(
        false
      );
    }
  });
});

describe('a link to a slide', () => {
  const linked = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          shapes: shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(
              run('Go', '', '<a:hlinkClick r:id="rIdJump" action="ppaction://hlinksldjump"/>')
            ),
          }),
          rels: [['rIdJump', `${R}slide`, 'slide2.xml']],
        },
        { shapes: '' },
      ],
    });

  test('stays when pasted in the same deck', async () => {
    const e = await open(linked());
    e.setSelection({ kind: 'shapes', nodeIds: [nodeOf(e, 'T')] });
    const payload = e.copy()!;
    e.goToSlide(1);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const link = find(xmlPartOf(pkgOf(e), partOf(e, 1))!.root, 'a:hlinkClick');
    expect(targetOf(pkgOf(e), partOf(e, 1), attr(link, 'r:id'))).toBe(partOf(e, 1));
  });

  test('goes when pasted in another deck', async () => {
    const src = await open(linked());
    src.setSelection({ kind: 'shapes', nodeIds: [nodeOf(src, 'T')] });
    const payload = JSON.parse(JSON.stringify(src.copy())) as ClipboardPayload;
    const e = await open(linked());
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const shapes = xmlPartOf(pkgOf(e), partOf(e, 0))!.root;
    const links: Element[] = [];
    const walk = (el: Element): void => {
      if (el.local === 'hlinkClick') links.push(el);
      for (const c of el.children) if (c.type === 'element') walk(c);
    };
    walk(shapes);
    // Only the deck's own link is left.
    expect(links).toHaveLength(1);
  });
});

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

describe('the checks of an untrusted payload', () => {
  test('refuse an embedded object type that can hold macros', async () => {
    const payload = await copied('preserved-content', 'Embedded workbook');
    const relationships = payload.relationships.map((r) =>
      r.type === `${R}package` ? { ...r, contentType: 'application/vnd.ms-office.vbaProject' } : r
    );
    const e = await open(fixture('slide-structure'));
    expect(e.exec({ type: 'paste', payload: { ...payload, relationships } }).ok).toBe(false);
  });

  test('refuse a chart part that does not parse or has another root', async () => {
    const payload = await copied('charts', 'Chart 1');
    for (const xml of ['<c:chartSpace><not closed', '<p:sld xmlns:p="urn:x"/>']) {
      const chart = { ...payload.relationships[0]!, base64: Buffer.from(xml).toString('base64') };
      const e = await open(fixture('slide-structure'));
      expect(e.exec({ type: 'paste', payload: { ...payload, relationships: [chart] } }).ok).toBe(
        false
      );
    }
  });

  test('refuse an mc:AlternateContent with an undeclared prefix or a Fallback first', async () => {
    const payload = await copied('charts', 'Shape 2');
    const branch = (local: string, extra: object = {}) => ({
      ns: MC,
      local,
      children: [payload.shapes[0]!],
      ...extra,
    });
    const requires = { attrs: [{ ns: '', local: 'Requires', value: 'zz' }] };
    const declared = { ...requires, nsDecls: [{ prefix: 'zz', uri: 'urn:zz' }] };
    for (const children of [
      [branch('Choice', requires), branch('Fallback')],
      [branch('Fallback'), branch('Choice', declared)],
    ]) {
      const bad = { ...payload, shapes: [{ ns: MC, local: 'AlternateContent', children }] };
      const e = await open(fixture('slide-structure'));
      expect(e.exec({ type: 'paste', payload: bad as ClipboardPayload }).ok).toBe(false);
    }
  });
});

describe('new part names and shape ids', () => {
  test('a part name that a content-type override names is taken', async () => {
    const e = await open(fixture('slide-structure'));
    const batch = new OpBatch(pkgOf(e));
    const ct = xmlPartOf(batch.pkg, '/[Content_Types].xml')!;
    batch.push({
      op: 'insertNode',
      part: '/[Content_Types].xml',
      parent: ct.root.id,
      index: ct.root.children.length,
      spec: {
        ns: 'http://schemas.openxmlformats.org/package/2006/content-types',
        local: 'Override',
        attrs: [
          { ns: '', local: 'PartName', value: '/ppt/charts/extra1.xml' },
          { ns: '', local: 'ContentType', value: 'application/xml' },
        ],
      },
    });
    expect(freePartName(batch, '/ppt/charts/extra', '.xml')).toBe('/ppt/charts/extra2.xml');
  });

  test('copies in two branches share a new id, and two shapes in one branch do not', () => {
    const sp = (id: string) => ({
      ns: NS.p,
      local: 'sp',
      children: [
        {
          ns: NS.p,
          local: 'nvSpPr',
          children: [{ ns: NS.p, local: 'cNvPr', attrs: [{ ns: '', local: 'id', value: id }] }],
        },
      ],
    });
    const wrapper = {
      ns: MC,
      local: 'AlternateContent',
      children: [
        { ns: MC, local: 'Choice', children: [sp('9'), sp('9')] },
        { ns: MC, local: 'Fallback', children: [sp('9')] },
      ],
    };
    const ids: string[] = [];
    const walk = (n: NodeSpec): void => {
      if (n.local === 'cNvPr') ids.push(n.attrs![0]!.value);
      for (const c of n.children ?? []) if (!('text' in c)) walk(c);
    };
    walk(renumberShapes([wrapper], 4).specs[0]!);
    expect(ids).toEqual(['4', '5', '4']);
  });
});

describe('an mc:AlternateContent', () => {
  test('copies as its drawn copy alone when a branch needs a part that cannot travel', async () => {
    const box = (name: string, extra = '') =>
      shape({ id: 3, name, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(name)) }).replace(
        '<p:cNvPr ',
        `<p:cNvPr ${extra}`
      );
    const deck = buildPptx({
      slides: [
        {
          shapes:
            `<mc:AlternateContent xmlns:mc="${MC}">` +
            '<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">' +
            box('Ink', 'r:id="rIdInk" ') +
            `</mc:Choice><mc:Fallback>${box('Ink')}</mc:Fallback></mc:AlternateContent>`,
          rels: [['rIdInk', `${R}customXml`, '../customXml/item1.xml']],
        },
      ],
    });
    const e = await open(deck);
    e.setSelection({ kind: 'shapes', nodeIds: [nodeOf(e, 'Ink')] });
    expect(e.copy()!.shapes.map((s) => s.local)).toEqual(['sp']);
  });
});
