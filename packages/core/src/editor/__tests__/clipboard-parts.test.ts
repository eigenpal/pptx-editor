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
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import type { ClipboardPayload, ClipboardRelationship } from '../../contracts/index.ts';
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
