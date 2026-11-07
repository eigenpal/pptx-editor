// Structural edits of kept content: the save evidence that a feature row with a `preserved`
// round trip keeps its content when the slide or the object moves, not only when an edit is
// next to it (`feature-edits.ts`). Each recipe duplicates the slide, pastes the kept objects on
// a new slide of the same deck, pastes them into another deck, or groups them, then saves and
// reopens. Its check finds each kept object with the same XML (relationship and drawing ids
// aside) and relationship targets with the same bytes, and every relationship of the copy names
// a part that exists.
//
// `kept-content-edits.test.ts` runs every recipe. `bun scripts/fidelity/kept-content-edits.ts
// --out <dir>` writes each edited copy as `<id>.pptx`, to check for a repair prompt
// (CONTRIBUTING.md, "Fidelity harness").

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { createPptxEditor } from '../../packages/core/src/editor/index.ts';
import {
  NS,
  parseXmlPart,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';
import {
  added,
  all,
  FIXTURES,
  runFeatureEdit,
  shapeElement,
  shapeXmlOf,
  slidePart,
  type EditTools,
  type FeatureEdit,
  type PartRule,
  type SavedDeck,
} from './feature-edit-kit.ts';

/** The kept content of one feature row: a fixture slide and the objects that hold it. */
interface Kept {
  readonly fixture: string;
  readonly slide?: number;
  /** The kept objects, by name. None: the slide itself holds the kept content. */
  readonly names?: readonly string[];
  /** A shape of the same slide to group the kept objects with, when they are one object. */
  readonly partner?: string;
}

const KEPT: Readonly<Record<string, Kept>> = {
  'shapes.custom-geometry': { fixture: 'custom-geometry', names: ['Freeform'] },
  'effects.fill-overlay': { fixture: 'fill-fidelity', slide: 2, names: ['Shape 5', 'Shape 6'] },
  'effects.3d-and-other': { fixture: 'preserved-content', names: ['3-D shape'], partner: 'Title' },
  'pictures.emf-wmf': { fixture: 'preserved-content', names: ['EMF picture', 'WMF picture'] },
  'objects.charts': { fixture: 'charts', names: ['Chart 1'], partner: 'Shape 2' },
  'objects.diagrams': { fixture: 'diagrams', names: ['Diagram 1'], partner: 'Shape 2' },
  'objects.media-ole': {
    fixture: 'preserved-content',
    names: ['Audio clip', 'Video clip', 'Embedded workbook'],
  },
  'objects.equations': { fixture: 'preserved-content', names: ['Equation'] },
  'review.comments': { fixture: 'slide-structure' },
  'themes.placeholder-inheritance': { fixture: 'placeholders-background' },
};

/** Parts that a structural edit may change or add. */
const PARTS: readonly PartRule[] = [
  /^\/ppt\/slides\//,
  /^\/ppt\/presentation\.xml$/,
  /^\/ppt\/_rels\/presentation\.xml\.rels$/,
  /^\/\[Content_Types\]\.xml$/,
  /^\/docProps\/app\.xml$/,
  /^\/ppt\/commentAuthors\.xml$/,
  added(/^\/ppt\/(charts|diagrams|embeddings|media|comments|notesSlides|drawings)\//),
];

/** XML with the ids that a copy renumbers left out: relationship ids and drawing ids. */
const normalized = (xml: string): string =>
  xml
    .replace(/\br:(\w+)="[^"]*"/g, 'r:$1=""')
    .replace(/(<p:cNvPr [^>]*?)\bid="\d+"/g, '$1id=""')
    .replace(/\brelId="[^"]*"/g, 'relId=""');

/** The bytes of a part as a save writes it, as text, or undefined. */
function contentOf(pkg: PptxPackage, part: string | undefined): string | undefined {
  const p = part ? pkg.parts.get(part) : undefined;
  if (!p) return undefined;
  if (p.kind === 'xml') return normalized(serializeXmlPart(p.xml));
  // A part that the reader keeps as bytes (a chart, a diagram, comments) may still be XML: it
  // compares as the writer serializes its tree, so a paste that writes it again compares equal.
  const tree = parseXmlPart(p.partName, new TextDecoder().decode(p.bytes));
  return tree.ok
    ? normalized(serializeXmlPart(tree.part))
    : Buffer.from(p.bytes).toString('base64');
}

/** The targets that the `r:*` attributes of `el` name, in document order, as their content. */
function targetsOf(pkg: PptxPackage, source: string, el: Element): (string | undefined)[] {
  const rels = new Map(relationshipsOf(pkg, source).map((r) => [r.id, r]));
  const out: (string | undefined)[] = [];
  const walk = (node: Element): void => {
    for (const a of node.attrs)
      if (a.ns === NS.r && a.value !== '') {
        const r = rels.get(a.value);
        out.push(r?.external ? r.externalTarget : contentOf(pkg, r?.target));
      }
    for (const c of node.children) if (c.type === 'element') walk(c);
  };
  walk(el);
  return out;
}

/** A failure when the kept object `name` of slide `index` differs from the fixture's copy. */
function sameObject(d: SavedDeck, index: number, name: string, from: number): string | undefined {
  const xml = normalized(shapeXmlOf(d.pkg, index, name));
  if (xml !== normalized(shapeXmlOf(d.original, from, name))) return `"${name}" changed`;
  const now = targetsOf(d.pkg, slidePart(d.pkg, index), shapeElement(d.pkg, index, name));
  const was = targetsOf(
    d.original,
    slidePart(d.original, from),
    shapeElement(d.original, from, name)
  );
  if (now.some((t) => t === undefined)) return `a relationship of "${name}" names no part`;
  return now.join('\n') === was.join('\n') ? undefined : `a part that "${name}" names changed`;
}

/** A failure when a relationship of the package names a part that does not exist. */
function resolved(pkg: PptxPackage): string | undefined {
  for (const name of pkg.parts.keys()) {
    const m = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(name);
    if (!m) continue;
    const source = `${m[1]}/${m[2]}`;
    for (const r of relationshipsOf(pkg, source))
      if (!r.external && !pkg.parts.has(r.target!)) return `${source} names a missing ${r.target}`;
  }
  return undefined;
}

/** The slide children that hold slide-level kept content, as XML without ids. */
function slideContent(pkg: PptxPackage, index: number): string {
  const root = xmlPartOf(pkg, slidePart(pkg, index))!.root;
  return normalized(serializeXmlPart({ partName: '', nextId: 0, root })).replace(
    /(<p14:creationId\b[^>]*\bval=")\d+"/,
    '$1"'
  );
}

/** The comment list that slide `index` names, without its indexes. */
function commentsOf(pkg: PptxPackage, index: number): string | undefined {
  const rel = relationshipsOf(pkg, slidePart(pkg, index)).find((r) => r.type.endsWith('/comments'));
  return contentOf(pkg, rel?.target)?.replace(/\bidx="\d+"/g, 'idx=""');
}

/** The check that the kept content of `k` is on slide `index` as on the fixture slide. */
function keptOn(k: Kept, index: number): (d: SavedDeck) => string | undefined {
  const from = k.slide ?? 0;
  return (d) =>
    all(
      ...(k.names ?? []).map((n) => sameObject(d, index, n, from)),
      k.names
        ? undefined
        : slideContent(d.pkg, index) === slideContent(d.original, from)
          ? undefined
          : 'the slide changed',
      commentsOf(d.original, from) === undefined ||
        commentsOf(d.pkg, index) === commentsOf(d.original, from)
        ? undefined
        : 'the comments changed',
      resolved(d.pkg)
    );
}

/** Copy the kept objects of `k` in `t`'s editor (or from a fresh copy of the fixture). */
async function payloadOf(k: Kept, t: EditTools, fresh: boolean) {
  if (!fresh) {
    t.select(...k.names!);
    return t.editor.copy();
  }
  const src = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await src.load(new Uint8Array(readFileSync(join(FIXTURES, `${k.fixture}.pptx`))));
  if (!r.ok) throw new Error('the source deck did not open');
  src.goToSlide(k.slide ?? 0);
  const ids = k.names!.map((n) => src.snapshot().shapes!.find((s) => s.name === n)!.nodeId);
  src.setSelection({ kind: 'shapes', nodeIds: ids });
  const payload = JSON.parse(JSON.stringify(src.copy())) as ReturnType<typeof src.copy>;
  src.destroy();
  return payload;
}

/** The recipes of one row. */
function recipesOf(id: string, k: Kept): [string, FeatureEdit][] {
  const slide = k.slide ?? 0;
  const base = { fixture: k.fixture, slide, parts: PARTS };
  const out: [string, FeatureEdit][] = [
    [
      `${id}: duplicate slide`,
      {
        ...base,
        edit: (t) => t.run({ type: 'duplicateSlide', index: slide }),
        check: keptOn(k, slide + 1),
      },
    ],
  ];
  if (!k.names) return out;
  for (const fresh of [false, true])
    out.push([
      `${id}: paste ${fresh ? 'into another deck' : 'on a new slide'}`,
      {
        ...base,
        edit: async (t) => {
          const payload = await payloadOf(k, t, fresh);
          if (!payload) throw new Error('nothing was copied');
          t.run({ type: 'addSlide' });
          t.editor.goToSlide(slide + 1);
          t.run({ type: 'paste', payload });
        },
        check: keptOn(k, slide + 1),
      },
    ]);
  if (k.names.length > 1 || k.partner)
    out.push([
      `${id}: group`,
      {
        ...base,
        edit: (t) => {
          t.select(...k.names!, ...(k.partner ? [k.partner] : []));
          t.run({ type: 'group' });
        },
        check: (d) =>
          all(
            serializeXmlPart(xmlPartOf(d.pkg, slidePart(d.pkg, slide))!).includes('<p:grpSp>')
              ? undefined
              : 'no group',
            ...k.names!.map((n) => sameObject(d, slide, n, slide)),
            resolved(d.pkg)
          ),
      },
    ]);
  return out;
}

/** The recipes, by `<row id>: <edit>`. */
export const KEPT_CONTENT_EDITS: Readonly<Record<string, FeatureEdit>> = Object.fromEntries(
  Object.entries(KEPT).flatMap(([id, k]) => recipesOf(id, k))
);

/** The rows that the recipes cover. */
export const KEPT_CONTENT_ROWS: readonly string[] = Object.keys(KEPT);

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  for (const [id, recipe] of Object.entries(KEPT_CONTENT_EDITS)) {
    const result = await runFeatureEdit(id, recipe);
    if (!result.ok) failed++;
    console.log(`${id}: ${result.ok ? `ok ${result.fingerprint}` : `FAIL ${result.failure}`}`);
    const file = id.replace(/[^a-z0-9.-]+/gi, '-');
    if (values.out && result.bytes) writeFileSync(join(values.out, `${file}.pptx`), result.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
