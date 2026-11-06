// The kit of the feature edits (`feature-edits.ts`): the recipe types, the helpers that the
// recipes use, and the runner. A recipe opens a fixture of its row, makes the row's own edit
// through the public editor API, saves, and opens the saved copy again. The edit must be in the
// reopened copy, its check must fail on the unedited deck, a second save must give the same
// package, and only the parts the edit owns may change.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { EditorCommand } from '../../packages/core/src/contracts/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../../packages/core/src/editor/index.ts';
import {
  attr,
  canonicalForm,
  canonicalPackageFingerprint,
  elements,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';

export const FIXTURES = join(import.meta.dir, '../../fixtures');

/** A 2 × 2 PNG, for the picture edits. */
export const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

/** What a recipe uses to edit the open deck. */
export interface EditTools {
  readonly editor: PptxEditorInstance;
  /** Run a command; a refusal fails the recipe. */
  run(command: EditorCommand): void;
  /** The node id of the shape named `name` on the current slide, at any depth. */
  node(name: string): string;
  /** Select the shapes named `names` on the current slide. */
  select(...names: string[]): void;
  /** Select text in the shape named `name`, in paragraph `paragraph`, from `from` to `to`. */
  selectText(name: string, from: number, to: number, paragraph?: number): void;
  /** Select text in a table cell. */
  selectCell(table: string, row: number, col: number, to?: number): void;
  /** Type text at the selection. */
  type(text: string): void;
  /** A small PNG picture. */
  readonly png: Uint8Array;
}

/** What a recipe uses to check the reopened copy. */
export interface SavedDeck {
  readonly pkg: PptxPackage;
  readonly editor: PptxEditorInstance;
  /** The XML of slide `index` (0-based), with prefixes, as saved. */
  slideXml(index: number): string;
  /** The XML of the shape named `name` on slide `index`, as saved. */
  shapeXml(index: number, name: string): string;
  /** The XML of any part, as saved. */
  partXml(part: string): string;
}

/** The edit of one feature row. */
export interface FeatureEdit {
  /** A fixture of the row (`pptx-feature-evidence.ts`). */
  readonly fixture: string;
  /** The slide the edit runs on, 0-based. */
  readonly slide?: number;
  readonly edit: (t: EditTools) => void;
  /** A failure message, or undefined when the reopened copy has the edit. */
  readonly check: (d: SavedDeck) => string | undefined;
  /**
   * Parts the edit may change besides the slide it runs on and that slide's relationships, as
   * patterns of part names. Absent: only that slide.
   */
  readonly parts?: readonly RegExp[];
}

export const SHAPE_KINDS = new Set(['sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame']);

/** Every shape element under `el`, at any depth. */
export function shapesIn(el: Element): Element[] {
  const out: Element[] = [];
  for (const c of elements(el)) {
    if (SHAPE_KINDS.has(c.local)) out.push(c);
    out.push(...shapesIn(c));
  }
  return out;
}

export const nameOf = (shape: Element): string | undefined => {
  const nv = elements(shape).find((c) => c.local.startsWith('nv'));
  const cNvPr = nv && elements(nv).find((c) => c.local === 'cNvPr');
  return attr(cNvPr, 'name');
};

/** The XML of `el` with the namespace prefixes of the part it is in. */
const xmlOf = (el: Element, root: Element = el): string =>
  serializeXmlPart({
    partName: '',
    nextId: 0,
    root: { ...el, nsDecls: [...root.nsDecls, ...el.nsDecls] },
  });

export function slidePart(pkg: PptxPackage, index: number): string {
  const ref = presentationInfo(pkg).slides[index];
  if (!ref) throw new Error(`no slide ${index + 1}`);
  return ref.partName;
}

export function shapeElement(pkg: PptxPackage, index: number, name: string): Element {
  const root = xmlPartOf(pkg, slidePart(pkg, index))?.root;
  const hit = root && shapesIn(root).find((s) => nameOf(s) === name);
  if (!hit) throw new Error(`no shape "${name}" on slide ${index + 1}`);
  return hit;
}

/** The XML of the shape named `name` on slide `index`. */
export function shapeXmlOf(pkg: PptxPackage, index: number, name: string): string {
  return xmlOf(shapeElement(pkg, index, name), xmlPartOf(pkg, slidePart(pkg, index))!.root);
}

export const has = (xml: string, ...parts: string[]): string | undefined => {
  const missing = parts.filter((p) => !xml.includes(p));
  return missing.length ? `missing ${missing.map((m) => JSON.stringify(m)).join(', ')}` : undefined;
};

export const lacks = (xml: string, ...parts: string[]): string | undefined => {
  const present = parts.filter((p) => xml.includes(p));
  return present.length
    ? `still has ${present.map((m) => JSON.stringify(m)).join(', ')}`
    : undefined;
};

/** The runs of the first paragraph of a shape, in order: their `a:rPr` and text. */
export function runsOf(shapeXml: string): { readonly rPr: string; readonly text: string }[] {
  const p = shapeXml.split(/<a:p[ >]/)[1] ?? '';
  return [
    ...p.matchAll(/<a:r>(<a:rPr[^>]*?(?:\/>|>.*?<\/a:rPr>))?<a:t>([^<]*)<\/a:t><\/a:r>/g),
  ].map((m) => ({ rPr: m[1] ?? '', text: m[2] }));
}

/** The `a:rPr` of the first run of the first paragraph. */
export const firstRun = (shapeXml: string): string => runsOf(shapeXml)[0]?.rPr ?? '';

/** The first failure of several checks. */
export const all = (...results: (string | undefined)[]): string | undefined =>
  results.find(Boolean);

/** Parts that every slide edit may touch when it adds a picture. */
export const MEDIA = [/^\/ppt\/media\//, /^\/\[Content_Types\]\.xml$/];
export const PRESENTATION = [
  /^\/ppt\/presentation\.xml$/,
  /^\/ppt\/_rels\/presentation\.xml\.rels$/,
];
export const NEW_SLIDES = [
  /^\/ppt\/slides\//,
  /^\/ppt\/notesSlides\//,
  /^\/\[Content_Types\]\.xml$/,
  /^\/docProps\/app\.xml$/,
];

/** The result of one recipe. */
export interface FeatureEditResult {
  readonly id: string;
  readonly ok: boolean;
  readonly failure?: string;
  /** The edited copy, as saved. */
  readonly bytes?: Uint8Array;
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const result = await editor.load(bytes);
  if (!result.ok) throw new Error(`open failed: ${JSON.stringify(result)}`);
  return editor;
}

function tools(editor: PptxEditorInstance): EditTools {
  const node = (name: string): string => {
    type Item = { nodeId: string; name: string; children?: readonly Item[] };
    const walk = (items: readonly Item[]): Item | undefined => {
      for (const i of items) {
        if (i.name === name) return i;
        const hit = i.children && walk(i.children);
        if (hit) return hit;
      }
      return undefined;
    };
    const hit = walk(editor.snapshot().shapes as readonly Item[]);
    if (!hit) throw new Error(`no shape "${name}" on the current slide`);
    return hit.nodeId;
  };
  const run = (command: EditorCommand): void => {
    const result = editor.exec(command);
    if (!result.ok) throw new Error(`${command.type} refused: ${JSON.stringify(result)}`);
  };
  return {
    editor,
    run,
    node,
    select: (...names) => editor.setSelection({ kind: 'shapes', nodeIds: names.map(node) }),
    selectText: (name, from, to, paragraph = 0) =>
      editor.setSelection({
        kind: 'text',
        nodeId: node(name),
        anchor: { paragraph, offset: from },
        focus: { paragraph, offset: to },
      }),
    selectCell: (table, row, col, to = 0) =>
      editor.setSelection({
        kind: 'text',
        nodeId: node(table),
        cell: { row, column: col },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: to },
      }),
    type: (text) => {
      const result = editor.input('text', text);
      if (!result.ok) throw new Error(`typing refused: ${JSON.stringify(result)}`);
    },
    png: PNG,
  };
}

function saved(pkg: PptxPackage, editor: PptxEditorInstance): SavedDeck {
  const partXml = (part: string): string => {
    const xml = xmlPartOf(pkg, part);
    if (!xml) throw new Error(`no part ${part}`);
    return serializeXmlPart(xml);
  };
  return {
    pkg,
    editor,
    partXml,
    slideXml: (index) => partXml(slidePart(pkg, index)),
    shapeXml: (index, name) => shapeXmlOf(pkg, index, name),
  };
}

/** The content of a part, for comparing two packages that were read separately. */
function partContent(pkg: PptxPackage, name: string): string | undefined {
  const part = pkg.parts.get(name);
  if (!part) return undefined;
  if (part.kind === 'xml') return canonicalForm(part.xml.root);
  return Buffer.from(part.bytes).toString('base64');
}

/** The names of the parts whose content differs, or that only one package has. */
function differentParts(a: PptxPackage, b: PptxPackage): string[] {
  const names = new Set([...a.parts.keys(), ...b.parts.keys()]);
  return [...names].filter((n) => partContent(a, n) !== partContent(b, n)).sort();
}

/** Run one recipe: edit, save, reopen, check, save again. */
export async function runFeatureEdit(id: string, recipe: FeatureEdit): Promise<FeatureEditResult> {
  const fixture = new Uint8Array(readFileSync(join(FIXTURES, `${recipe.fixture}.pptx`)));
  const slide = recipe.slide ?? 0;
  const editor = await open(fixture);
  let reopened: PptxEditorInstance | undefined;
  try {
    const untouched = editor.save();
    if (!untouched) return { id, ok: false, failure: 'the untouched deck did not save' };
    editor.goToSlide(slide);
    recipe.edit(tools(editor));
    const bytes = editor.save();
    if (!bytes) return { id, ok: false, failure: 'the edited deck did not save' };
    const before = readPptx(untouched);
    const after = readPptx(bytes);
    if (!before.ok || !after.ok) return { id, ok: false, failure: 'a saved copy does not read' };
    // A check that the unedited deck passes proves nothing about the edit.
    const baseline = await open(untouched);
    try {
      baseline.goToSlide(slide);
      let vacuous: boolean;
      try {
        vacuous = recipe.check(saved(before.pkg, baseline)) === undefined;
      } catch {
        vacuous = false;
      }
      if (vacuous) return { id, ok: false, failure: 'the check passes on the unedited deck too' };
    } finally {
      baseline.destroy();
    }
    reopened = await open(bytes);
    reopened.goToSlide(slide);
    const failure = recipe.check(saved(after.pkg, reopened));
    if (failure) return { id, ok: false, failure, bytes };
    const again = reopened.save();
    const second = again && readPptx(again);
    if (!second || !second.ok)
      return { id, ok: false, failure: 'the reopened copy did not save', bytes };
    if (canonicalPackageFingerprint(second.pkg) !== canonicalPackageFingerprint(after.pkg))
      return {
        id,
        ok: false,
        failure: `a second save changed ${differentParts(after.pkg, second.pkg).join(', ')}`,
        bytes,
      };
    const own = slidePart(before.pkg, slide);
    const allowed = [
      (p: string) => p === own,
      (p: string) => p === own.replace(/\/([^/]+)$/, '/_rels/$1.rels'),
      ...(recipe.parts ?? []).map((re) => (p: string) => re.test(p)),
    ];
    const collateral = differentParts(before.pkg, after.pkg).filter(
      (p) => !allowed.some((ok) => ok(p))
    );
    if (collateral.length)
      return { id, ok: false, failure: `the edit also changed ${collateral.join(', ')}`, bytes };
    return { id, ok: true, bytes };
  } catch (error) {
    return { id, ok: false, failure: error instanceof Error ? error.message : String(error) };
  } finally {
    editor.destroy();
    reopened?.destroy();
  }
}
