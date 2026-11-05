// An edit of the copy that the editor draws in `mc:AlternateContent` leaves no out-of-date copy
// in another branch (`automation/commands/alternate-copies.ts`, ECMA-376 Part 3, 10.2.1). A reader that reads
// only `mc:Fallback` then never shows the object as it was before the edit.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../../scripts/fixtures/timing.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { withoutStaleCopies } from '../../automation/commands/alternate-copies.ts';
import { findMatches } from '../../automation/commands/find.ts';

const IN = 914_400;
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const VML = 'urn:schemas-microsoft-com:vml';

const fixture = (): Uint8Array =>
  new Uint8Array(
    readFileSync(join(import.meta.dir, '../../../../../fixtures/preserve-unsupported.pptx'))
  );

const box = (id: number, color: string, text: string, x = IN): string =>
  shape({
    id,
    name: `Box ${id}`,
    x,
    y: IN,
    cx: 3 * IN,
    cy: IN,
    fill: solid(color),
    text: para(run(text)),
  });

/** A deck whose only slide holds `body` in its shape tree, with optional timing. */
const deck = (body: string, ext?: string): Uint8Array =>
  buildPptx({ slides: [{ layout: 'blank', shapes: body, ...(ext ? { ext } : {}) }] });

const wrap = (branches: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}">${branches}</mc:AlternateContent>`;
const choice = (requires: string, decl: string, inner: string): string =>
  `<mc:Choice ${decl} Requires="${requires}">${inner}</mc:Choice>`;
const fallback = (inner: string): string => `<mc:Fallback>${inner}</mc:Fallback>`;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const slideRoot = (pkg: PptxPackage): Element => xmlPartOf(pkg, '/ppt/slides/slide1.xml')!.root;
const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, '/ppt/slides/slide1.xml')!);
const drawnId = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;

/** Save, reopen, and return the slide root of the saved file. */
function saved(e: PptxEditorInstance): Element {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen failed');
  return slideRoot(r.pkg);
}

/**
 * The text a reader that understands no `mc:Choice` shows: each `mc:Fallback` in place of its
 * `mc:AlternateContent`, and the other shapes as they are.
 */
function fallbackReaderText(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    if (el.ns === MC && el.local === 'Choice') return;
    if (el.local === 't' && el.ns === 'http://schemas.openxmlformats.org/drawingml/2006/main') {
      const t = el.children[0];
      out.push(t && t.type === 'text' ? t.value : '');
      return;
    }
    for (const c of elements(el)) walk(c);
  };
  walk(find(root, 'p:spTree')!);
  return out;
}

/** The `a:off/@x` values in the `mc:Fallback` of a slide root. */
const fallbackOffsets = (root: Element): string[] => {
  const fb = find(root, 'mc:Fallback');
  const out: string[] = [];
  const walk = (el: Element): void => {
    if (el.local === 'off') out.push(attr(el, 'x') ?? '');
    for (const c of elements(el)) walk(c);
  };
  if (fb) walk(fb);
  return out;
};

describe('an edit of a drawn mc:Choice copy', () => {
  test('a move leaves no stale copy for a reader of mc:Fallback, after save and reopen', async () => {
    const e = await open(fixture());
    const id = drawnId(e, 'Shape 3');
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    const root = saved(e);
    // The Choice holds the moved shape.
    const moved = find(find(root, 'mc:Choice'), 'a:off');
    expect(attr(moved, 'x')).toBe(String(5_486_400 + IN));
    // A Fallback reader no longer sees the shape at its old place, nor its old text.
    expect(fallbackOffsets(root)).toEqual([]);
    expect(fallbackReaderText(root)).not.toContain('Fallback branch');
  });

  test('typing removes the copy, and undo brings it back', async () => {
    const e = await open(fixture());
    const before = canonicalPackageFingerprint(e.session!.store.package);
    const id = drawnId(e, 'Shape 3');
    const at = { paragraph: 0, offset: 0 };
    e.setSelection({ kind: 'text', nodeId: id, anchor: at, focus: at });
    for (const ch of 'New ') expect(e.input('text', ch).ok).toBe(true);
    expect(slideXml(e)).toContain('New Choice branch');
    expect(slideXml(e)).not.toContain('Fallback branch');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(before);
  });

  test('a rename changes both copies, so both stay', async () => {
    const e = await open(fixture());
    const id = drawnId(e, 'Shape 3');
    expect(e.exec({ type: 'renameShape', nodeId: id, name: 'Kept' }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml.match(/name="Kept"/g)).toHaveLength(2);
    expect(xml).toContain('Fallback branch');
  });

  test('every other branch loses its copy, and the shapes beside it stay', async () => {
    const e = await open(
      deck(
        wrap(
          choice('v', `xmlns:v="${VML}"`, box(3, '000000', 'vml copy')) +
            choice(
              'p14',
              `xmlns:p14="${P14}"`,
              box(3, 'C00000', 'drawn') + box(4, '00B050', 'p14 other', 5 * IN)
            ) +
            fallback(box(3, '70AD47', 'fallback copy') + box(4, '00B050', 'fallback other', 5 * IN))
        )
      )
    );
    e.setSelection({ kind: 'shapes', nodeIds: [drawnId(e, 'Box 3')] });
    expect(e.exec({ type: 'setShapeFill', color: '0000FF' }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).not.toContain('vml copy');
    expect(xml).not.toContain('fallback copy');
    expect(xml).toContain('drawn');
    expect(xml).toContain('fallback other');
    // The `v` Choice keeps no content, so a reader that understands `v` falls to nothing for it.
    expect(xml).toContain(`<mc:Choice xmlns:v="${VML}" Requires="v"/>`);
  });

  test('a delete removes every copy, the empty wrapper, and the animation of the shape', async () => {
    const body =
      box(2, '4472C4', 'plain', 5 * IN) +
      wrap(
        choice('p14', `xmlns:p14="${P14}"`, box(3, 'C00000', 'drawn')) +
          fallback(box(3, '70AD47', 'fallback copy'))
      );
    const e = await open(deck(body, clickEntrances([2, 3])));
    e.setSelection({ kind: 'shapes', nodeIds: [drawnId(e, 'Box 3')] });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    const root = saved(e);
    expect(find(root, 'mc:AlternateContent')).toBeUndefined();
    expect(fallbackReaderText(root)).toEqual(['plain']);
    const xml = serializeXmlPart(xmlPartOf(e.session!.store.package, '/ppt/slides/slide1.xml')!);
    expect(xml).not.toContain('spid="3"');
    expect(xml).toContain('spid="2"');
  });

  test('an edit inside a drawn group removes the group copy in mc:Fallback', async () => {
    const group = (id: number, inner: string) =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      '<p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/>' +
      '<a:chOff x="914400" y="914400"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>' +
      `${inner}</p:grpSp>`;
    const e = await open(
      deck(
        wrap(
          choice('p14', `xmlns:p14="${P14}"`, group(5, box(3, 'C00000', 'inner drawn'))) +
            fallback(group(5, box(3, '70AD47', 'inner fallback')))
        )
      )
    );
    const g = e.snapshot().shapes!.find((s) => s.name === 'Group 5')!;
    const inner = g.children!.find((s) => s.name === 'Box 3')!.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [inner] });
    expect(e.exec({ type: 'setShapeFill', color: '0000FF' }).ok).toBe(true);
    expect(slideXml(e)).not.toContain('inner fallback');
    expect(slideXml(e)).toContain('<mc:Fallback/>');
  });
});

describe('an edit of a drawn mc:Fallback copy', () => {
  const vmlDeck = () =>
    deck(
      wrap(
        choice('v', `xmlns:v="${VML}"`, box(3, 'C00000', 'choice copy')) +
          fallback(box(3, '70AD47', 'drawn fallback'))
      )
    );

  test('a move keeps the mc:Choice copy, as the fidelity page lists', async () => {
    const e = await open(vmlDeck());
    e.setSelection({ kind: 'shapes', nodeIds: [drawnId(e, 'Box 3')] });
    expect(e.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    expect(slideXml(e)).toContain('choice copy');
  });

  test('a delete removes the mc:Choice copy too', async () => {
    const e = await open(vmlDeck());
    e.setSelection({ kind: 'shapes', nodeIds: [drawnId(e, 'Box 3')] });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    expect(find(saved(e), 'mc:AlternateContent')).toBeUndefined();
  });
});

describe('withoutStaleCopies', () => {
  test('returns the same ops when no op touches a branch', () => {
    const r = readPptx(deck(box(2, '4472C4', 'plain')));
    if (!r.ok) throw new Error('open failed');
    const sp = find(find(slideRoot(r.pkg), 'p:spTree'), 'p:sp')!;
    const ops = [
      {
        op: 'setAttribute' as const,
        part: '/ppt/slides/slide1.xml',
        node: child(child(sp, 'p:nvSpPr'), 'p:cNvPr')!.id,
        name: { ns: '', local: 'name' },
        value: 'x',
      },
    ];
    expect(withoutStaleCopies(r.pkg, ops)).toBe(ops);
  });
});

describe('find in mc:AlternateContent', () => {
  test('finds the text of the drawn branch, and only that', async () => {
    const e = await open(fixture());
    const pkg = e.session!.store.package;
    const hits = findMatches(pkg, 'branch');
    expect(hits.map((h) => h.text)).toEqual(['branch']);
    expect(hits[0].nodeId).toBe(drawnId(e, 'Shape 3'));
    expect(findMatches(pkg, 'Fallback')).toEqual([]);
  });
});
