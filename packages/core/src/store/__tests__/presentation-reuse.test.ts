// The presentation structure is shared by packages that differ only in parts it does not
// read, and relationship ids resolve in one lookup. A shared structure must equal the one a
// fresh read gives, so each test compares it with a read of a copy that no cache knows.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  applyOps,
  changedPartNames,
  presentationInfo,
  readPptx,
  relationshipById,
  relationshipsOf,
  relsPartName,
  relTarget,
  repairedPresentationRoot,
  writePptx,
  xmlPartOf,
  type PptxOp,
  type PptxPackage,
} from '../index.ts';

function deck(n = 3): PptxPackage {
  const r = readPptx(
    buildPptx({
      slides: Array.from({ length: n }, (_, i) => ({
        shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100, text: para(run(`S${i}`)) }),
      })),
    })
  );
  if (!r.ok) throw new Error('deck');
  return r.pkg;
}

function apply(pkg: PptxPackage, ops: PptxOp[]): PptxPackage {
  const r = applyOps(pkg, ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  return r.pkg;
}

/** Set an attribute on the root of a part. */
function rootAttr(pkg: PptxPackage, part: string, local: string, value: string): PptxPackage {
  const root = xmlPartOf(pkg, part)!.root;
  return apply(pkg, [{ op: 'setAttribute', part, node: root.id, name: { ns: '', local }, value }]);
}

/** The info a fresh read gives: the same package, saved and read again. */
function freshInfo(pkg: PptxPackage) {
  const r = readPptx(writePptx(pkg));
  if (!r.ok) throw new Error('reopen');
  const strip = (p: PptxPackage) =>
    presentationInfo(p).slides.map(({ nodeId: _n, ...rest }) => rest);
  return strip(r.pkg);
}

const withoutNodeIds = (pkg: PptxPackage) =>
  presentationInfo(pkg).slides.map(({ nodeId: _n, ...rest }) => rest);

describe('presentationInfo across edits', () => {
  test('a slide content edit shares the info', () => {
    const pkg = deck();
    const info = presentationInfo(pkg);
    const slide = info.slides[1]!.partName;
    const next = rootAttr(pkg, slide, 'showMasterSp', '0');
    expect(presentationInfo(next)).toBe(info);
    expect(withoutNodeIds(next)).toEqual(freshInfo(next));
  });

  test('hiding a slide, or a relationship or presentation edit, reads it again', () => {
    const pkg = deck();
    const info = presentationInfo(pkg);
    const slide = info.slides[2]!.partName;
    const hidden = rootAttr(pkg, slide, 'show', '0');
    expect(presentationInfo(hidden)).not.toBe(info);
    expect(presentationInfo(hidden).slides[2]!.hidden).toBe(true);
    expect(withoutNodeIds(hidden)).toEqual(freshInfo(hidden));

    // The slide loses its layout relationship.
    const rels = relsPartName(slide);
    const rel = xmlPartOf(pkg, rels)!.root.children.find(
      (c) => c.type === 'element' && c.attrs.some((a) => a.value.endsWith('/slideLayout'))
    )!;
    const noLayout = apply(pkg, [{ op: 'removeNode', part: rels, node: rel.id }]);
    expect(presentationInfo(noLayout).slides[2]!.layoutPart).toBeUndefined();
    expect(withoutNodeIds(noLayout)).toEqual(freshInfo(noLayout));

    // The slide list loses a slide.
    const entry = info.slides[0]!;
    const shorter = apply(pkg, [
      { op: 'removeNode', part: info.presentationPart, node: entry.nodeId },
    ]);
    expect(presentationInfo(shorter).slides.map((s) => s.partName)).toEqual(
      info.slides.slice(1).map((s) => s.partName)
    );
  });

  test('an added or removed part reads it again', () => {
    const pkg = deck();
    const info = presentationInfo(pkg);
    const added = apply(pkg, [
      { op: 'addPart', part: '/ppt/extra.xml', xml: '<x xmlns="urn:x"/>' } as PptxOp,
    ]);
    expect(presentationInfo(added)).not.toBe(info);
    expect(presentationInfo(added)).toEqual(info);
  });

  test('the repaired presentation root is shared after a slide edit', () => {
    const pkg = deck();
    const root = repairedPresentationRoot(pkg);
    const slide = presentationInfo(pkg).slides[0]!.partName;
    expect(repairedPresentationRoot(rootAttr(pkg, slide, 'show', '0'))).toBe(root);
  });
});

/** The `p:sldId` entries of a presentation root. */
const slideEntries = (root: ReturnType<typeof repairedPresentationRoot>) =>
  (root?.children ?? [])
    .filter((c) => c.type === 'element' && c.local === 'sldIdLst')
    .flatMap((l) => (l.type === 'element' ? l.children : []))
    .filter((c) => c.type === 'element' && c.local === 'sldId');

describe('repairedPresentationRoot across edits', () => {
  test('a presentation relationship edit repairs again', () => {
    const pkg = deck();
    expect(slideEntries(repairedPresentationRoot(pkg))).toHaveLength(3);
    const info = presentationInfo(pkg);
    const rels = relsPartName(info.presentationPart);
    const rel = xmlPartOf(pkg, rels)!.root.children.find(
      (c) =>
        c.type === 'element' &&
        c.attrs.some((a) => a.local === 'Id' && a.value === info.slides[0]!.rId)
    )!;
    const next = apply(pkg, [{ op: 'removeNode', part: rels, node: rel.id }]);
    // The entry whose relationship is gone is not saved.
    expect(slideEntries(repairedPresentationRoot(next))).toHaveLength(2);
  });

  test('two documents edited in turn each keep their shared root', () => {
    const a = deck();
    const b = deck(2);
    const ra = repairedPresentationRoot(a);
    const rb = repairedPresentationRoot(b);
    const ia = presentationInfo(a);
    const ib = presentationInfo(b);
    const a2 = rootAttr(a, ia.slides[0]!.partName, 'showMasterSp', '0');
    const b2 = rootAttr(b, ib.slides[0]!.partName, 'showMasterSp', '0');
    expect(repairedPresentationRoot(a2)).toBe(ra);
    expect(repairedPresentationRoot(b2)).toBe(rb);
    expect(presentationInfo(a2)).toBe(ia);
    expect(presentationInfo(b2)).toBe(ib);
  });
});

describe('relationship lookups', () => {
  test('relationshipById finds the first relationship with an id', () => {
    const pkg = deck();
    const main = presentationInfo(pkg).presentationPart;
    for (const rel of relationshipsOf(pkg, main)) {
      expect(relationshipById(pkg, main, rel.id)).toBe(
        relationshipsOf(pkg, main).find((r) => r.id === rel.id)!
      );
      expect(relTarget(pkg, main, rel.id)).toBe(rel.external ? undefined : rel.target);
    }
    expect(relationshipById(pkg, main, 'rIdNone')).toBeUndefined();
    expect(relationshipById(pkg, '/ppt/none.xml', 'rId1')).toBeUndefined();
  });

  test('changedPartNames lists replaced, added, and removed parts', () => {
    const pkg = deck();
    expect(changedPartNames(pkg, pkg, 8)).toEqual([]);
    const slide = presentationInfo(pkg).slides[0]!.partName;
    const edited = rootAttr(pkg, slide, 'show', '0');
    expect(changedPartNames(pkg, edited, 8)).toEqual([slide]);
    const parts = new Map(pkg.parts);
    parts.delete(slide);
    const removed: PptxPackage = { parts };
    expect(changedPartNames(pkg, removed, 8)).toEqual([slide]);
    expect(changedPartNames(removed, pkg, 8)).toEqual([slide]);
    expect(changedPartNames(pkg, removed, 0)).toBeNull();
  });
});
