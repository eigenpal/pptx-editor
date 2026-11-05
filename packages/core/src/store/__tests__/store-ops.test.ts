import { describe, expect, test } from 'bun:test';
import { buildPptx, shape, para, run } from '../../../../../scripts/fixtures/builder.ts';
import {
  find,
  nodeById,
  NS,
  PPTX_OP_KINDS,
  PptxDocumentStore,
  readPptx,
  semanticDigest,
  textOf,
  xmlPartOf,
  type PptxOp,
} from '../index.ts';

const SLIDE = '/ppt/slides/slide1.xml';

function store(): PptxDocumentStore {
  const r = readPptx(
    buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100, text: para(run('abc')) }) }],
    })
  );
  if (!r.ok) throw new Error('fixture');
  return new PptxDocumentStore(r.pkg);
}

const slideRoot = (s: PptxDocumentStore) => xmlPartOf(s.package, SLIDE)!.root;

describe('ops', () => {
  test('setAttribute moves a shape and is undoable', () => {
    const s = store();
    const off = find(slideRoot(s), 'a:off')!;
    const r = s.transact([
      { op: 'setAttribute', part: SLIDE, node: off.id, name: { ns: '', local: 'x' }, value: '500' },
    ]);
    expect(r.ok).toBe(true);
    expect(find(slideRoot(s), 'a:off')!.attrs.find((a) => a.local === 'x')!.value).toBe('500');
    expect(s.canUndo).toBe(true);
    s.undo();
    expect(find(slideRoot(s), 'a:off')!.attrs.find((a) => a.local === 'x')!.value).toBe('0');
    s.redo();
    expect(find(slideRoot(s), 'a:off')!.attrs.find((a) => a.local === 'x')!.value).toBe('500');
  });

  test('setText replaces run text and keeps node ids of untouched parts', () => {
    const s = store();
    const t = find(slideRoot(s), 'a:t')!;
    const offBefore = find(slideRoot(s), 'a:off')!;
    s.transact([{ op: 'setText', part: SLIDE, node: t.id, text: 'xyz' }]);
    expect(textOf(find(slideRoot(s), 'a:t'))).toBe('xyz');
    expect(find(slideRoot(s), 'a:off')).toBe(offBefore);
  });

  test('a refused op leaves the store unchanged (atomic transaction)', () => {
    const s = store();
    const before = s.package;
    const t = find(slideRoot(s), 'a:t')!;
    const r = s.transact([
      { op: 'setText', part: SLIDE, node: t.id, text: 'changed' },
      { op: 'removeNode', part: SLIDE, node: 'n99999' },
    ]);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.rejection.code).toBe('unknown-node');
    expect(s.package).toBe(before);
    expect(s.revision).toBe(0);
  });

  test('invalid XML text is refused, never written', () => {
    const s = store();
    const t = find(slideRoot(s), 'a:t')!;
    const r = s.transact([{ op: 'setText', part: SLIDE, node: t.id, text: 'bad\u0001' }]);
    expect(r.ok).toBe(false);
  });

  test('insert, move, and remove nodes', () => {
    const s = store();
    const tree = find(slideRoot(s), 'p:spTree')!;
    const r = s.transact([
      {
        op: 'insertNode',
        part: SLIDE,
        parent: tree.id,
        index: tree.children.length,
        spec: {
          ns: NS.p,
          local: 'sp',
          prefix: 'p',
          children: [{ ns: NS.p, local: 'nvSpPr', prefix: 'p' }],
        },
      },
    ]);
    expect(r.ok).toBe(true);
    const tree2 = find(slideRoot(s), 'p:spTree')!;
    const added = tree2.children[tree2.children.length - 1];
    expect(added.type === 'element' && added.kind).toBe('sp');
    expect(
      s.transact([{ op: 'moveNode', part: SLIDE, node: added.id, parent: tree2.id, index: 2 }]).ok
    ).toBe(true);
    expect(
      s.transact([{ op: 'moveNode', part: SLIDE, node: tree2.id, parent: added.id, index: 0 }]).ok
    ).toBe(false);
    expect(s.transact([{ op: 'removeNode', part: SLIDE, node: added.id }]).ok).toBe(true);
    expect(nodeById(xmlPartOf(s.package, SLIDE)!, added.id)).toBeUndefined();
  });

  test('addPart refuses unsafe names and duplicates', () => {
    const s = store();
    expect(s.transact([{ op: 'addPart', part: '/../x.xml', xml: '<x/>' }]).ok).toBe(false);
    expect(s.transact([{ op: 'addPart', part: SLIDE, xml: '<x/>' }]).ok).toBe(false);
    // A part needs a content type: the same transaction declares one for its extension.
    const types = xmlPartOf(s.package, '/[Content_Types].xml')!.root.id;
    const binDefault: PptxOp = {
      op: 'insertNode',
      part: '/[Content_Types].xml',
      parent: types,
      index: 0,
      spec: {
        ns: NS.ct,
        local: 'Default',
        attrs: [
          { ns: '', local: 'Extension', value: 'bin' },
          { ns: '', local: 'ContentType', value: 'application/octet-stream' },
        ],
      },
    };
    expect(
      s.transact([binDefault, { op: 'addPart', part: '/ppt/media/a.bin', base64: 'AAEC' }]).ok
    ).toBe(true);
    expect(s.transact([{ op: 'removePart', part: '/ppt/media/a.bin' }]).ok).toBe(true);
  });

  test('typing merges into one history entry by history group', () => {
    const s = store();
    const t = find(slideRoot(s), 'a:t')!;
    const historyGroup = Symbol('typing');
    for (const text of ['a', 'ab', 'abc'])
      s.transact([{ op: 'setText', part: SLIDE, node: t.id, text }], { historyGroup });
    s.undo();
    expect(textOf(find(slideRoot(s), 'a:t'))).toBe('abc');
    expect(s.canUndo).toBe(false);
  });

  test('ops are plain JSON and every kind is listed', () => {
    const op: PptxOp = { op: 'setText', part: SLIDE, node: 'n1', text: 'x' };
    expect(JSON.parse(JSON.stringify(op))).toEqual(op);
    expect(PPTX_OP_KINDS).toHaveLength(7);
  });

  test('an edited package still round-trips through save', () => {
    const s = store();
    const t = find(slideRoot(s), 'a:t')!;
    s.transact([{ op: 'setText', part: SLIDE, node: t.id, text: 'saved & <escaped>' }]);
    expect(() => semanticDigest(s.package)).not.toThrow();
  });
});
