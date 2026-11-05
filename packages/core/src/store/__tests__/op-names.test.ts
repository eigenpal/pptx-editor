// Regression tests for audit finding S3: ops validate every element and
// attribute name, prefix, namespace URI, and namespace declaration, and refuse each bad case
// with a typed rejection. The node specs below are the shapes a clipboard paste hands to
// `insertNode`; the store is the boundary that refuses them.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  find,
  NS,
  PptxDocumentStore,
  readPptx,
  serializeXmlPart,
  writePptx,
  xmlPartOf,
  type AttrName,
  type NodeSpec,
  type OpRejection,
  type PptxOp,
} from '../index.ts';

const SLIDE = '/ppt/slides/slide1.xml';

function store(): PptxDocumentStore {
  const r = readPptx(
    buildPptx({ slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }] })
  );
  if (!r.ok) throw new Error('fixture');
  return new PptxDocumentStore(r.pkg);
}

function refusal(s: PptxDocumentStore, op: PptxOp): OpRejection | null {
  const r = s.transact([op]);
  return r.ok ? null : r.rejection;
}

const root = (s: PptxDocumentStore) => xmlPartOf(s.package, SLIDE)!.root;
const spTree = (s: PptxDocumentStore) => find(root(s), 'p:spTree')!;

const setAttr = (s: PptxDocumentStore, name: AttrName): PptxOp => ({
  op: 'setAttribute',
  part: SLIDE,
  node: root(s).id,
  name,
  value: 'v',
});

const insert = (s: PptxDocumentStore, spec: unknown): PptxOp => ({
  op: 'insertNode',
  part: SLIDE,
  parent: spTree(s).id,
  index: 0,
  spec: spec as NodeSpec,
});

describe('S3: names are XML names', () => {
  test('a local name carrying markup is refused, not written verbatim', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: '', local: 'x="1"><p:evil/><y z' }))).toMatchObject({
      code: 'invalid-name',
    });
    expect(refusal(s, insert(s, { ns: 'urn:x', local: 'e f' }))).toMatchObject({
      code: 'invalid-name',
    });
  });

  test('a prefix carrying markup is refused on elements, attributes, and declarations', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: 'urn:x', local: 'a', prefix: 'q="1" r' }))).toMatchObject({
      code: 'invalid-name',
    });
    expect(refusal(s, insert(s, { ns: 'urn:x', local: 'e', prefix: 'a b="c"' }))).toMatchObject({
      code: 'invalid-name',
    });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'z="1" w', uri: 'u' }] }))
    ).toMatchObject({ code: 'invalid-name' });
  });

  test('non-ASCII XML names are accepted', () => {
    const s = store();
    expect(refusal(s, insert(s, { ns: 'urn:x', local: 'Größe', prefix: 'ü' }))).toBeNull();
    expect(() => writePptx(s.package)).not.toThrow();
  });
});

describe('S3: reserved prefixes and namespaces', () => {
  test('an unprefixed xmlns attribute is refused (it would rebind the default namespace)', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: '', local: 'xmlns' }))).toMatchObject({
      code: 'reserved-name',
    });
    expect(
      refusal(
        s,
        insert(s, { ns: '', local: 'e', attrs: [{ ns: '', local: 'xmlns', value: NS.p }] })
      )
    ).toMatchObject({ code: 'reserved-name' });
  });

  test('the xmlns prefix and namespace are refused (the hijack of DrawingML)', () => {
    const s = store();
    const hijack = { ns: 'urn:y', prefix: 'xmlns', local: 'a', value: 'urn:hijack' };
    expect(refusal(s, insert(s, { ns: NS.p, local: 'sp', attrs: [hijack] }))).toMatchObject({
      code: 'reserved-name',
    });
    expect(refusal(s, setAttr(s, { ns: NS.xmlns, local: 'a' }))).toMatchObject({
      code: 'reserved-name',
    });
    expect(refusal(s, insert(s, { ns: 'urn:x', local: 'e', prefix: 'xmlns' }))).toMatchObject({
      code: 'reserved-name',
    });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'xmlns', uri: 'u' }] }))
    ).toMatchObject({ code: 'reserved-name' });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'p', uri: NS.xmlns }] }))
    ).toMatchObject({ code: 'reserved-name' });
  });

  test('the xml prefix is accepted only for the XML namespace', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: 'urn:x', local: 'lang', prefix: 'xml' }))).toMatchObject({
      code: 'reserved-name',
    });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'xml', uri: 'urn:x' }] }))
    ).toMatchObject({ code: 'reserved-name' });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'q', uri: NS.xml }] }))
    ).toMatchObject({ code: 'reserved-name' });
    expect(refusal(s, setAttr(s, { ns: NS.xml, local: 'space', prefix: 'xml' }))).toBeNull();
  });
});

describe('S3: namespace URIs and declarations', () => {
  test('a namespace URI that is not XML text is refused (it made every save fail)', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: 'urn:\u0000', local: 'a' }))).toMatchObject({
      code: 'invalid-namespace',
    });
    expect(refusal(s, insert(s, { ns: 'urn:￿', local: 'e' }))).toMatchObject({
      code: 'invalid-namespace',
    });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'q', uri: '\u0001' }] }))
    ).toMatchObject({ code: 'invalid-namespace' });
    expect(() => writePptx(s.package)).not.toThrow();
  });

  test('bindings XML 1.0 cannot express are refused', () => {
    const s = store();
    expect(refusal(s, setAttr(s, { ns: '', local: 'a', prefix: 'p' }))).toMatchObject({
      code: 'invalid-namespace',
    });
    expect(
      refusal(s, insert(s, { ns: 'urn:x', local: 'e', nsDecls: [{ prefix: 'q', uri: '' }] }))
    ).toMatchObject({ code: 'invalid-namespace' });
    expect(
      refusal(s, insert(s, { ns: '', local: 'e', nsDecls: [{ prefix: '', uri: 'urn:d' }] }))
    ).toMatchObject({ code: 'invalid-namespace' });
  });

  test('one attribute or prefix twice on an element is refused', () => {
    const s = store();
    expect(
      refusal(
        s,
        insert(s, {
          ns: 'urn:x',
          local: 'e',
          attrs: [
            { ns: 'urn:y', local: 'a', prefix: 'q', value: '1' },
            { ns: 'urn:y', local: 'a', prefix: 'r', value: '2' },
          ],
        })
      )
    ).toMatchObject({ code: 'duplicate-name' });
    expect(
      refusal(
        s,
        insert(s, {
          ns: 'urn:x',
          local: 'e',
          nsDecls: [
            { prefix: 'q', uri: 'urn:1' },
            { prefix: 'q', uri: 'urn:2' },
          ],
        })
      )
    ).toMatchObject({ code: 'duplicate-name' });
  });

  test('names deep in a spec are checked too, and a refused op changes nothing', () => {
    const s = store();
    const before = s.package;
    const deep = { ns: NS.p, local: 'sp', children: [{ ns: NS.p, local: 'x', prefix: 'xmlns' }] };
    expect(refusal(s, insert(s, deep))).toMatchObject({ code: 'reserved-name' });
    expect(s.package).toBe(before);
    expect(s.revision).toBe(0);
  });
});

describe('S3: what is accepted saves and reopens with the same names', () => {
  test('a declared prefix that shadows DrawingML cannot capture the shape children', () => {
    const s = store();
    const spec: NodeSpec = {
      ns: NS.p,
      local: 'sp',
      prefix: 'p',
      nsDecls: [{ prefix: 'a', uri: 'urn:hijack' }],
      children: [
        { ns: NS.a, local: 'off', prefix: 'a', attrs: [{ ns: '', local: 'x', value: '1' }] },
      ],
    };
    expect(refusal(s, insert(s, spec))).toBeNull();
    const reopened = readPptx(writePptx(s.package));
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    const sp = find(find(xmlPartOf(reopened.pkg, SLIDE)!.root, 'p:spTree')!, 'p:sp')!;
    expect(sp.children[0].type === 'element' && sp.children[0].ns).toBe(NS.a);
    expect(serializeXmlPart(xmlPartOf(s.package, SLIDE)!)).toContain('xmlns:a="urn:hijack"');
  });
});
