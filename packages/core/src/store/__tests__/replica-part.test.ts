import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  buildReplicaPart,
  PptxDocumentStore,
  readPptx,
  type ReplicaElementSpec,
} from '../index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

const spec = (over: Partial<ReplicaElementSpec> = {}): ReplicaElementSpec => ({
  key: 'root',
  ns: P,
  local: 'sld',
  prefix: 'p',
  nsDecls: [
    { prefix: 'p', uri: P },
    { prefix: 'a', uri: A },
  ],
  attrs: [],
  children: [
    {
      key: 'body',
      ns: A,
      local: 'p',
      prefix: 'a',
      children: [
        { key: 'r', ns: A, local: 'r', prefix: 'a', children: [{ key: 'txt', text: 'Hi' }] },
      ],
    },
  ],
  ...over,
});

const none = { idOf: () => undefined, nextId: 1 };

describe('buildReplicaPart', () => {
  test('builds a part with fresh ids and typed kinds', () => {
    const r = buildReplicaPart('/ppt/slides/slide9.xml', spec(), none);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ids.size).toBe(4);
    expect(new Set(r.ids.values()).size).toBe(4);
    expect(r.part.nextId).toBeGreaterThan(4);
  });

  test('keeps the ids this replica knows and never reuses them for new keys', () => {
    const known = new Map([
      ['root', 'n1'],
      ['txt', 'n7'],
    ]);
    const r = buildReplicaPart('/ppt/slides/slide9.xml', spec(), {
      idOf: (k) => known.get(k),
      nextId: 5,
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ids.get('root')).toBe('n1');
    expect(r.ids.get('txt')).toBe('n7');
    expect(new Set(r.ids.values()).size).toBe(4);
    expect(r.part.nextId).toBeGreaterThan(7);
  });

  test('refuses an invalid name, a repeated key, and an invalid text', () => {
    expect(buildReplicaPart('/x.xml', spec({ local: 'bad name' }), none).ok).toBe(false);
    const repeated = spec({
      children: [
        { key: 'dup', ns: A, local: 'p', prefix: 'a' },
        { key: 'dup', ns: A, local: 'p', prefix: 'a' },
      ],
    });
    expect(buildReplicaPart('/x.xml', repeated, none).ok).toBe(false);
    const badText = spec({ children: [{ key: 't', text: '\u0001' }] });
    expect(buildReplicaPart('/x.xml', badText, none).ok).toBe(false);
  });

  test('refuses a tree deeper than the XML reader accepts', () => {
    let deep: ReplicaElementSpec = { key: 'leaf', ns: A, local: 'p', prefix: 'a' };
    for (let i = 0; i < 300; i++)
      deep = { key: `k${i}`, ns: A, local: 'p', prefix: 'a', children: [deep] };
    expect(buildReplicaPart('/x.xml', spec({ children: [deep] }), none).ok).toBe(false);
  });
});

describe('PptxDocumentStore.replaceParts', () => {
  const store = () => {
    const r = readPptx(buildPptx({ slides: [{ shapes: '', layout: 'blank' }] }));
    if (!r.ok) throw new Error('fixture');
    return new PptxDocumentStore(r.pkg);
  };

  test('publishes one remote revision and clears local history', () => {
    const s = store();
    const built = buildReplicaPart('/ppt/slides/slide9.xml', spec(), none);
    if (!built.ok) throw new Error('build');
    const origins: string[] = [];
    s.subscribe((c) => origins.push(c.origin));
    const r = s.replaceParts([
      { kind: 'xml', partName: '/ppt/slides/slide9.xml', xml: built.part },
    ]);
    expect(r.ok).toBe(true);
    expect(origins).toEqual(['remote']);
    expect(s.canUndo).toBe(false);
    expect(s.package.parts.has('/ppt/slides/slide9.xml')).toBe(true);
  });

  test('refuses an unsafe part name and removing the content types part', () => {
    const s = store();
    const bytes = new Uint8Array([1]);
    expect(s.replaceParts([{ kind: 'binary', partName: '/../evil.bin', bytes }]).ok).toBe(false);
    expect(s.replaceParts([], ['/[Content_Types].xml']).ok).toBe(false);
  });
});
