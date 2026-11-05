// Tests for the two XML readers at the trust boundary: DTD and entity rejection, the depth and
// element caps, attribute normalization, and that the one-pass reader from DOCX Editor
// (`xml-fast-read.ts` at fc28dba9) builds exactly the tree the bounded reader builds, or gives
// the part to it.

import { describe, expect, test } from 'bun:test';
import { fastReadXml } from '../xml-fast-read.ts';
import {
  MAX_XML_DEPTH,
  XML_DEFAULT_MAX_ELEMENTS,
  readXml,
  readXmlBounded,
  type XmlNode,
  type XmlResult,
} from '../xml-reader.ts';
import { decodeXmlEntities, normalizeAttributeValue } from '../xml-text.ts';

const READERS: readonly [string, (xml: string) => XmlResult][] = [
  ['readXml', (xml) => readXml(xml)],
  ['readXmlBounded', (xml) => readXmlBounded(xml)],
];

type Element = Extract<XmlNode, { type: 'element' }>;

describe.each(READERS)('%s refuses DTDs and entities', (_, read) => {
  test.each([
    ['<!DOCTYPE r><r/>', 'dtd-forbidden'],
    ['<?xml version="1.0"?><!DOCTYPE r [<!ENTITY e "x">]><r>&e;</r>', 'dtd-forbidden'],
    ['<!ENTITY e "x"><r/>', 'entity-forbidden'],
    ['<r>&e;</r>', 'entity-forbidden'],
    ['<r a="&e;"/>', 'entity-forbidden'],
    ['<!ELEMENT r ANY><r/>', 'dtd-forbidden'],
    ['<!ATTLIST r a CDATA "x"><r/>', 'dtd-forbidden'],
    ['<!NOTATION n SYSTEM "x"><r/>', 'dtd-forbidden'],
    ['<r a="<?"><!DOCTYPE x [<!ENTITY e "BOOM">]><y>&e;</y><z b="?>"/></r>', 'dtd-forbidden'],
    ['<r a="<!--"/><!DOCTYPE r><r b="-->">x</r>', 'dtd-forbidden'],
    ["<r a='<![CDATA['/><!DOCTYPE r><q b=']]>'>x</q>", 'dtd-forbidden'],
  ])('%s', (xml, reason) => {
    expect(read(xml)).toMatchObject({ ok: false, reason });
  });

  test('the predefined entities and character references still read', () => {
    const r = read('<r a="&lt;&amp;&#x41;">&gt;&quot;&apos;&#65;</r>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const el = r.nodes[0] as Element;
    expect(el.attributes.a).toBe('<&A');
    expect(el.children).toEqual([{ type: 'text', value: '>"\'A' }]);
  });

  test('literal markup inside CDATA, comments, and PIs is not markup', () => {
    expect(read('<r><![CDATA[<!DOCTYPE x>&e;]]><!-- <!ENTITY --><?pi <!DOCTYPE?></r>').ok).toBe(
      true
    );
  });

  test('a numeric reference to a character XML 1.0 forbids is refused', () => {
    for (const xml of ['<r>&#0;</r>', '<r>&#xFFFE;</r>', '<r a="&#xD800;"/>'])
      expect(read(xml)).toMatchObject({ ok: false });
  });
});

describe.each(READERS)('%s caps depth at %s levels', (_, read) => {
  const nested = (n: number): string => '<a>'.repeat(n) + '</a>'.repeat(n);

  test('depth 256 opens and depth 257 is refused', () => {
    expect(MAX_XML_DEPTH).toBe(256);
    expect(read(nested(256)).ok).toBe(true);
    expect(read(nested(257))).toMatchObject({ ok: false, reason: 'too-deep' });
  });
});

describe('element and size budgets', () => {
  test('an element budget refuses before the tree is built', () => {
    const xml = '<r>' + '<a/>'.repeat(10) + '</r>';
    expect(readXml(xml, { maxBytes: 1e6, maxElements: 5 })).toMatchObject({
      ok: false,
      reason: 'too-many-elements',
    });
    expect(readXml(xml, { maxBytes: 1e6, maxElements: 11 }).ok).toBe(true);
  });

  test('a byte budget counts UTF-8 bytes', () => {
    const xml = '<r>' + 'é'.repeat(10) + '</r>';
    expect(readXml(xml, { maxBytes: xml.length })).toMatchObject({
      ok: false,
      reason: 'too-large',
    });
  });

  test('a limit that is not a finite, non-negative integer refuses', () => {
    for (const maxBytes of [Number.NaN, Infinity, -1, 1.5])
      expect(readXml('<r/>', { maxBytes })).toMatchObject({ ok: false, reason: 'invalid-limits' });
    expect(XML_DEFAULT_MAX_ELEMENTS).toBeGreaterThan(0);
  });
});

describe.each(READERS)('%s normalizes attribute values (XML 1.0 3.3.3)', (_, read) => {
  test('literal tab, line feed, and carriage return become spaces; references are kept', () => {
    const r = read('<r a="x\ny\tz" b="1\r\n2" c="p&#xA;q&#x9;r&#xD;s"/>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const el = r.nodes[0] as Element;
    expect(el.attributes.a).toBe('x y z');
    expect(el.attributes.b).toBe('1 2');
    expect(el.attributes.c).toBe('p\nq\tr\rs');
  });

  test('attribute records have no prototype, so a file name cannot reach one', () => {
    const r = read('<r __proto__="x" constructor="y" toString="z"/>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const el = r.nodes[0] as Element;
    expect(Object.getPrototypeOf(el.attributes)).toBe(null);
    const own = (name: string): unknown =>
      Object.getOwnPropertyDescriptor(el.attributes, name)?.value;
    expect([own('__proto__'), own('constructor'), own('toString')]).toEqual(['x', 'y', 'z']);
  });
});

describe('the one-pass reader', () => {
  const PARTS = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><a:t xml:space="preserve"> a &amp; b </a:t></p:sp></p:spTree></p:cSld></p:sld>',
    '<r a="1" b=\'2\'><x c="3"/><x c="3"/><x c="4"/></r>',
    '<r a="x\ny\tz"/>',
    '<r>line\r\nend</r>',
    '<r><a/>  text  <b>&#x1F600;</b></r>',
  ];

  test.each(PARTS)('builds the tree the bounded reader builds: %s', (xml) => {
    const fast = fastReadXml(xml, XML_DEFAULT_MAX_ELEMENTS);
    expect(fast).not.toBeNull();
    const bounded = readXmlBounded(xml);
    // JSON text keeps attribute order, which `toEqual` ignores; a save writes that order.
    expect(bounded.ok && JSON.stringify(bounded.nodes)).toBe(JSON.stringify(fast));
  });

  test.each([
    '<!DOCTYPE r><r/>',
    '<r><!-- c --></r>',
    '<r><![CDATA[x]]></r>',
    '<r>&e;</r>',
    '<r a="<"/>',
    '<r></s>',
    '<r a="1" a="2"/>',
    '<é/>',
    '<r/><r/>',
    '<__proto__/>',
    '<a>'.repeat(91) + '</a>'.repeat(91),
  ])('gives up and leaves the decision to the bounded reader: %s', (xml) => {
    expect(fastReadXml(xml, XML_DEFAULT_MAX_ELEMENTS)).toBeNull();
  });

  test('never builds more elements than the budget', () => {
    // It counts every tag, end tags too, before it builds anything: four here.
    expect(fastReadXml('<r><a/><a/></r>', 3)).toBeNull();
    expect(fastReadXml('<r><a/><a/></r>', 4)).not.toBeNull();
  });

  test('shared attribute records are frozen, so one element cannot change another', () => {
    const fast = fastReadXml('<r><x c="3"/><x c="3"/></r>', 10)!;
    const [first, second] = (fast[0] as Element).children as Element[];
    expect(Object.isFrozen(first!.attributes)).toBe(true);
    expect(second!.attributes).toEqual(first!.attributes);
  });
});

describe('xml-text helpers', () => {
  test('decodeXmlEntities decodes only the predefined entities and character references', () => {
    expect(decodeXmlEntities('&amp;&lt;&gt;&quot;&apos;&#65;&#x42;&e;')).toBe('&<>"\'AB&e;');
    expect(() => decodeXmlEntities('&#x0;')).toThrow();
  });

  test('normalizeAttributeValue turns each literal tab, line feed, and carriage return into a space', () => {
    expect(normalizeAttributeValue('a\tb\nc\rd\r\ne')).toBe('a b c d  e');
  });
});

// Ported from docx-editor/packages/core/src/store/__tests__/xml-fast-read.test.ts at fc28dba9
// (EigenPal), with a PresentationML part as the base.
describe('mutations of a part', () => {
  /** The tree with attribute order spelled out, so an order change is a difference. */
  function shape(nodes: readonly XmlNode[] | null): unknown {
    if (nodes === null) return null;
    return nodes.map((node) =>
      node.type === 'text'
        ? ['text', node.value]
        : [node.name, Object.entries(node.attributes), shape(node.children)]
    );
  }

  /** The fast read, held to the bounded read: same tree when accepted. */
  function agrees(xml: string): readonly XmlNode[] | null {
    const fast = fastReadXml(xml, 1_000_000);
    if (fast !== null) {
      const bounded = readXmlBounded(xml);
      expect(bounded.ok).toBe(true);
      if (bounded.ok) expect(shape(fast)).toEqual(shape(bounded.nodes));
    }
    return fast;
  }

  // A deterministic generator, so a failure reproduces.
  let seed = 0x2545f491;
  const random = (n: number): number => {
    seed ^= seed << 13;
    seed ^= seed >>> 17;
    seed ^= seed << 5;
    return (seed >>> 0) % n;
  };
  const base =
    '<?xml version="1.0" standalone="yes"?>\n<p:sld xmlns:p="u" xmlns:a="v">' +
    '<p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr wrap="square"/>' +
    '<a:p><a:pPr algn="ctr"/><a:r><a:t xml:space="preserve"> a &amp; b </a:t></a:r></a:p>' +
    "<a:p><a:r><a:rPr lang='en-US'/><a:t>c'd</a:t></a:r></a:p>" +
    '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
  const pieces = ['<', '>', '/', '"', "'", '=', '&', ';', ' ', '\r', '!', '?', 'a:', 'x', '#'];

  test('never accepts what the bounded reader refuses, and agrees on what it accepts', () => {
    let accepted = 0;
    for (let round = 0; round < 4000; round += 1) {
      let xml = base;
      for (let edits = 1 + random(3); edits > 0; edits -= 1) {
        const at = random(xml.length);
        const kind = random(3);
        if (kind === 0) xml = xml.slice(0, at) + xml.slice(at + 1);
        else if (kind === 1) xml = xml.slice(0, at) + pieces[random(pieces.length)] + xml.slice(at);
        else xml = xml.slice(0, at) + pieces[random(pieces.length)] + xml.slice(at + 1);
      }
      if (agrees(xml) !== null) accepted += 1;
    }
    // Both outcomes occur, so the comparison is not vacuous.
    expect(accepted).toBeGreaterThan(100);
    expect(accepted).toBeLessThan(4000);
  }, 30_000);
});
