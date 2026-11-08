// Bounded XML reader tests: DTD/entity rejection, order/attribute/whitespace preservation, and
// no value coercion.
//
// The first two blocks are adapted from
// docx-editor/packages/core/src/store/__tests__/xml-reader.test.ts at d0bac831c (EigenPal,
// Apache-2.0), with WordprocessingML names replaced by DrawingML ones. The rest are regression
// tests for audit findings S4, S5, K2, and K6.

import { describe, expect, test } from 'bun:test';
import {
  readXml,
  MAX_XML_DEPTH,
  XML_DEFAULT_MAX_ELEMENTS,
  findElement,
  childElements,
  textContent,
  requireXmlStringScalar,
} from '@eigenpal/ooxml-core';
import { parseXmlPart } from '../package/tree.ts';
import { serializeXmlPart } from '../package/serialize.ts';

describe('trust boundary rejections', () => {
  test('refuses DTDs, entity declarations, and custom entity refs', () => {
    expect(readXml('<!DOCTYPE x><x/>')).toMatchObject({ ok: false, reason: 'dtd-forbidden' });
    expect(readXml('<!ENTITY lol "z"><x/>')).toMatchObject({
      ok: false,
      reason: 'entity-forbidden',
    });
    // Billion-laughs style reference to a custom entity.
    expect(readXml('<x>&lol;</x>')).toMatchObject({ ok: false, reason: 'entity-forbidden' });
    // The five predefined entities are allowed.
    expect(readXml('<x>a &amp; b</x>').ok).toBe(true);
  });
  test('enforces a size bound', () => {
    expect(readXml('<x/>', { maxBytes: 2 })).toMatchObject({ ok: false, reason: 'too-large' });
  });
  test('enforces an element-count bound', () => {
    expect(readXml('<x><a/><b/></x>', { maxBytes: 100, maxElements: 3 }).ok).toBe(true);
    expect(readXml('<x><a/><b/></x>', { maxBytes: 100, maxElements: 2 })).toMatchObject({
      ok: false,
      reason: 'too-many-elements',
    });
  });
  test('honors an element budget above the default before parser allocation', () => {
    // Intentionally malformed: distinguish the element preflight from XML validation
    // without allocating millions of parsed nodes in the test process.
    const xml = '<x>' + '<a/>'.repeat(XML_DEFAULT_MAX_ELEMENTS) + '</x><unclosed';
    expect(readXml(xml)).toMatchObject({ ok: false, reason: 'too-many-elements' });
    expect(
      readXml(xml, { maxBytes: xml.length, maxElements: XML_DEFAULT_MAX_ELEMENTS + 2 })
    ).toMatchObject({ ok: false, reason: 'parse-error' });
    // Two reads of a 40 MB part, each scanned by the one-pass reader first and then by the
    // bounded reader; on a loaded machine that takes longer than the default 5 s.
  }, 30_000);
  test('counts elements before parser allocation and before full XML validation', () => {
    expect(readXml('<x><a/><b/><unclosed', { maxBytes: 100, maxElements: 2 })).toMatchObject({
      ok: false,
      reason: 'too-many-elements',
    });
  });
  test('rejects NaN and non-finite configured limits', () => {
    expect(readXml('<x/>', { maxBytes: Number.NaN })).toMatchObject({
      ok: false,
      reason: 'invalid-limits',
    });
    expect(readXml('<x/>', { maxBytes: Number.POSITIVE_INFINITY })).toMatchObject({
      ok: false,
      reason: 'invalid-limits',
    });
    expect(readXml('<x/>', { maxBytes: 100, maxElements: Number.NaN })).toMatchObject({
      ok: false,
      reason: 'invalid-limits',
    });
    expect(readXml('<x/>', { maxBytes: 100, maxElements: Number.NEGATIVE_INFINITY })).toMatchObject(
      { ok: false, reason: 'invalid-limits' }
    );
  });

  test('rejects decoded numeric references forbidden by XML 1.0', () => {
    for (const reference of [
      '&#0;',
      '&#1;',
      '&#x8;',
      '&#xB;',
      '&#xC;',
      '&#xD800;',
      '&#xFFFE;',
      '&#xFFFF;',
      '&#x110000;',
      '&#999999999999999999999999;',
    ]) {
      expect(readXml(`<x>${reference}</x>`)).toMatchObject({
        ok: false,
        reason: 'parse-error',
      });
      expect(readXml(`<x a="${reference}"/>`)).toMatchObject({
        ok: false,
        reason: 'parse-error',
      });
    }
  });

  test('rejects literal XML 1.0-invalid characters in every value context', () => {
    for (const invalid of ['\u0000', '￾', '￿', '\uD800', '\uDC00']) {
      expect(readXml(`<x>${invalid}</x>`)).toMatchObject({
        ok: false,
        reason: 'parse-error',
      });
      expect(readXml(`<x a="${invalid}"/>`)).toMatchObject({
        ok: false,
        reason: 'parse-error',
      });
      expect(readXml(`<x><![CDATA[${invalid}]]></x>`)).toMatchObject({
        ok: false,
        reason: 'parse-error',
      });
    }
  });

  test('measures maxBytes as UTF-8 bytes rather than UTF-16 code units', () => {
    const xml = '<x>é</x>';
    expect(xml.length).toBe(8);
    expect(readXml(xml, { maxBytes: 8 })).toMatchObject({
      ok: false,
      reason: 'too-large',
    });
    expect(readXml(xml, { maxBytes: 9 }).ok).toBe(true);
  });
});

describe('fidelity preservation', () => {
  test('preserves significant child order', () => {
    const r = readXml('<p><a/><b/><a/></p>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const p = findElement(r.nodes, 'p')!;
    expect(
      p.children.filter((c) => c.type === 'element').map((c) => (c as { name: string }).name)
    ).toEqual(['a', 'b', 'a']);
  });

  test('preserves attributes and raw lexical values without coercion', () => {
    const r = readXml('<a:t xml:space="preserve" n="007" b="true">0042</a:t>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const t = findElement(r.nodes, 'a:t')!;
    // Zero-padded / boolean-looking attribute values stay strings, verbatim.
    expect(t.attributes).toEqual({ 'xml:space': 'preserve', n: '007', b: 'true' });
    expect(Object.getPrototypeOf(t.attributes)).toBeNull();
    expect(textContent(t)).toBe('0042'); // not coerced to number 42
  });

  test('preserves significant whitespace in text', () => {
    const r = readXml('<a:t>  spaced  </a:t>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(textContent(findElement(r.nodes, 'a:t')!)).toBe('  spaced  ');
  });

  test('reads a DrawingML paragraph structure', () => {
    const xml =
      '<p:txBody><a:p><a:r><a:t>Hello</a:t></a:r><a:r><a:t> world</a:t></a:r></a:p></p:txBody>';
    const r = readXml(xml);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const body = findElement(r.nodes, 'p:txBody')!;
    const paras = childElements(body, 'a:p');
    expect(paras).toHaveLength(1);
    expect(textContent(paras[0])).toBe('Hello world');
  });

  test('preserves valid supplementary numeric references', () => {
    const result = readXml('<x>&#x1F600;</x>');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(textContent(findElement(result.nodes, 'x')!)).toBe('😀');
  });

  test('treats numeric-reference spelling inside CDATA as literal text', () => {
    const result = readXml('<x><![CDATA[&#0; &lol; <!DOCTYPE x>]]></x>');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(textContent(findElement(result.nodes, 'x')!)).toBe('&#0; &lol; <!DOCTYPE x>');
    expect(readXml('<x>&#0;</x>')).toMatchObject({
      ok: false,
      reason: 'parse-error',
    });
  });

  test('requireXmlStringScalar fails closed on non-string text/CDATA/attribute values', () => {
    expect(requireXmlStringScalar('ok', 'text')).toBe('ok');
    expect(() => requireXmlStringScalar({} as unknown, 'text')).toThrow(/non-scalar text/);
    expect(() => requireXmlStringScalar(42 as unknown, 'cdata')).toThrow(/non-scalar cdata/);
    expect(() => requireXmlStringScalar(['x'] as unknown, 'attribute')).toThrow(
      /non-scalar attribute/
    );
  });
});

describe('S4: one depth cap of 256 levels', () => {
  const nested = (n: number): string => '<a>'.repeat(n) + '</a>'.repeat(n);
  const nestedSelfClosing = (n: number): string =>
    '<a>'.repeat(n - 1) + '<a/>' + '</a>'.repeat(n - 1);

  test('depth 256 opens and depth 257 is refused', () => {
    expect(MAX_XML_DEPTH).toBe(256);
    expect(readXml(nested(101)).ok).toBe(true);
    expect(readXml(nested(256)).ok).toBe(true);
    expect(readXml(nested(257))).toMatchObject({ ok: false, reason: 'too-deep' });
    expect(readXml(nestedSelfClosing(256)).ok).toBe(true);
    expect(readXml(nestedSelfClosing(257))).toMatchObject({ ok: false, reason: 'too-deep' });
    expect(parseXmlPart('/x.xml', nested(256)).ok).toBe(true);
  });
});

describe('S5: the DTD preflight is quote-aware', () => {
  test('a literal-section opener inside an attribute value cannot hide a DOCTYPE', () => {
    for (const xml of [
      '<r a="<?"><!DOCTYPE x [<!ENTITY e "BOOM">]><y>&e;</y><z b="?>"/></r>',
      '<r a="<!--"/><!DOCTYPE r [<!ENTITY e "BOOM">]><r b="-->">x</r>',
      '<r a="<![CDATA["/><!DOCTYPE r><q b="]]>">x</q>',
      "<r a='<?'><!DOCTYPE x><z b='?>'/></r>",
    ])
      expect(readXml(xml)).toMatchObject({ ok: false, reason: 'dtd-forbidden' });
  });

  test('a custom entity inside an attribute value is still refused', () => {
    expect(readXml('<r a="&e;"/>')).toMatchObject({ ok: false, reason: 'entity-forbidden' });
  });

  test('DTD-only declarations are refused', () => {
    expect(readXml('<!ELEMENT r ANY><r/>')).toMatchObject({ ok: false, reason: 'dtd-forbidden' });
  });

  test('quoted markup characters and processing instructions still read', () => {
    const r = readXml('<?xml version="1.0"?><r a="x>y" b=\'1\'><?pi data?><c/></r>');
    expect(r.ok).toBe(true);
  });
});

describe('K6: attribute values are normalized as XML 1.0 requires', () => {
  test('literal tab, line feed, and carriage return become spaces; references are kept', () => {
    const r = readXml('<r a="x\ny\tz" b="1\r\n2" c="p&#xA;q&#x9;r&#xD;s"/>');
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const el = findElement(r.nodes, 'r')!;
    expect(el.attributes.a).toBe('x y z');
    expect(el.attributes.b).toBe('1 2');
    expect(el.attributes.c).toBe('p\nq\tr\rs');
  });

  test('a literal line feed in an attribute is not saved as a character reference', () => {
    const part = parseXmlPart('/x.xml', '<r a="one\ntwo"/>');
    expect(part.ok).toBe(true);
    if (!part.ok) return;
    expect(serializeXmlPart(part.part)).toContain('a="one two"');
  });
});

describe('K2: XML names use the full XML Name character set', () => {
  test('non-ASCII element and attribute names open and round-trip', () => {
    const xml = '<r xmlns:x="urn:x"><Größe x:é="1">2</Größe><名前/></r>';
    const part = parseXmlPart('/x.xml', xml);
    expect(part.ok).toBe(true);
    if (!part.ok) return;
    const again = parseXmlPart('/x.xml', serializeXmlPart(part.part));
    expect(
      again.ok && again.part.root.children.map((c) => c.type === 'element' && c.local)
    ).toEqual(['Größe', '名前']);
  });

  test('names that are not XML names are still refused', () => {
    expect(parseXmlPart('/x.xml', '<r xmlns:x="urn:x"><x:a:b/></r>').ok).toBe(false);
  });
});

describe('namespace bindings are checked on read, as they are on save', () => {
  test('a reserved or impossible binding refuses the part instead of an unsaveable tree', () => {
    for (const xml of [
      '<r xmlns:foo="http://www.w3.org/XML/1998/namespace"/>',
      '<r xmlns:xml="urn:not-xml"/>',
      '<p:r xmlns:p="u"><p:x xmlns:p=""/></p:r>',
      '<r xmlns:x="u" xmlns:y="u" x:a="1" y:a="2"/>',
    ])
      expect(parseXmlPart('/x.xml', xml)).toMatchObject({ ok: false, reason: 'bad-name' });
  });
});
