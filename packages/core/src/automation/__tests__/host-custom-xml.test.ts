import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
const NS_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
const ID_A = '{11111111-2222-3333-4444-555555555555}';
const ID_B = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}';

const props = (id: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${id}" xmlns:ds="${NS_DS}"><ds:schemaRefs/></ds:datastoreItem>`;
const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

/** Rewrite named files of a zip. */
function patch(bytes: Uint8Array, edits: Record<string, (text: string) => string>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [name, change] of Object.entries(edits))
    files[name] = strToU8(change(files[name] ? strFromU8(files[name]) : ''));
  return zipSync(files);
}

/** Two slides; a box on the first. */
function deck(): Uint8Array {
  return buildPptx({
    slides: [
      { layout: 'blank', shapes: shape({ id: 2, name: 'Box', x: IN, y: IN, cx: IN, cy: IN }) },
      { layout: 'blank', shapes: '' },
    ],
  });
}

/**
 * A deck whose presentation holds part A, and whose two slides both name part B (as a
 * duplicated slide does). Slide 1 also has a tags entry in its customer data list.
 */
function deckWithParts(): Uint8Array {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: '',
        rels: [
          ['rIdB', `${R}/customXml`, '../../customXml/item2.xml'],
          ['rIdT', `${R}/tags`, '../tags/tag1.xml'],
        ],
      },
      {
        layout: 'blank',
        shapes: '',
        rels: [['rIdB', `${R}/customXml`, '../../customXml/item2.xml']],
      },
    ],
    presentationRels: [['rIdA', `${R}/customXml`, '../customXml/item1.xml']],
    extraParts: [
      ['customXml/item1.xml', 'application/xml', '<a:root xmlns:a="urn:a"><a:v>1</a:v></a:root>'],
      ['customXml/itemProps1.xml', CT_PROPS, props(ID_A)],
      ['customXml/item2.xml', 'application/xml', '<b xmlns="urn:b">shared</b>'],
      ['customXml/itemProps2.xml', CT_PROPS, props(ID_B)],
      [
        'ppt/tags/tag1.xml',
        'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
        '<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="K" val="v"/></p:tagLst>',
      ],
    ],
  });
  const list = (extra: string) =>
    `</p:spTree><p:custDataLst><p:custData r:id="rIdB"/>${extra}</p:custDataLst></p:cSld>`;
  return patch(bytes, {
    'customXml/_rels/item1.xml.rels': () =>
      rels(`<Relationship Id="rId1" Type="${R}/customXmlProps" Target="itemProps1.xml"/>`),
    'customXml/_rels/item2.xml.rels': () =>
      rels(`<Relationship Id="rId1" Type="${R}/customXmlProps" Target="itemProps2.xml"/>`),
    'ppt/slides/slide1.xml': (t) =>
      t.replace('</p:spTree></p:cSld>', list('<p:tags r:id="rIdT"/>')),
    'ppt/slides/slide2.xml': (t) => t.replace('</p:spTree></p:cSld>', list('')),
  });
}

function open(bytes = deck()): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function exec(host: AutomationHost, operations: AutomationOperation[]): AutomationBatchResponse {
  return host.execute({ operations });
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = exec(host, operations);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function refusal(host: AutomationHost, operations: AutomationOperation[]) {
  const response = exec(host, operations);
  expect(response.ok).toBe(false);
  expect(response.changed).toBe(false);
  const failed = response.results.find((r) => r.status === 'error');
  if (failed?.status !== 'error') throw new Error('expected a refusal');
  return { code: failed.error.code, detail: failed.error.detail };
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

/** The presentation, the first slide, its layout and master, and the box on it. */
function owners(host: AutomationHost) {
  const values = ok(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getSlideLayout', slide: { result: 1 } },
    { op: 'getSlideMaster', owner: { result: 1 } },
    { op: 'listShapes', container: { result: 1 } },
  ]);
  const shapes = values[4]?.kind === 'handles' ? values[4].handles : [];
  return {
    presentation: handle(values[0]),
    slide: handle(values[1]),
    layout: handle(values[2]),
    master: handle(values[3]),
    shape: shapes[0],
  };
}

function list(host: AutomationHost, owner: AutomationHandle, namespaceUri?: string) {
  const [value] = ok(host, [
    { op: 'listCustomXmlParts', owner, ...(namespaceUri === undefined ? {} : { namespaceUri }) },
  ]);
  if (value?.kind !== 'customXmlParts') throw new Error(JSON.stringify(value));
  return value.parts;
}

function xmlOf(host: AutomationHost, owner: AutomationHandle, id: string): string {
  const [value] = ok(host, [{ op: 'readCustomXml', owner, id }]);
  if (value?.kind !== 'text') throw new Error(JSON.stringify(value));
  return value.text;
}

function save(host: AutomationHost): Uint8Array {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return saved.bytes;
}

const files = (bytes: Uint8Array): Record<string, string> =>
  Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));

describe('custom XML parts in a file', () => {
  test('the presentation and each slide list the parts they name', () => {
    const host = open(deckWithParts());
    const { presentation, slide } = owners(host);
    expect(list(host, presentation)).toEqual([{ id: ID_A, namespaceUri: 'urn:a' }]);
    expect(list(host, slide)).toEqual([{ id: ID_B, namespaceUri: 'urn:b' }]);
    expect(list(host, slide, 'urn:a')).toEqual([]);
    expect(xmlOf(host, presentation, ID_A.toLowerCase())).toBe(
      '<a:root xmlns:a="urn:a"><a:v>1</a:v></a:root>'
    );
    // An unbraced id finds the part too.
    expect(xmlOf(host, slide, ID_B.slice(1, -1))).toBe('<b xmlns="urn:b">shared</b>');
  });

  test('a delete from one slide keeps a part another slide still names', () => {
    const host = open(deckWithParts());
    const { slide } = owners(host);
    const [second] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    ok(host, [{ op: 'deleteCustomXmlPart', owner: slide, id: ID_B }]);
    expect(list(host, slide)).toEqual([]);
    expect(list(host, handle(second))).toEqual([{ id: ID_B, namespaceUri: 'urn:b' }]);
    let out = files(save(host));
    // The tags entry stays, so the list stays.
    expect(out['ppt/slides/slide1.xml']).toContain(
      '<p:custDataLst><p:tags r:id="rIdT"/></p:custDataLst>'
    );
    expect(out['ppt/slides/_rels/slide1.xml.rels']).not.toContain('customXml');
    expect(out['customXml/item2.xml']).toBeDefined();
    ok(host, [{ op: 'deleteCustomXmlPart', owner: handle(second), id: ID_B }]);
    out = files(save(host));
    expect(out['customXml/item2.xml']).toBeUndefined();
    expect(out['customXml/itemProps2.xml']).toBeUndefined();
    expect(out['customXml/_rels/item2.xml.rels']).toBeUndefined();
    expect(out['[Content_Types].xml']).not.toContain('itemProps2.xml');
    // A list with nothing left goes.
    expect(out['ppt/slides/slide2.xml']).not.toContain('custDataLst');
    expect(out['customXml/item1.xml']).toBeDefined();
  });

  test('planted relationships do not make a part', () => {
    const bytes = patch(deckWithParts(), {
      // Point part A's properties at a part of another content type.
      'customXml/_rels/item1.xml.rels': () =>
        rels(`<Relationship Id="rId1" Type="${R}/customXmlProps" Target="item2.xml"/>`),
      // Name a relationships part as a custom XML part.
      'ppt/_rels/presentation.xml.rels': (t) =>
        t.replace(
          '</Relationships>',
          `<Relationship Id="rIdX" Type="${R}/customXml" Target="slides/_rels/slide1.xml.rels"/></Relationships>`
        ),
    });
    const host = open(bytes);
    expect(list(host, owners(host).presentation)).toEqual([]);
  });

  test('many planted ids that lead nowhere list in bounded time', () => {
    const count = 40_000;
    const planted = Array.from(
      { length: count },
      (_, i) =>
        `<Relationship Id="rIdP${i}" Type="${R}/customXml" Target="../customXml/none${i}.xml"/>`
    ).join('');
    const names = Array.from({ length: count }, (_, i) => `<p:custData r:id="rIdQ${i}"/>`).join('');
    const bytes = patch(deckWithParts(), {
      'ppt/_rels/presentation.xml.rels': (t) =>
        t.replace('</Relationships>', `${planted}</Relationships>`),
      'ppt/slides/slide1.xml': (t) =>
        t.replace('<p:custData r:id="rIdB"/>', `<p:custData r:id="rIdB"/>${names}`),
    });
    const host = open(bytes);
    const { presentation, slide } = owners(host);
    const started = performance.now();
    expect(list(host, presentation)).toEqual([{ id: ID_A, namespaceUri: 'urn:a' }]);
    expect(list(host, slide)).toEqual([{ id: ID_B, namespaceUri: 'urn:b' }]);
    // A scan per id over every relationship took seconds at this size; a map lookup does not.
    expect(performance.now() - started).toBeLessThan(1000);
  });
});

describe('custom XML part writes', () => {
  test('add on every owner kind survives a save and reopen', () => {
    const host = open();
    const o = owners(host);
    const kinds = [o.presentation, o.slide, o.layout, o.master, o.shape!];
    const added = ok(
      host,
      kinds.map((owner, i) => ({
        op: 'addCustomXmlPart' as const,
        owner,
        xml: `<?xml version="1.0" encoding="UTF-8"?><d${i} xmlns="urn:k${i}">v${i}</d${i}>`,
        id: `{0000000${i}-0000-4000-8000-000000000000}`,
      }))
    );
    expect(added[4]).toEqual({
      kind: 'customXmlParts',
      parts: [{ id: '{00000004-0000-4000-8000-000000000000}', namespaceUri: 'urn:k4' }],
    });
    const out = files(save(host));
    expect(out['customXml/item1.xml']).toBe(
      '<?xml version="1.0" encoding="UTF-8"?><d0 xmlns="urn:k0">v0</d0>'
    );
    expect(out['customXml/itemProps1.xml']).toContain(
      'ds:itemID="{00000000-0000-4000-8000-000000000000}"'
    );
    expect(out['customXml/itemProps1.xml']).toContain('<ds:schemaRef ds:uri="urn:k0"/>');
    expect(out['customXml/_rels/item1.xml.rels']).toContain('customXmlProps');
    expect(out['[Content_Types].xml']).toContain(
      `PartName="/customXml/itemProps5.xml" ContentType="${CT_PROPS}"`
    );
    expect(out['ppt/_rels/presentation.xml.rels']).toContain('../customXml/item1.xml');
    expect(out['ppt/slides/slide1.xml']).toMatch(
      /<\/p:spTree><p:custDataLst><p:custData r:id="rId\d+"\/>/
    );
    expect(out['ppt/slideLayouts/slideLayout2.xml']).toContain('<p:custDataLst><p:custData');
    expect(out['ppt/slideMasters/slideMaster1.xml']).toContain('<p:custDataLst><p:custData');
    expect(out['ppt/slides/slide1.xml']).toMatch(
      /<p:nvPr><p:custDataLst><p:custData r:id="rId\d+"\/><\/p:custDataLst><\/p:nvPr>/
    );
    const again = open(save(host));
    const p = owners(again);
    [p.presentation, p.slide, p.layout, p.master, p.shape!].forEach((owner, i) => {
      const id = `{0000000${i}-0000-4000-8000-000000000000}`;
      expect(list(again, owner)).toEqual([{ id, namespaceUri: `urn:k${i}` }]);
      expect(xmlOf(again, owner, id)).toContain(`>v${i}</d${i}>`);
    });
  });

  test('without an id the host draws a GUID; a part without a namespace lists an empty one', () => {
    const host = open();
    const { presentation } = owners(host);
    const [value] = ok(host, [{ op: 'addCustomXmlPart', owner: presentation, xml: '<plain/>' }]);
    if (value?.kind !== 'customXmlParts') throw new Error('no part');
    expect(value.parts[0]!.id).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/
    );
    expect(list(host, presentation, '')).toEqual([{ id: value.parts[0]!.id, namespaceUri: '' }]);
  });

  test('setXml replaces the text and can change the namespace', () => {
    const host = open(deckWithParts());
    const { presentation } = owners(host);
    ok(host, [
      { op: 'setCustomXml', owner: presentation, id: ID_A, xml: '<n xmlns="urn:new">2</n>' },
    ]);
    expect(list(host, presentation)).toEqual([{ id: ID_A, namespaceUri: 'urn:new' }]);
    const again = open(save(host));
    expect(xmlOf(again, owners(again).presentation, ID_A)).toBe('<n xmlns="urn:new">2</n>');
  });

  test('setXml points the schema references at a new root namespace', () => {
    const bytes = patch(deckWithParts(), {
      'customXml/itemProps1.xml': () =>
        props(ID_A).replace(
          '<ds:schemaRefs/>',
          '<ds:schemaRefs><ds:schemaRef ds:uri="urn:a"/><ds:schemaRef ds:uri="urn:other"/></ds:schemaRefs>'
        ),
    });
    const host = open(bytes);
    const { presentation } = owners(host);
    // The same namespace leaves the references alone.
    ok(host, [
      { op: 'setCustomXml', owner: presentation, id: ID_A, xml: '<a:r xmlns:a="urn:a"/>' },
    ]);
    expect(files(save(host))['customXml/itemProps1.xml']).toContain('ds:uri="urn:a"');
    ok(host, [
      { op: 'setCustomXml', owner: presentation, id: ID_A, xml: '<n xmlns="urn:new">2</n>' },
    ]);
    const out = files(save(host))['customXml/itemProps1.xml']!;
    expect(out).toContain(
      '<ds:schemaRefs><ds:schemaRef ds:uri="urn:new"/><ds:schemaRef ds:uri="urn:other"/></ds:schemaRefs>'
    );
    expect(out).toContain(`ds:itemID="${ID_A}"`);
    const again = open(save(host));
    expect(list(again, owners(again).presentation)).toEqual([
      { id: ID_A, namespaceUri: 'urn:new' },
    ]);
  });

  test('a delete of a new part leaves no trace', () => {
    const host = open();
    const before = files(save(host));
    const { shape } = owners(host);
    ok(host, [{ op: 'addCustomXmlPart', owner: shape!, xml: '<x/>', id: ID_A }]);
    ok(host, [{ op: 'deleteCustomXmlPart', owner: shape!, id: ID_A }]);
    const after = files(save(host));
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    expect(after['ppt/slides/slide1.xml']).not.toContain('custDataLst');
  });

  test('untrusted XML and bad ids are refused, and nothing is written', () => {
    const host = open(deckWithParts());
    const { presentation, slide } = owners(host);
    const add = (xml: unknown, id?: string): AutomationOperation =>
      ({
        op: 'addCustomXmlPart',
        owner: presentation,
        xml,
        ...(id ? { id } : {}),
      }) as AutomationOperation;
    for (const xml of [
      '',
      'not xml',
      '<a/><b/>',
      '<a>',
      '<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>',
      '<?xml version="1.0" encoding="UTF-16"?><a/>',
      `<a>${'x'.repeat(8 * 1024 * 1024)}</a>`,
      42,
    ])
      expect(refusal(host, [add(xml)])).toEqual({ code: 'invalid-argument', detail: 'custom-xml' });
    expect(refusal(host, [add('<a/>', 'nope')]).detail).toBe('id');
    expect(refusal(host, [add('<a/>', ID_B)]).detail).toBe('custom-xml-id');
    expect(refusal(host, [{ op: 'readCustomXml', owner: slide, id: ID_A }]).code).toBe(
      'item-not-found'
    );
    expect(
      refusal(host, [{ op: 'setCustomXml', owner: presentation, id: ID_A, xml: '<<' }]).detail
    ).toBe('custom-xml');
    expect(
      refusal(host, [
        add('<ok/>', '{99999999-0000-4000-8000-000000000000}'),
        { op: 'deleteCustomXmlPart', owner: slide, id: ID_A },
      ]).code
    ).toBe('item-not-found');
    expect(list(host, presentation)).toEqual([{ id: ID_A, namespaceUri: 'urn:a' }]);
  });
});
