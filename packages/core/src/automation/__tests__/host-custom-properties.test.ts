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

const NS_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const REL_CUSTOM =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CT_CUSTOM = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';

function deck(custom?: string): Uint8Array {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, name: 'Box', x: 0, y: 0, cx: 914400, cy: 914400 }),
      },
    ],
    ...(custom ? { extraParts: [['docProps/custom.xml', CT_CUSTOM, custom]] as const } : {}),
  });
  if (!custom) return bytes;
  // Relate the part from the package, as a writer does.
  const files = unzipSync(bytes);
  const rels = strFromU8(files['_rels/.rels']!).replace(
    '</Relationships>',
    `<Relationship Id="rId9" Type="${REL_CUSTOM}" Target="docProps/custom.xml"/></Relationships>`
  );
  return rezip(files, { '_rels/.rels': rels });
}

function rezip(files: Record<string, Uint8Array>, changed: Record<string, string>): Uint8Array {
  const out: Record<string, Uint8Array> = { ...files };
  for (const [name, text] of Object.entries(changed)) out[name] = strToU8(text);
  return zipSync(out);
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

function errorOf(response: AutomationBatchResponse): { code: string; detail?: string } {
  const failed = response.results.find((r) => r.status === 'error');
  if (failed?.status !== 'error') throw new Error('expected a refusal');
  return {
    code: failed.error.code,
    ...(failed.error.detail ? { detail: failed.error.detail } : {}),
  };
}

function presentation(host: AutomationHost): AutomationHandle {
  const [value] = ok(host, [{ op: 'getPresentation' }]);
  if (value?.kind !== 'handle') throw new Error('no presentation');
  return value.handle;
}

function list(host: AutomationHost) {
  const [value] = ok(host, [{ op: 'readCustomProperties', owner: presentation(host) }]);
  if (value?.kind !== 'customProperties') throw new Error(JSON.stringify(value));
  return value.properties;
}

function save(host: AutomationHost): Uint8Array {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return saved.bytes;
}

const customXml = (bytes: Uint8Array): string =>
  strFromU8(unzipSync(bytes)['docProps/custom.xml']!);

describe('custom document properties', () => {
  test('a deck without the part lists nothing', () => {
    expect(list(open())).toEqual([]);
  });

  test('the first add creates the part, its relationship, and its content type', () => {
    const host = open();
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Client',
        value: { type: 'String', value: 'A & B <x>' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Count',
        value: { type: 'Number', value: 42 },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Ratio',
        value: { type: 'Number', value: 0.25 },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Done',
        value: { type: 'Boolean', value: true },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Due',
        value: { type: 'Date', value: '2026-03-04T05:06:07.000Z' },
      },
    ]);
    const bytes = save(host);
    const files = unzipSync(bytes);
    expect(strFromU8(files['_rels/.rels']!)).toContain(REL_CUSTOM);
    expect(strFromU8(files['[Content_Types].xml']!)).toContain(
      `PartName="/docProps/custom.xml" ContentType="${CT_CUSTOM}"`
    );
    const xml = customXml(bytes);
    expect(xml).toContain('pid="2" name="Client"');
    expect(xml).toContain('pid="6" name="Due"');
    expect(xml).toContain('<vt:lpwstr>A &amp; B &lt;x&gt;</vt:lpwstr>');
    expect(xml).toContain('<vt:i4>42</vt:i4>');
    expect(xml).toContain('<vt:r8>0.25</vt:r8>');
    expect(xml).toContain('<vt:bool>true</vt:bool>');
    expect(xml).toContain('<vt:filetime>2026-03-04T05:06:07Z</vt:filetime>');
    expect(list(open(bytes))).toEqual([
      { key: 'Client', type: 'String', value: 'A & B <x>' },
      { key: 'Count', type: 'Number', value: 42 },
      { key: 'Ratio', type: 'Number', value: 0.25 },
      { key: 'Done', type: 'Boolean', value: true },
      { key: 'Due', type: 'Date', value: '2026-03-04T05:06:07.000Z' },
    ]);
  });

  test('a set of an existing key keeps its name and pid, ignoring case, and changes the type', () => {
    const host = open();
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'Stage',
        value: { type: 'String', value: 'draft' },
      },
    ]);
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'STAGE',
        value: { type: 'Number', value: 3 },
      },
    ]);
    expect(list(host)).toEqual([{ key: 'Stage', type: 'Number', value: 3 }]);
    expect(customXml(save(host))).toContain('pid="2" name="Stage"><vt:i4>3</vt:i4>');
  });

  test('types the object model does not show stay through every edit', () => {
    const custom =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}">` +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="List"><vt:vector size="1" baseType="lpwstr"><vt:lpwstr>a</vt:lpwstr></vt:vector></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Old"><vt:lpstr>kept</vt:lpstr></property>' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Flag"><vt:bool>0</vt:bool></property>' +
      '</Properties>';
    const host = open(deck(custom));
    expect(list(host)).toEqual([
      { key: 'Old', type: 'String', value: 'kept' },
      { key: 'Flag', type: 'Boolean', value: false },
    ]);
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'New',
        value: { type: 'String', value: 'x' },
      },
    ]);
    let xml = customXml(save(host));
    expect(xml).toContain(
      '<vt:vector size="1" baseType="lpwstr"><vt:lpwstr>a</vt:lpwstr></vt:vector>'
    );
    expect(xml).toContain('pid="8" name="New"');
    ok(host, [{ op: 'deleteCustomProperties', owner: presentation(host) }]);
    expect(list(host)).toEqual([]);
    xml = customXml(save(host));
    expect(xml).toContain('name="List"');
    expect(xml).not.toContain('name="Old"');
  });

  test('add with the key of an unlisted property replaces it, as a key names one property', () => {
    const custom =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}">` +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="List"><vt:vector size="1" baseType="lpwstr"><vt:lpwstr>a</vt:lpwstr></vt:vector></property>' +
      '</Properties>';
    const host = open(deck(custom));
    expect(list(host)).toEqual([]);
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'list',
        value: { type: 'String', value: 'x' },
      },
    ]);
    expect(list(host)).toEqual([{ key: 'List', type: 'String', value: 'x' }]);
    const xml = customXml(save(host));
    expect(xml).not.toContain('vt:vector');
    expect(xml).toContain('pid="2" name="List"><vt:lpwstr>x</vt:lpwstr>');
  });

  test('delete removes one property; an unknown key is item-not-found', () => {
    const host = open();
    ok(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'A',
        value: { type: 'String', value: '1' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'B',
        value: { type: 'String', value: '2' },
      },
    ]);
    ok(host, [{ op: 'deleteCustomProperty', owner: presentation(host), key: 'a' }]);
    expect(list(host).map((p) => p.key)).toEqual(['B']);
    expect(
      errorOf(exec(host, [{ op: 'deleteCustomProperty', owner: presentation(host), key: 'A' }]))
        .code
    ).toBe('item-not-found');
    // A delete of all on a deck without the part writes nothing.
    const plain = open();
    const response = exec(plain, [{ op: 'deleteCustomProperties', owner: presentation(plain) }]);
    expect([response.ok, response.changed]).toEqual([true, false]);
  });

  test('invalid keys and values are refused and write nothing', () => {
    const host = open();
    const bad: AutomationOperation[] = [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: '',
        value: { type: 'String', value: 'x' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k'.repeat(256),
        value: { type: 'String', value: 'x' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'bad\u0001',
        value: { type: 'String', value: 'x' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'String', value: 'a\u0000' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'Number', value: Number.NaN },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'Date', value: 'not a date' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'Date', value: '0100-01-01T00:00:00Z' },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'Boolean', value: 'yes' as never },
      },
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'k',
        value: { type: 'Blob', value: 'x' } as never,
      },
      { op: 'setCustomProperty', owner: presentation(host), key: 'k', value: null as never },
    ];
    for (const op of bad) {
      const response = exec(host, [op]);
      expect(errorOf(response).code).toBe('invalid-argument');
      expect(response.changed).toBe(false);
    }
    expect(list(host)).toEqual([]);
  });

  test('a custom properties part with another root is refused', () => {
    const host = open(deck(`<Other xmlns="urn:x"/>`));
    expect(
      errorOf(exec(host, [{ op: 'readCustomProperties', owner: presentation(host) }]))
    ).toEqual({
      code: 'unsupported-content',
      detail: 'custom-properties-root',
    });
  });

  test('a write is one atomic batch with the rest of the request', () => {
    const host = open();
    const response = exec(host, [
      {
        op: 'setCustomProperty',
        owner: presentation(host),
        key: 'A',
        value: { type: 'String', value: '1' },
      },
      { op: 'deleteCustomProperty', owner: presentation(host), key: 'missing' },
    ]);
    expect(response.ok).toBe(false);
    expect(list(host)).toEqual([]);
  });
});
