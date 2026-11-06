/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Offline checks of the Office.js compatibility tooling. Reads only committed files under
// compat/; never fetches the network.
import { describe, expect, test } from 'bun:test';
import ts from 'typescript';
import { createCompatibilityReport, renderDetails } from '../compat-report.mjs';
import { inventoryModule } from '../lib/signature-inventory.mjs';
import {
  validateManifestAgainstReference,
  validateManifestSchemaVersion,
} from '../lib/manifest-integrity.mjs';
import { validateReferenceFixture } from '../lib/reference-normalize.mjs';
import { validateProvenance } from '../lib/provenance.mjs';
import manifest from '../../compat/manifest.json';
import reference from '../../compat/reference/presentation.reference.json';
import inventory from '../../compat/reference/presentation.full-inventory.json';
import editingScope from '../../compat/editing-scope.json';
import runtimeNotes from '../../compat/runtime-notes.json';
import provenance from '../../compat/provenance.json';
import pins from '../../compat/definitely-typed-commits.json';

type Selection = { members: string[]; namespace?: string; isFunction?: boolean };
type Row = { uid: string; scope: string; kind: string; readonly?: boolean };

const symbols = manifest.symbols as Record<string, Selection>;
const rows = inventory.endpoints as Row[];
const byUid = new Map(rows.map((row) => [row.uid, row]));

function selectedUids(): string[] {
  const uids: string[] = [];
  for (const [name, selection] of Object.entries(symbols)) {
    const namespace = selection.namespace ?? 'PowerPoint';
    if (selection.isFunction) uids.push(`${namespace}.${name}`);
    for (const member of selection.members) uids.push(`${namespace}.${name}#${member}`);
  }
  return uids;
}

describe('pinned reference', () => {
  test('the reference, the full inventory, and the provenance name one release', () => {
    expect(validateReferenceFixture(reference)).toEqual([]);
    expect(validateProvenance(provenance)).toEqual([]);
    const upstream = provenance.upstreamPackage;
    expect(upstream.name).toBe('@types/office-js');
    expect(upstream.version).toBe('1.0.605');
    expect(reference.generatedFrom).toEqual({ package: upstream.name, version: upstream.version });
    expect(inventory.upstream.name).toBe(upstream.name);
    expect(inventory.upstream.version).toBe(upstream.version);
    expect(inventory.upstream.integrity).toBe(upstream.integrity);
    expect(inventory.upstream.sourceUrl).toBe(upstream.sourceRepository.sourceUrl);
    expect((pins.commits as Record<string, string>)[upstream.version]).toBe(
      upstream.sourceRepository.commit
    );
  });

  test('the full inventory has unique UIDs from the presentation namespace only', () => {
    expect(new Set(rows.map((row) => row.uid)).size).toBe(rows.length);
    for (const row of rows) expect(row.uid).toMatch(/^(PowerPoint|OfficeExtension)\./);
  });
});

describe('manifest integrity', () => {
  test('the manifest is a valid subset of the reference', () => {
    expect(validateManifestSchemaVersion(manifest)).toEqual([]);
    expect(validateManifestAgainstReference(manifest, reference)).toEqual([]);
  });

  test('every selected UID exists in the pinned inventory', () => {
    const missing = selectedUids().filter((uid) => !byUid.has(uid));
    expect(missing).toEqual([]);
  });

  test('every selected symbol is in exactly one category, except the run function', () => {
    const categorized = Object.values(manifest.categories).flat();
    expect(new Set(categorized).size).toBe(categorized.length);
    const uncategorized = Object.keys(symbols).filter((name) => !categorized.includes(name));
    expect(uncategorized).toEqual(['run']);
  });

  test('every omission names a real upstream UID that the manifest does not select', () => {
    const selected = new Set(selectedUids());
    for (const omission of manifest.omissions) {
      const known =
        byUid.has(omission.uid) || rows.some((row) => row.uid.startsWith(`${omission.uid}#`));
      expect({ uid: omission.uid, known }).toEqual({ uid: omission.uid, known: true });
      expect(selected.has(omission.uid)).toBe(false);
      expect(omission.reason.trim().length).toBeGreaterThan(0);
    }
  });

  test('every selected class accounts for each upstream member: selected, omitted, or plumbing', () => {
    const plumbing = new Set(['context', 'isNullObject', 'load', 'toJSON', 'track', 'untrack']);
    const omitted = new Set(manifest.omissions.map((omission) => omission.uid));
    const unaccounted: string[] = [];
    for (const [name, selection] of Object.entries(symbols)) {
      if (selection.isFunction || selection.namespace || name === 'RequestContext') continue;
      const prefix = `PowerPoint.${name}#`;
      for (const row of rows) {
        if (!row.uid.startsWith(prefix)) continue;
        const member = row.uid.slice(prefix.length);
        if (member === 'constructor' || plumbing.has(member)) continue;
        if (!selection.members.includes(member) && !omitted.has(row.uid)) unaccounted.push(row.uid);
      }
    }
    expect(unaccounted).toEqual([]);
  });

  test('every listed string enum exists in the pinned inventory', () => {
    for (const name of manifest.enums.names) {
      const members = rows.filter(
        (row) => row.kind === 'enum-member' && row.uid.startsWith(`PowerPoint.${name}.`)
      );
      expect({ name, count: members.length > 0 }).toEqual({ name, count: true });
    }
  });
});

describe('editing scope', () => {
  test('every editing UID exists and is a selected editing method or property write', () => {
    const selected = new Set(selectedUids());
    const endpoints = editingScope.endpoints;
    expect(new Set(endpoints).size).toBe(endpoints.length);
    for (const uid of endpoints) {
      const row = byUid.get(uid);
      expect({ uid, exists: Boolean(row) }).toEqual({ uid, exists: true });
      expect({ uid, selected: selected.has(uid) }).toEqual({ uid, selected: true });
      expect(row!.scope).toBe('api');
      expect(row!.kind === 'method' || (row!.kind === 'property' && !row!.readonly)).toBe(true);
    }
  });

  test('the report accepts the committed scope and runtime notes', () => {
    const report = createCompatibilityReport(inventory, [], runtimeNotes, editingScope);
    expect(report.summary.total).toBe(editingScope.endpoints.length);
    expect(report.summary.missing).toBe(editingScope.endpoints.length);
  });
});

describe('report on in-memory exports', () => {
  /** Inventories a tiny in-memory module with the same compiler pass that reads src/index.ts. */
  function localInventory(source: string) {
    const name = '/index.ts';
    const options = { strict: true, target: ts.ScriptTarget.ES2022 };
    const host = ts.createCompilerHost(options);
    const original = host.getSourceFile;
    host.getSourceFile = (file, ...args) =>
      file === name
        ? ts.createSourceFile(name, source, options.target, true)
        : original(file, ...args);
    const program = ts.createProgram([name], options, host);
    const checker = program.getTypeChecker();
    return inventoryModule(
      checker,
      checker.getSymbolAtLocation(program.getSourceFile(name)!)!,
      'PptxEditor'
    );
  }

  const actual = localInventory(`
    export class ClientObject { readonly isNullObject: boolean = false; }
    export class Shape extends ClientObject {
      name: string = '';
      left: string = '';
      delete(): void {}
    }
    class FontBase extends ClientObject {
      get bold(): boolean | null { return null; }
      set bold(value: boolean | null) {}
    }
    export class ShapeFont extends FontBase {}
  `);
  const scope = {
    schemaVersion: 1,
    description: 'Test scope.',
    endpoints: [
      'PowerPoint.Shape#name',
      'PowerPoint.Shape#left',
      'PowerPoint.Shape#delete',
      'PowerPoint.ShapeFont#bold',
      'PowerPoint.Slide#delete',
    ],
  };

  test('maps PowerPoint.* UIDs to PptxEditor.* exports and follows inheritance', () => {
    const report = createCompatibilityReport(inventory, actual, {}, scope);
    const status = Object.fromEntries(report.endpoints.map((entry) => [entry.uid, entry.status]));
    expect(status).toEqual({
      'PowerPoint.Shape#delete': 'match',
      'PowerPoint.Shape#left': 'different',
      'PowerPoint.Shape#name': 'match',
      'PowerPoint.ShapeFont#bold': 'match',
      'PowerPoint.Slide#delete': 'missing',
    });
    expect(report.summary).toEqual({
      total: 5,
      match: 3,
      different: 1,
      missing: 1,
      percent: 60,
    });
    expect(report.groups['Editing methods'].total).toBe(2);
    expect(report.groups['Property writes'].total).toBe(3);
    expect(renderDetails(report)).toContain('| PowerPoint.Slide#delete | missing | unverified |');
  });

  test('refuses a scope entry that is not an editing member', () => {
    expect(() =>
      createCompatibilityReport(
        inventory,
        actual,
        {},
        {
          schemaVersion: 1,
          endpoints: ['PowerPoint.Shape#id'],
        }
      )
    ).toThrow('Not an editing method or property write');
  });

  test('refuses a runtime note on an unknown UID or with an unknown status', () => {
    expect(() =>
      createCompatibilityReport(
        inventory,
        actual,
        { 'PowerPoint.Ghost#member': { status: 'partial', notes: 'x' } },
        scope
      )
    ).toThrow('Unknown runtime note endpoint');
    expect(() =>
      createCompatibilityReport(
        inventory,
        actual,
        { 'PowerPoint.Shape#name': { status: 'maybe', notes: 'x' } },
        scope
      )
    ).toThrow('Invalid runtime note');
  });
});
