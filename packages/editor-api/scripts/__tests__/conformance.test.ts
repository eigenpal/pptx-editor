/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The generated Office.js conformance assertions and the source-compatibility fixtures. The
// compiler runs offline on committed files; nothing here fetches the network.
import { afterAll, describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ASSERTIONS_PATH,
  generateConformance,
  generateFromCompat,
  qualifyTypeText,
} from '../generate-conformance.mjs';
import { TYPECHECK_TIMEOUT_MS, typecheckProject } from '../lib/typecheck-compat.mjs';
import manifest from '../../compat/manifest.json';
import inventory from '../../compat/reference/presentation.full-inventory.json';

type Row = { uid: string; [key: string]: unknown };

const packageRoot = path.resolve(import.meta.dir, '..', '..');
const compatTsconfig = path.join(packageRoot, 'compat', 'tsconfig.json');
const rows = inventory.endpoints as Row[];
const rowOf = (uid: string) => {
  const row = structuredClone(rows.find((r) => r.uid === uid)!);
  // `manifest.json` lists `Shape#type` as a conformance exception: the package leaves out the
  // one value that names a trademarked feature. These probes compare the rest of the union.
  if (uid === 'PowerPoint.Shape#type') row.type = String(row.type).replace('"SmartArt" | ', '');
  return row;
};
const temporary: string[] = [];

afterAll(() => {
  for (const directory of temporary) fs.rmSync(directory, { recursive: true, force: true });
});

/**
 * Generate assertions for `members` from `endpoints` in a temporary project, and compile it.
 * The assertions name the real package, or `declarations` when it is given. Answers the
 * diagnostics of the generated file, each as the alias or function that the diagnostic line
 * declares, then the message.
 */
function compileAgainstPackage(
  members: Record<string, string[]>,
  endpoints: Row[],
  declarations?: string
): string[] {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-conformance-'));
  temporary.push(directory);
  if (declarations !== undefined)
    fs.writeFileSync(path.join(directory, 'declarations.ts'), declarations);
  const symbols = Object.fromEntries(
    Object.entries(members).map(([name, list]) => [name, { members: list }])
  );
  const { assertionsSource } = generateConformance({
    inventory: { endpoints },
    manifest: { symbols },
    packageImport:
      declarations === undefined ? path.join(packageRoot, 'src', 'index.ts') : './declarations.ts',
    typeAssertImport: path.join(packageRoot, 'compat', 'type-assert.ts'),
  });
  const file = path.join(directory, 'assertions.ts');
  fs.writeFileSync(file, assertionsSource);
  fs.writeFileSync(
    path.join(directory, 'tsconfig.json'),
    JSON.stringify({
      extends: path.join(packageRoot, 'tsconfig.json'),
      compilerOptions: { noEmit: true },
      include: ['assertions.ts'],
    })
  );
  const lines = assertionsSource.split('\n');
  return typecheckProject(path.join(directory, 'tsconfig.json'), { scope: 'project' }).map(
    (diagnostic) => {
      const [, row, message] = /assertions\.ts:(\d+):\d+ - ([^\n]*)/.exec(diagnostic) ?? [];
      let declared: RegExpExecArray | null = null;
      for (let index = Number(row) - 1; index >= 0 && !declared; index--)
        declared = /^(?:export function|type) (\w+)/.exec(lines[index] ?? '');
      const text = (message ?? diagnostic).replace(/Namespace '"[^"]*"'/, 'The package');
      return `${declared?.[1] ?? '?'}: ${text}`;
    }
  );
}

describe('the conformance assertions', () => {
  test(
    'the compat project (generated assertions and fixtures) type-checks with zero diagnostics',
    () => {
      expect(typecheckProject(compatTsconfig, { scope: 'project' })).toEqual([]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test('the committed generated file is current', () => {
    expect(fs.readFileSync(ASSERTIONS_PATH, 'utf8')).toBe(generateFromCompat().assertionsSource);
  });

  test('every selected member is asserted or named in conformanceExceptions', () => {
    const result = generateFromCompat();
    const selected = Object.values(
      manifest.symbols as Record<string, { members: string[]; isFunction?: boolean }>
    ).reduce(
      (count, selection) => count + (selection.isFunction ? 1 : selection.members.length),
      0
    );
    expect(result.assertions + result.skipped.length).toBe(selected);
    expect(result.skipped.map((entry: { uid: string }) => entry.uid)).toEqual(
      manifest.conformanceExceptions.map((entry) => entry.uid)
    );
    for (const { uid } of result.skipped)
      expect(result.assertionsSource).toContain(
        `// Not asserted: manifest.json conformanceExceptions "${uid}".`
      );
  });

  const members = { Shape: ['fill', 'type'], ShapeCollection: ['addTextBox'], ShapeFont: ['bold'] };
  const uids = [
    'PowerPoint.Shape#fill',
    'PowerPoint.Shape#type',
    'PowerPoint.ShapeCollection#addTextBox',
    'PowerPoint.ShapeFont#bold',
  ];

  test(
    'the pinned reference compiles against the real package',
    () => {
      expect(compileAgainstPackage(members, uids.map(rowOf))).toEqual([]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test(
    'a reference that the package does not match fails its assertion, in either direction',
    () => {
      const [fill, type, addTextBox, bold] = uids.map(rowOf);
      // A writable reference against a read-only property.
      Object.assign(fill!, { readonly: false, writeType: 'ShapeFill' });
      // A narrower union than the package declares: the check is not one-directional.
      type!.type = String(type!.type).replace('"Unsupported" | ', '');
      // A required parameter where the package has an optional one.
      (
        addTextBox!.overloads as { parameters: { optional: boolean }[] }[]
      )[0]!.parameters[1]!.optional = false;
      // A type name that the package does not export.
      bold!.type = 'BoldState';
      expect(compileAgainstPackage(members, [fill!, type!, addTextBox!, bold!])).toEqual([
        "_assert_Shape_fill: Type 'false' does not satisfy the constraint 'true'.",
        "_write_Shape_fill: Cannot assign to 'fill' because it is a read-only property.",
        "_assert_Shape_type: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_ShapeCollection_addTextBox: Type 'false' does not satisfy the constraint 'true'.",
        "Ref_ShapeFont_bold: The package has no exported member 'BoldState'.",
        "_assert_ShapeFont_bold: Type 'false' does not satisfy the constraint 'true'.",
      ]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test(
    'a declaration that drifts from the reference fails its assertion',
    () => {
      // The reference is unchanged. The package side drops `readonly` from `fill`, leaves
      // most values out of ShapeType, and makes the `addTextBox` options required.
      const declarations = `
        export declare class ShapeFill {}
        export interface ShapeAddOptions { left?: number }
        export declare enum ShapeType { image = 'Image' }
        export declare class Shape {
          fill: ShapeFill;
          readonly type: ShapeType | \`\${ShapeType}\`;
        }
        export declare class ShapeCollection {
          addTextBox(text: string, options: ShapeAddOptions | undefined): Shape;
        }
      `;
      const drifted = { Shape: ['fill', 'type'], ShapeCollection: ['addTextBox'] };
      expect(compileAgainstPackage(drifted, uids.slice(0, 3).map(rowOf), declarations)).toEqual([
        "_assert_Shape_fill: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_Shape_type: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_ShapeCollection_addTextBox: Type 'false' does not satisfy the constraint 'true'.",
      ]);
    },
    TYPECHECK_TIMEOUT_MS
  );
});

describe('generateConformance', () => {
  test('qualifies package type names and leaves string literals and type parameters alone', () => {
    expect(qualifyTypeText('"Shape" | ShapeType | null')).toBe(
      '"Shape" | PptxEditor.ShapeType | null'
    );
    expect(qualifyTypeText('( context : RequestContext ) => Promise < T >', ['T'])).toBe(
      '(context: PptxEditor.RequestContext) => Promise<T>'
    );
  });

  test('refuses a conformance exception that names a member the manifest does not select', () => {
    expect(() =>
      generateConformance({
        inventory: { endpoints: [rowOf('PowerPoint.Shape#fill')] },
        manifest: {
          symbols: { Shape: { members: ['fill'] } },
          conformanceExceptions: [{ uid: 'PowerPoint.Shape#name', reason: 'Not selected.' }],
        },
      })
    ).toThrow('is not a selected member');
  });

  test('refuses a selected member that the reference inventory does not list', () => {
    expect(() =>
      generateConformance({
        inventory: { endpoints: [] },
        manifest: { symbols: { Shape: { members: ['fill'] } } },
      })
    ).toThrow('has no row for PowerPoint.Shape#fill');
  });
});
