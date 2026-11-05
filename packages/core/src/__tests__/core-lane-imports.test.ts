// Adapted from docx-editor/packages/core/src/__tests__/core-lane-imports.test.ts at 5993bd223.
//
// The lane DAG, enforced against the real import statements.
//
// `CORE_LANES` declares which lane may import which (`core-lane-graph.ts`). The relative-import
// check in `core-lane-graph.test.ts` reads only `from './...'` with single quotes; it does not
// see a package-subpath import (`@pptx-editor.dev/core/layout`), a side-effect import, a
// dynamic import, or an `import x = require`. So this walks every source file of every lane
// and asserts each import's target lane is one the DAG allows.
//
// VALUE imports must obey the DAG, no exceptions. TYPE-ONLY imports (`import type` /
// `export type ... from` / `import('m').T`) are erased at compile time and cannot move code
// between environments, but they are still coupling, so a cross-lane type edge is permitted
// when the importing lane declares the target in `mayImportTypes`, and otherwise only when
// pinned in GRANDFATHERED_TYPE_EDGES below. The pin RATCHETS: an entry that stops matching
// fails, so the list can only shrink or be consciously edited.
//
// Divergence from DOCX Editor, on purpose: the imports come from `import-scan.ts`, which parses
// each file with the TypeScript compiler. The regex scanner of the original dropped
// `import x = require`, template-literal and commented dynamic imports without a sign. This
// version also refuses an import of the package root (`src/index.ts`, `@pptx-editor.dev/core`)
// from inside a lane, because the root re-exports the editor lane.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { CORE_LANES, CORE_PACKAGE_ROOT, existingLanePath, type LaneName } from './core-lane-graph';
import { failGuard, scanImports } from './import-scan';

const SRC = join(CORE_PACKAGE_ROOT, 'src');
const LANES = Object.keys(CORE_LANES) as LaneName[];
const CORE = '@pptx-editor.dev/core';
const GRAPH_FILE = 'packages/core/src/__tests__/core-lane-graph.ts';
const THIS_FILE = 'packages/core/src/__tests__/core-lane-imports.test.ts';

/**
 * Type-only cross-lane edges the DAG does not allow, pinned file by file.
 *
 * Adding an entry needs the same justification a DAG edit does: say WHY the type cannot
 * live in a lane both sides may import (usually `contracts` or `store`). Empty today: every
 * cross-lane type import follows `mayImport`.
 */
const GRANDFATHERED_TYPE_EDGES: readonly { readonly file: string; readonly to: LaneName }[] = [];

/**
 * VALUE edges the DAG does not allow, pinned until their code moves lane.
 *
 * These are live coupling, not just compile-time: each one is a debt with a tracked issue,
 * and the pin ratchets exactly like the type list. Empty today.
 */
const GRANDFATHERED_VALUE_EDGES: readonly { readonly file: string; readonly to: LaneName }[] = [];

/** The package root barrel. It re-exports the editor lane, so no lane may import it. */
type Target = LaneName | 'package root';

interface ImportEdge {
  readonly file: string; // relative to src/, posix separators
  readonly line: number;
  readonly to: Target;
  readonly typeOnly: boolean;
  readonly specifier: string;
  readonly statement: string;
}

function laneOfFile(relativePath: string): LaneName | null {
  const first = relativePath.split('/')[0]!;
  return LANES.find((lane) => CORE_LANES[lane].directory === `src/${first}`) ?? null;
}

/** The lane a specifier lands in, the package root, or null for anything outside the lanes. */
function targetOf(specifier: string, fromFile: string): Target | null {
  if (specifier.startsWith('.')) {
    const target = relative(SRC, resolve(join(SRC, dirname(fromFile)), specifier));
    const posix = target.split(sep).join('/');
    if (/^index(\.tsx?)?$/.test(posix)) return 'package root';
    return laneOfFile(posix);
  }
  if (specifier === CORE) return 'package root';
  if (!specifier.startsWith(`${CORE}/`)) return null;
  const subpath = `./${specifier.slice(CORE.length + 1)}`;
  for (const lane of LANES) {
    const declared = CORE_LANES[lane];
    if (declared.subpath === subpath || subpath.startsWith(`${declared.subpath}/`)) return lane;
  }
  return null;
}

function* sourceFiles(directory: string): Generator<string> {
  for (const entry of readdirSync(directory)) {
    const path = join(directory, entry);
    if (statSync(path).isDirectory()) {
      if (entry === '__tests__' || entry === 'node_modules') continue;
      yield* sourceFiles(path);
    } else if (/\.tsx?$/.test(entry) && !/\.test\.tsx?$/.test(entry)) {
      yield path;
    }
  }
}

/**
 * The cross-lane edges of one source text, as if it lived at `relativePath` under src/.
 * Imports whose specifier is computed at run time go to `computed`: no guard can follow them.
 */
function edgesOfSource(
  relativePath: string,
  source: string,
  computed: string[] = []
): ImportEdge[] {
  const ownLane = laneOfFile(relativePath);
  if (!ownLane) return [];
  const edges: ImportEdge[] = [];
  for (const found of scanImports(relativePath, source)) {
    if (found.specifier === null) {
      computed.push(`packages/core/src/${relativePath}:${found.line}  ${found.statement}`);
      continue;
    }
    const target = targetOf(found.specifier, relativePath);
    if (target && target !== ownLane) {
      edges.push({
        file: relativePath,
        line: found.line,
        to: target,
        typeOnly: found.typeOnly,
        specifier: found.specifier,
        statement: found.statement,
      });
    }
  }
  return edges;
}

function edgesOf(file: string, computed: string[]): ImportEdge[] {
  const relativePath = relative(SRC, file).split(sep).join('/');
  return edgesOfSource(relativePath, readFileSync(file, 'utf8'), computed);
}

const isLane = (target: Target): target is LaneName => target !== 'package root';

/** Value edges the DAG forbids: judged against `mayImport` only, never `mayImportTypes`. */
function offDagValueEdges(edges: readonly ImportEdge[]): ImportEdge[] {
  return edges.filter(
    (edge) =>
      !edge.typeOnly &&
      isLane(edge.to) &&
      !CORE_LANES[laneOfFile(edge.file)!].mayImport.includes(edge.to)
  );
}

/** Type edges the DAG forbids: neither `mayImport` nor `mayImportTypes` names the target. */
function offDagTypeEdges(edges: readonly ImportEdge[]): ImportEdge[] {
  return edges.filter((edge) => {
    if (!edge.typeOnly || !isLane(edge.to)) return false;
    const lane = CORE_LANES[laneOfFile(edge.file)!];
    return !lane.mayImport.includes(edge.to) && !(lane.mayImportTypes ?? []).includes(edge.to);
  });
}

const allEdges: ImportEdge[] = [];
const computedImports: string[] = [];
let filesScanned = 0;
for (const lane of LANES) {
  for (const file of sourceFiles(existingLanePath(lane))) {
    allEdges.push(...edgesOf(file, computedImports));
    filesScanned += 1;
  }
}

const keyOf = (edge: { readonly file: string; readonly to: Target }): string =>
  `${edge.file} → ${edge.to}`;

/** One line per edge: where it is, what it imports, and which rule it breaks. */
function describeEdge(edge: ImportEdge): string {
  const from = laneOfFile(edge.file)!;
  const allowed = CORE_LANES[from].mayImport;
  const kind = edge.typeOnly ? 'type-only import' : 'value import';
  return [
    `  packages/core/src/${edge.file}:${edge.line}  ${edge.statement}`,
    `    ${kind} of '${edge.specifier}' lands in lane '${edge.to}'; lane '${from}' may import [${allowed.join(', ') || 'nothing'}]`,
  ].join('\n');
}

/** The ratchet: report new edges and stale pins, each with its own fix. */
function checkRatchet(
  kind: 'VALUE' | 'TYPE',
  offDag: readonly ImportEdge[],
  pinned: readonly { readonly file: string; readonly to: LaneName }[],
  pinList: string
): void {
  const pinnedKeys = new Set(pinned.map(keyOf));
  const actualKeys = new Set(offDag.map(keyOf));
  const fresh = offDag.filter((edge) => !pinnedKeys.has(keyOf(edge)));
  const stale = pinned.filter((pin) => !actualKeys.has(keyOf(pin)));
  const lines: string[] = [];
  if (fresh.length > 0) {
    lines.push(`${fresh.length} ${kind} import(s) cross a lane edge the DAG forbids:`);
    lines.push(...fresh.map(describeEdge));
    lines.push(
      'Fix: import the name from a lane that this lane may import. Move a shared type to `contracts` or `store`.',
      `Or, if the edge is intended, add the target to the lane's ${
        kind === 'VALUE' ? '`mayImport`' : '`mayImport` or `mayImportTypes`'
      } in ${GRAPH_FILE} and to docs/architecture/engine-lanes.md, in a diff a reviewer sees.`,
      `Or pin the edge in ${pinList} in ${THIS_FILE}, with the reason it cannot move.`
    );
  }
  if (stale.length > 0) {
    lines.push(`${stale.length} pin(s) in ${pinList} no longer match a real import:`);
    lines.push(...stale.map((pin) => `  ${keyOf(pin)}`));
    lines.push(`Fix: delete these entries from ${pinList} in ${THIS_FILE}. The debt is paid.`);
  }
  failGuard(lines);
}

describe('every lane import obeys the DAG (core-lane-graph.ts)', () => {
  test('the walk is not vacuous: it scans every lane and sees the editor lane import the others', () => {
    expect(filesScanned).toBeGreaterThan(100);
    const fromEditor = allEdges.filter((edge) => edge.file.startsWith('editor/'));
    expect(fromEditor.some((edge) => edge.to === 'layout')).toBe(true);
    expect(fromEditor.some((edge) => edge.to === 'store')).toBe(true);
    expect(fromEditor.some((edge) => edge.to === 'contracts')).toBe(true);
  });

  test('no VALUE import crosses an edge the DAG forbids, beyond the pinned debt', () => {
    checkRatchet(
      'VALUE',
      offDagValueEdges(allEdges),
      GRANDFATHERED_VALUE_EDGES,
      'GRANDFATHERED_VALUE_EDGES'
    );
  });

  test('every off-DAG TYPE edge is pinned, and every pin is still real', () => {
    checkRatchet(
      'TYPE',
      offDagTypeEdges(allEdges),
      GRANDFATHERED_TYPE_EDGES,
      'GRANDFATHERED_TYPE_EDGES'
    );
  });

  test('no lane imports the package root, which re-exports the editor lane', () => {
    const rootEdges = allEdges.filter((edge) => edge.to === 'package root');
    failGuard(
      rootEdges.length === 0
        ? []
        : [
            `${rootEdges.length} import(s) reach the package root (src/index.ts or '${CORE}'):`,
            ...rootEdges.map(
              (edge) => `  packages/core/src/${edge.file}:${edge.line}  ${edge.statement}`
            ),
            'Fix: import the name from the lane that owns it, with a relative path into packages/core/src/<lane>/.',
          ]
    );
  });

  test('every import names a literal module, so the walk can follow it', () => {
    failGuard(
      computedImports.length === 0
        ? []
        : [
            `${computedImports.length} import(s) compute their specifier at run time, so this guard cannot check them:`,
            ...computedImports.map((line) => `  ${line}`),
            "Fix: write the specifier as a string literal, for example import('../layout/index.ts').",
          ]
    );
  });
});

describe('the scanner itself, against the statement shapes that broke it', () => {
  const specifiers = (source: string) =>
    scanImports('example.ts', source).map(({ specifier, typeOnly }) => ({ specifier, typeOnly }));

  test('a clause comment carrying quotes, backticks and semicolons does not drop the import', () => {
    const source = [
      'export {',
      '  layoutSlide,',
      "  // Referenced by `SlideLayout`: don't remove; it is load-bearing.",
      '  layoutTextBody,',
      "} from './slide-layout.ts';",
    ].join('\n');
    expect(specifiers(source)).toEqual([{ specifier: './slide-layout.ts', typeOnly: false }]);
  });

  test('a from-less export type cannot bridge to the next statement and launder it', () => {
    const source = [
      'export type CommandGate = (command: string) => boolean;',
      "import { layoutSlide } from '../layout/slide-layout.ts';",
    ].join('\n');
    expect(specifiers(source)).toEqual([
      { specifier: '../layout/slide-layout.ts', typeOnly: false },
    ]);
  });

  test('a very long specifier list is still scanned: no length budget', () => {
    const names = Array.from({ length: 120 }, (_, index) => `  someLongExportedName${index},`);
    const source = `import {\n${names.join('\n')}\n} from '../layout/index.ts';`;
    expect(specifiers(source)).toEqual([{ specifier: '../layout/index.ts', typeOnly: false }]);
  });

  test('every import form is seen, with its type-only flag', () => {
    const source = [
      "import type {\n  A,\n  B,\n} from 'type-multiline';",
      "import { type C } from 'inline-type-only';",
      "export * as ns from 'export-star-as';",
      "export type * from 'export-type-star';",
      "import legacy = require('import-equals');",
      "import type Legacy = require('import-type-equals');",
      'const lazy = () => import(`template-literal`);',
      "const lazier = () => import(/* chunk */ 'commented-dynamic');",
      "const old = require('plain-require');",
      "type T = import('import-type').T;",
      "import 'side-effect';",
      "// import { ignored } from 'in-a-comment';",
      'const text = "import { ignored } from \'in-a-string\'";',
    ].join('\n');
    expect(specifiers(source)).toEqual([
      { specifier: 'type-multiline', typeOnly: true },
      // Kept as `import {} from` under verbatimModuleSyntax, so the module still loads.
      { specifier: 'inline-type-only', typeOnly: false },
      { specifier: 'export-star-as', typeOnly: false },
      { specifier: 'export-type-star', typeOnly: true },
      { specifier: 'import-equals', typeOnly: false },
      { specifier: 'import-type-equals', typeOnly: true },
      { specifier: 'template-literal', typeOnly: false },
      { specifier: 'commented-dynamic', typeOnly: false },
      { specifier: 'plain-require', typeOnly: false },
      { specifier: 'import-type', typeOnly: true },
      { specifier: 'side-effect', typeOnly: false },
    ]);
  });

  test('a computed dynamic import is reported, not skipped', () => {
    expect(specifiers('const m = (name: string) => import(name);')).toEqual([
      { specifier: null, typeOnly: false },
    ]);
    expect(specifiers('const m = (b: string) => import(`${b}/x`);')).toEqual([
      { specifier: null, typeOnly: false },
    ]);
  });

  test('a value import across a forbidden edge is reported, a type-only one is not a value', () => {
    const value = edgesOfSource(
      'store/example.ts',
      "import { layoutSlide } from '../layout/slide-layout.ts';"
    );
    expect(offDagValueEdges(value).map(keyOf)).toEqual(['store/example.ts → layout']);
    const typeOnly = edgesOfSource(
      'store/example.ts',
      "import type { SlideLayout } from '../layout/slide-layout.ts';"
    );
    expect(offDagValueEdges(typeOnly)).toEqual([]);
    expect(offDagTypeEdges(typeOnly).map(keyOf)).toEqual(['store/example.ts → layout']);
  });

  test('a package-subpath import lands in its lane', () => {
    const edges = edgesOfSource(
      'cascade/example.ts',
      "import { paintSlide } from '@pptx-editor.dev/core/output';"
    );
    expect(offDagValueEdges(edges).map(keyOf)).toEqual(['cascade/example.ts → output']);
  });

  test('an import of the package root is recognized, by path and by name', () => {
    const byPath = edgesOfSource('store/example.ts', "export { x } from '../index.ts';");
    const byName = edgesOfSource('store/example.ts', `export { x } from '${CORE}';`);
    expect([...byPath, ...byName].map(keyOf)).toEqual([
      'store/example.ts → package root',
      'store/example.ts → package root',
    ]);
  });

  test('a failure names the file, the line, the statement, and the rule', () => {
    const [edge] = edgesOfSource(
      'store/example.ts',
      "\nimport { layoutSlide } from '../layout/slide-layout.ts';"
    );
    expect(describeEdge(edge!)).toBe(
      [
        "  packages/core/src/store/example.ts:2  import { layoutSlide } from '../layout/slide-layout.ts';",
        "    value import of '../layout/slide-layout.ts' lands in lane 'layout'; lane 'store' may import [nothing]",
      ].join('\n')
    );
  });
});
