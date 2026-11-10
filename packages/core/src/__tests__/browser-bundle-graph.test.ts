// Adapted from docx-editor/packages/core/src/__tests__/browser-bundle-graph.test.ts at 5993bd223.
//
// What a browser import actually pulls in.
//
// The lane DAG says which lane may import which. That is a rule about DECLARED dependencies,
// and a bundler does not read it: it follows `import` statements, through re-export barrels,
// into whatever they reach. A single `export *` in a barrel is enough to put a Yjs document or
// a PDF writer in every consumer's bundle while every package.json still looks correct.
//
// So this walks the real import graph from the browser entry points (every code subpath of
// `@pptx-editor.dev/core`, and the root of `@pptx-editor.dev/react` and `@pptx-editor.dev/vue`)
// and asserts what it can reach. It resolves relative imports within a package and package
// imports across the workspace, which is what a bundler does; it stops at third-party names and
// records them, which is enough to catch the dependencies that matter here.
//
// Divergences from DOCX Editor, on purpose:
//  - Imports come from `import-scan.ts` (the TypeScript parser), not a regex with a 400-char
//    budget. The regex dropped long import lists, `import x = require`, template-literal and
//    commented dynamic imports, and read `import("m")` inside comments.
//  - A Node builtin is caught with or without the `node:` prefix (`fs` as well as `node:fs`).
//  - Every runtime dependency of the pro and PDF packages must be in the forbidden list.
//  - A failure prints the import chain from the entry to the offending import.

import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import {
  BROWSER_FORBIDDEN_DEPENDENCIES,
  CORE_LANES,
  existingLanePath,
  type LaneName,
} from './core-lane-graph';
import { failGuard, scanImports } from './import-scan';

const PACKAGES = resolve(import.meta.dir, '..', '..', '..');
const REPO = resolve(PACKAGES, '..');
const GRAPH_FILE = 'packages/core/src/__tests__/core-lane-graph.ts';

const CORE = '@pptx-editor.dev/core';

/** Throws when a source root is missing, so a moved package cannot pass vacuously. */
function existingSourceRoot(relativeToPackages: string): string {
  const p = join(PACKAGES, relativeToPackages);
  if (!existsSync(p)) throw new Error(`source root ${p} does not exist`);
  return p;
}

/**
 * Importable specifier to the source root it resolves to.
 *
 * Built from the lane DAG rather than from directory names, so every lane subpath is followed
 * as `@pptx-editor.dev/core/<lane>`. The other workspace packages are mapped too, the
 * Node-only and pro ones included: if a browser entry ever reaches one, the walk follows it
 * into its third-party imports, and the checks below see them.
 */
const WORKSPACE: ReadonlyMap<string, string> = new Map([
  [CORE, existingSourceRoot('core/src')],
  ...(Object.keys(CORE_LANES) as LaneName[]).map(
    (lane) => [`${CORE}/${CORE_LANES[lane].subpath.slice(2)}`, existingLanePath(lane)] as const
  ),
  ['@pptx-editor.dev/i18n', existingSourceRoot('i18n/src')],
  ['@pptx-editor.dev/react', existingSourceRoot('react/src')],
  ['@pptx-editor.dev/vue', existingSourceRoot('vue/src')],
  ['@pptx-editor.dev/pro', existingSourceRoot('pro/src')],
  ['@pptx-editor.dev/pptx-to-pdf', existingSourceRoot('pptx-to-pdf/src')],
  ['@eigenpal/ooxml-core', existingSourceRoot('ooxml-core/src')],
]);

/** LONGEST first: every lane subpath begins with the core package's own name. */
const WORKSPACE_NAMES = [...WORKSPACE.keys()].sort((a, b) => b.length - a.length);

/** Workspace packages that are not part of a default browser import. */
const NOT_A_DEFAULT_BROWSER_IMPORT: readonly string[] = [
  '@pptx-editor.dev/pro',
  '@pptx-editor.dev/pptx-to-pdf',
];

/** Node builtins, by bare name: `fs` reaches a bundle as surely as `node:fs`. */
const NODE_BUILTINS = new Set(builtinModules.filter((name) => !name.startsWith('_')));

function isNodeBuiltin(specifier: string): boolean {
  if (specifier.startsWith('node:')) return true;
  return NODE_BUILTINS.has(specifier) || NODE_BUILTINS.has(specifier.split('/')[0]!);
}

function resolveFile(candidate: string): string | null {
  for (const suffix of ['', '.ts', '.tsx', '/index.ts', '/index.tsx']) {
    const path = `${candidate}${suffix}`;
    if (existsSync(path) && statSync(path).isFile()) return path;
  }
  return null;
}

/** Where an import was written: the file, the line, and the statement. */
interface Site {
  readonly file: string;
  readonly line: number;
  readonly statement: string;
}

interface Reach {
  /** Workspace packages reached, by name, with the first import that reached each. */
  readonly packages: Map<string, Site>;
  /** Third-party specifiers reached, with the first import that reached each. */
  readonly external: Map<string, Site>;
  /** Imports the walk could not follow: a walk with holes proves nothing. */
  readonly unresolved: Site[];
  /** Each reached file and the import that first reached it, for the chain in a failure. */
  readonly parent: Map<string, Site | null>;
}

/** Each file is parsed once, though ten entries walk it. */
const scanned = new Map<string, ReturnType<typeof scanImports>>();
function importsOf(file: string): ReturnType<typeof scanImports> {
  let found = scanned.get(file);
  if (!found) {
    found = scanImports(file, readFileSync(file, 'utf8'));
    scanned.set(file, found);
  }
  return found;
}

/** Walk the import graph from an entry file, breadth first so each chain is the shortest. */
function reachFrom(entry: string): Reach {
  const packages = new Map<string, Site>();
  const external = new Map<string, Site>();
  const unresolved: Site[] = [];
  const parent = new Map<string, Site | null>([[entry, null]]);
  const queue = [entry];
  const enqueue = (file: string, site: Site) => {
    if (parent.has(file)) return;
    parent.set(file, site);
    queue.push(file);
  };

  for (let next = 0; next < queue.length; next += 1) {
    const file = queue[next]!;
    for (const found of importsOf(file)) {
      const site: Site = { file, line: found.line, statement: found.statement };
      const specifier = found.specifier;
      if (specifier === null) {
        unresolved.push(site);
        continue;
      }
      if (specifier.startsWith('.')) {
        const resolved = resolveFile(resolve(dirname(file), specifier.replace(/\.tsx?$/, '')));
        if (resolved) enqueue(resolved, site);
        else unresolved.push(site);
        continue;
      }
      const workspace = WORKSPACE_NAMES.find(
        (name) => specifier === name || specifier.startsWith(`${name}/`)
      );
      if (workspace) {
        if (!packages.has(workspace)) packages.set(workspace, site);
        const entryFile = resolveFile(join(WORKSPACE.get(workspace)!, 'index'));
        if (entryFile) enqueue(entryFile, site);
        else unresolved.push(site);
        continue;
      }
      if (!external.has(specifier)) external.set(specifier, site);
    }
  }
  return { packages, external, unresolved, parent };
}

const shortPath = (file: string): string => relative(REPO, file);

/** The import chain from the entry to `site`, one import per line. */
function chainTo(reach: Reach, site: Site): string[] {
  const steps: Site[] = [site];
  for (let at = reach.parent.get(site.file); at; at = reach.parent.get(at.file)) steps.push(at);
  return steps
    .reverse()
    .map(
      (step, index) =>
        `${'  '.repeat(index + 2)}${shortPath(step.file)}:${step.line}  ${step.statement}`
    );
}

function existingEntry(path: string): string {
  if (!existsSync(path)) throw new Error(`browser entry ${path} does not exist`);
  return path;
}

/** Every entry a browser loads by default: core's code subpaths and the adapters' roots. */
const BROWSER_ENTRIES: readonly (readonly [string, string])[] = [
  [CORE, existingEntry(join(PACKAGES, 'core/src/index.ts'))],
  ...(Object.keys(CORE_LANES) as LaneName[]).map(
    (lane) =>
      [
        `${CORE}/${CORE_LANES[lane].subpath.slice(2)}`,
        existingEntry(join(existingLanePath(lane), 'index.ts')),
      ] as const
  ),
  ['@pptx-editor.dev/react', existingEntry(join(PACKAGES, 'react/src/index.ts'))],
  ['@pptx-editor.dev/vue', existingEntry(join(PACKAGES, 'vue/src/index.ts'))],
];

function forbiddenIn(specifiers: Iterable<string>): string[] {
  const all = [...specifiers];
  return BROWSER_FORBIDDEN_DEPENDENCIES.filter((forbidden) =>
    all.some((specifier) => specifier === forbidden || specifier.startsWith(`${forbidden}/`))
  );
}

/** A failure for each reached name: the rule, the chain that reaches it, and the fix. */
function reachFailure(
  entry: string,
  reach: Reach,
  reached: readonly [string, Site][],
  rule: string,
  fix: string
): string[] {
  if (reached.length === 0) return [];
  return [
    `A browser import of ${entry} reaches ${reached.length} name(s) that ${rule}:`,
    ...reached.flatMap(([name, site]) => [`  '${name}', through:`, ...chainTo(reach, site)]),
    `Fix: ${fix}`,
  ];
}

const FIX_ISOLATE =
  'move the import behind a separate entry that browsers do not load (for example ' +
  '`@pptx-editor.dev/pptx-to-pdf` or an `EditorModule` in `@pptx-editor.dev/pro`), or remove ' +
  'the re-export that pulls it into this graph.';

const runtimeDependencies = (pkg: string): string[] => {
  const manifest = JSON.parse(readFileSync(join(PACKAGES, pkg, 'package.json'), 'utf8')) as {
    dependencies?: Record<string, string>;
  };
  return Object.keys(manifest.dependencies ?? {});
};

describe('the browser entries are the package exports', () => {
  test('every code export of core, react, and vue is walked', () => {
    const exported: string[] = [];
    for (const pkg of ['core', 'react', 'vue']) {
      const manifest = JSON.parse(readFileSync(join(PACKAGES, pkg, 'package.json'), 'utf8')) as {
        name: string;
        exports: Record<string, unknown>;
      };
      for (const key of Object.keys(manifest.exports)) {
        if (key === './package.json' || /\.(css|json|wasm)$/.test(key)) continue;
        exported.push(key === '.' ? manifest.name : `${manifest.name}/${key.slice(2)}`);
      }
    }
    expect(BROWSER_ENTRIES.map(([name]) => name).sort()).toEqual(exported.sort());
  });

  test('every runtime dependency exclusive to pro and PDF is forbidden in a browser', () => {
    const missing = ['pro', 'pptx-to-pdf'].flatMap((pkg) =>
      runtimeDependencies(pkg)
        // The shared font catalog is also a core dependency and has browser-safe entries.
        .filter((dependency) => dependency !== '@docx-editor.dev/fonts')
        .filter((dependency) => !BROWSER_FORBIDDEN_DEPENDENCIES.includes(dependency))
        .map((dependency) => `  '${dependency}' (packages/${pkg}/package.json)`)
    );
    failGuard(
      missing.length === 0
        ? []
        : [
            'A Node-only or pro runtime dependency is not in BROWSER_FORBIDDEN_DEPENDENCIES:',
            ...missing,
            `Fix: add it to BROWSER_FORBIDDEN_DEPENDENCIES in ${GRAPH_FILE}.`,
          ]
    );
  });
});

describe.each(BROWSER_ENTRIES.map(([name, file]) => ({ name, file })))(
  'a browser import of $name cannot reach a Node-only or pro runtime',
  ({ name, file }) => {
    const reach = reachFrom(file);

    test('the walk resolves every import it follows', () => {
      failGuard(
        reach.unresolved.length === 0
          ? []
          : [
              `The walk from ${name} cannot follow ${reach.unresolved.length} import(s), so it cannot prove the graph is browser-safe:`,
              ...reach.unresolved.map(
                (site) => `  ${shortPath(site.file)}:${site.line}  ${site.statement}`
              ),
              'Fix: point the import at a file that exists, and write its specifier as a string literal.',
            ]
      );
    });

    test('it does not reach a forbidden runtime', () => {
      const reached = [...reach.external].filter(
        ([specifier]) => forbiddenIn([specifier]).length > 0
      );
      failGuard(
        reachFailure(
          name,
          reach,
          reached,
          `are in BROWSER_FORBIDDEN_DEPENDENCIES (${GRAPH_FILE})`,
          FIX_ISOLATE
        )
      );
    });

    test('it pulls in no Node builtin, so the graph is genuinely browser-safe', () => {
      // `node:` specifiers are never npm dependencies, so a package.json check cannot see
      // them; this is the only place that constraint can actually fail. Bare names count too.
      const reached = [...reach.external].filter(([specifier]) => isNodeBuiltin(specifier));
      failGuard(
        reachFailure(
          name,
          reach,
          reached,
          'are Node builtins',
          `${FIX_ISOLATE} Code that must run in Node and in browsers reaches Node APIs through \`globalThis.process.getBuiltinModule\`, as \`packages/core/src/layout/node-module-shim.ts\` does.`
        )
      );
    });

    test('it does not reach the pro or PDF packages', () => {
      const reached = [...reach.packages].filter(([pkg]) =>
        NOT_A_DEFAULT_BROWSER_IMPORT.includes(pkg)
      );
      failGuard(
        reachFailure(
          name,
          reach,
          reached,
          'are not part of a default browser import',
          'core and the adapters never import `@pptx-editor.dev/pro` or `@pptx-editor.dev/pptx-to-pdf`; pro code reaches the engine only through the `EditorModule` seam (packages/core/src/editor/modules.ts).'
        )
      );
    });

    if (name === '@pptx-editor.dev/react' || name === '@pptx-editor.dev/vue') {
      test('the walk is not vacuous: it reaches the engine lanes the adapter really uses', () => {
        // If resolution silently failed, every assertion above would pass over an empty graph.
        expect(reach.packages.has(`${CORE}/editor`)).toBe(true);
        expect(reach.packages.has('@pptx-editor.dev/i18n')).toBe(true);
        expect(reach.external.size).toBeGreaterThan(0);
      });
    }
  }
);

describe('the core walk is not vacuous', () => {
  test('the editor lane reaches store, layout, and the OPC layer', () => {
    const reach = reachFrom(join(existingLanePath('editor'), 'index.ts'));
    // Lanes are imported relatively inside core, so they show up as files reached through
    // the relative walk; the shared OPC layer and the font package show up by name.
    const files = [...reach.parent.keys()].map(shortPath);
    expect(files.some((file) => file.startsWith('packages/core/src/store/'))).toBe(true);
    expect(files.some((file) => file.startsWith('packages/core/src/layout/'))).toBe(true);
    expect(reach.packages.has('@eigenpal/ooxml-core')).toBe(true);
    expect(reach.external.has('@docx-editor.dev/fonts')).toBe(true);
  });
});

describe('the checks would FIRE if the graph ever reached a forbidden name', () => {
  test('the forbidden check fires on a pretend graph', () => {
    // The control. The assertions above pass, and a check that only ever passes is
    // indistinguishable from one that cannot fail, so the detection is exercised against a
    // graph that does contain forbidden names.
    const pretend = new Set([
      'yjs/dist/y.mjs',
      'node:fs',
      'pdf-lib',
      'harfbuzzjs',
      'fs',
      'path/posix',
    ]);
    expect(forbiddenIn(pretend)).toEqual(['yjs', 'pdf-lib', 'node:fs']);
    expect([...pretend].filter(isNodeBuiltin)).toEqual(['node:fs', 'fs', 'path/posix']);
  });

  test('the pro package, if reached, would bring its Yjs runtime with it', () => {
    const pro = reachFrom(join(WORKSPACE.get('@pptx-editor.dev/pro')!, 'index.ts'));
    expect(forbiddenIn(pro.external.keys())).toContain('yjs');
  });

  test('the PDF package, if reached, would bring its Node-only runtime with it', () => {
    const pdf = reachFrom(join(WORKSPACE.get('@pptx-editor.dev/pptx-to-pdf')!, 'index.ts'));
    expect(forbiddenIn(pdf.external.keys())).toEqual(
      expect.arrayContaining(['pdf-lib', 'fontkit'])
    );
  });

  test('a failure prints the chain from the entry to the import', () => {
    const entry = join(PACKAGES, 'react/src/index.ts');
    const middle = join(PACKAGES, 'core/src/editor/index.ts');
    const reach: Reach = {
      packages: new Map(),
      external: new Map(),
      unresolved: [],
      parent: new Map([
        [entry, null],
        [
          middle,
          { file: entry, line: 3, statement: "export * from '@pptx-editor.dev/core/editor';" },
        ],
      ]),
    };
    const site = { file: middle, line: 9, statement: "import * as Y from 'yjs';" };
    expect(
      reachFailure('@pptx-editor.dev/react', reach, [['yjs', site]], 'are forbidden', 'remove it.')
    ).toEqual([
      'A browser import of @pptx-editor.dev/react reaches 1 name(s) that are forbidden:',
      "  'yjs', through:",
      "    packages/react/src/index.ts:3  export * from '@pptx-editor.dev/core/editor';",
      "      packages/core/src/editor/index.ts:9  import * as Y from 'yjs';",
      'Fix: remove it.',
    ]);
  });
});
