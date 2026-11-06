// The internal lane DAG of the core package, adapted from
// docx-editor/packages/core/src/__tests__/core-lane-graph.ts.
//
// The machine-readable source of truth for which lane may import which. A lane that acquires a
// new dependency declares it here, in a diff a reviewer sees. `core-lane-graph.test.ts` proves
// the graph is acyclic and that real imports follow it; `core-lane-imports.test.ts` checks
// every import statement against it, and `browser-bundle-graph.test.ts` walks the browser
// entry points against `BROWSER_FORBIDDEN_DEPENDENCIES`.

import { existsSync } from 'node:fs';
import { join, resolve } from 'node:path';

export type LaneName =
  'contracts' | 'store' | 'cascade' | 'layout' | 'automation' | 'output' | 'editor';

export type LaneEnvironment = 'neutral' | 'browser';

export interface Lane {
  readonly directory: string;
  readonly mayImport: readonly LaneName[];
  /**
   * Lanes this one may import TYPES from, in addition to {@link Lane.mayImport}.
   *
   * Type-only imports (`import type`, `export type ... from`) are erased at build, so they can
   * never carry code or an environment across lanes. This is never a value edge: a value import
   * from one of these lanes still fails, and the cycle walk ignores it. No lane declares one
   * today; an off-DAG type edge is pinned in `core-lane-imports.test.ts` instead.
   */
  readonly mayImportTypes?: readonly LaneName[];
  /** `neutral` must run in Bun, Node, and browsers: no DOM. */
  readonly environment: LaneEnvironment;
  readonly subpath: string;
}

export const CORE_LANES: Readonly<Record<LaneName, Lane>> = Object.freeze({
  // Types only, but the editor contract (`contracts/editor.ts`) names DOM types
  // (`attach(element: HTMLElement)`), so the lane is not compiled DOM-free as a whole. DOCX
  // Editor leaves its contracts out of the neutral set too. The DOM-free lanes import
  // `contracts/neutral.ts`, which `check:lane-boundaries` compiles without the DOM library.
  contracts: {
    directory: 'src/contracts',
    mayImport: [],
    environment: 'browser',
    subpath: './contracts',
  },
  store: { directory: 'src/store', mayImport: [], environment: 'neutral', subpath: './store' },
  cascade: {
    directory: 'src/cascade',
    mayImport: ['store'],
    environment: 'neutral',
    subpath: './cascade',
  },
  // `contracts` for the font configuration types only (`contracts/editor-fonts.ts`, DOM-free),
  // as DOCX Editor's layout lane reads them for font composition.
  layout: {
    directory: 'src/layout',
    mayImport: ['contracts', 'store', 'cascade'],
    environment: 'neutral',
    subpath: './layout',
  },
  // The DOM-free command lane, after DOCX Editor's `core/src/automation`: the edits that lower
  // to store ops (`commands/`) and the automation host that batches them. `contracts` for the
  // command and selection types the edits take.
  automation: {
    directory: 'src/automation',
    mayImport: ['contracts', 'store', 'cascade', 'layout'],
    environment: 'neutral',
    subpath: './automation',
  },
  output: {
    directory: 'src/output',
    mayImport: ['store', 'cascade', 'layout'],
    environment: 'browser',
    subpath: './output',
  },
  editor: {
    directory: 'src/editor',
    mayImport: ['contracts', 'store', 'cascade', 'layout', 'automation', 'output'],
    environment: 'browser',
    subpath: './editor',
  },
});

/** Every lane, dependencies first. Throws on a cycle. */
export function laneTopologicalOrder(): LaneName[] {
  const order: LaneName[] = [];
  const visiting = new Set<LaneName>();
  const visit = (name: LaneName): void => {
    if (order.includes(name)) return;
    if (visiting.has(name)) throw new Error(`Lane cycle through ${name}`);
    visiting.add(name);
    for (const dependency of CORE_LANES[name].mayImport) visit(dependency);
    visiting.delete(name);
    order.push(name);
  };
  for (const name of Object.keys(CORE_LANES) as LaneName[]) visit(name);
  return order;
}

/**
 * Third-party dependencies and Node builtins that must NOT reach a default browser import.
 * `yjs` and `y-protocols` belong to the pro collaboration module; `pdf-lib` and `fontkit` to
 * `pptx-to-pdf`; `pdfkit` is DOCX Editor's PDF stack, kept so the two lists stay one list.
 */
export const BROWSER_FORBIDDEN_DEPENDENCIES: readonly string[] = [
  'yjs',
  'y-protocols',
  'pdfkit',
  'pdf-lib',
  'fontkit',
  'node:fs',
  'node:net',
  'node:http',
  // `@pptx-editor.dev/pptx-to-pdf` depends on the pro package for its drawing renderers.
  '@pptx-editor.dev/pro',
];

/** The core package root (`packages/core`). */
export const CORE_PACKAGE_ROOT = resolve(import.meta.dir, '..', '..');

/** Throws when a lane directory is missing, so a moved lane cannot pass vacuously. */
export function existingLanePath(lane: LaneName): string {
  const p = join(CORE_PACKAGE_ROOT, CORE_LANES[lane].directory);
  if (!existsSync(p)) throw new Error(`lane ${lane} directory ${p} does not exist`);
  return p;
}
