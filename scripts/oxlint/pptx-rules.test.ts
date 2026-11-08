// The lint gates that AGENTS.md relies on must keep firing: the `max-lines` caps, the
// HTML-sink ban, global `@keyframes` names, varargs spreads in the DOM-free lanes, the React/Vue
// import isolation, and the ESLint-era bans (`require`, `any`, `console`, `eval`). The repository
// has no violations, so a clean lint run cannot tell a working rule from one that silently
// stopped running. oxlint's JavaScript plugins are alpha and outside semver, so this test lints a
// fixture of every banned shape with the real `.oxlintrc.json`, at the paths its overrides name,
// and checks the exact reports. Same approach as DOCX Editor's
// `scripts/__tests__/oxlint-rules.test.ts`. Adapted from DOCX Editor at commit adf0e178.

import { afterAll, beforeAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import ts from 'typescript';

const root = join(import.meta.dir, '..', '..');
let workspace = '';

const PRELUDE = [
  'declare const el: HTMLElement;',
  'declare const popup: Window;',
  'declare const items: number[];',
  'declare const out: number[];',
  'declare function f(...a: number[]): void;',
].join('\n');
// Lines 6-12 are banned in shipped code.
const SINKS = [
  "el.innerHTML = '<b>x</b>';",
  "el.outerHTML = '<b>x</b>';",
  "el.insertAdjacentHTML('beforeend', '<b>x</b>');",
  "document.write('<b>x</b>');",
  "popup.document.write('<b>x</b>');",
  'export const slide = `@keyframes slideIn { from { opacity: 0 } }`;',
  "export const fade = '@keyframes fadeOut { }';",
].join('\n');
const ALLOWED = ['export const ok = `@keyframes pptx-fade { }`;', 'el.textContent = "x";'].join(
  '\n'
);
// Lines 6-10 are banned in the DOM-free lanes.
const SPREADS = [
  'f(...items);',
  'export const s = new Set(...items);',
  'out.push(...items);',
  'out.splice(0, 0, ...items);',
  'Math.max(...items);',
].join('\n');
// All four are banned in shipped code. Tests and scripts may use `any` and the console.
const MISC = [
  "export const m = require('node:fs');",
  'export const a: any = 1;',
  "export const e = eval('1');",
  "console.log('x');",
].join('\n');
const LONG = Array.from({ length: 1001 }, (_, i) => `export const x${i} = ${i};`).join('\n');

const LANES = [
  'packages/core/src/store',
  'packages/core/src/cascade',
  'packages/core/src/layout',
  'packages/core/src/automation',
  'packages/ooxml-core/src',
];

const FILES: Record<string, string> = {
  // Shipped code outside the lanes: every sink and keyframe name, no spread ban.
  'packages/core/src/editor/sample.ts': `${PRELUDE}\n${SINKS}\n${ALLOWED}\n${SPREADS}\n`,
  'packages/core/src/editor/misc.ts': `${MISC}\n`,
  // Tests are exempt from the sink, keyframes, spread, `any`, and console rules.
  'packages/core/src/store/__tests__/sample.test.ts': `${PRELUDE}\n${SINKS}\n${SPREADS}\n${MISC}\n`,
  // Scripts may build HTML and use `any`; the console is theirs.
  'scripts/sample.ts': `${PRELUDE}\n${SINKS}\n${MISC}\n`,
  // A React element with the sink attribute.
  'packages/react/src/sink.tsx':
    "export const x = <div dangerouslySetInnerHTML={{ __html: '' }} />;\n",
  // Framework isolation.
  'packages/react/src/iso.ts': "import { ref } from 'vue';\nexport { ref };\n",
  'packages/vue/src/iso.ts': [
    "import { useState } from 'react';",
    "import { createRoot } from 'react-dom/client';",
    "import { PptxEditor } from '@pptx-editor.dev/react';",
    'export { useState, createRoot, PptxEditor };',
  ].join('\n'),
  // A single-file component: the JavaScript plugin runs on its script block.
  'examples/vue/src/Sample.vue': [
    '<template><div /></template>',
    '<script setup lang="ts">',
    "document.body.innerHTML = '<b>x</b>';",
    '</script>',
    '',
  ].join('\n'),
  // The global cap, in shipped code and in scripts.
  'packages/core/src/editor/long.ts': `${LONG}\n`,
  'scripts/long.mjs': `${LONG}\n`,
};
for (const lane of LANES) FILES[`${lane}/spread.ts`] = `${PRELUDE}\n${SPREADS}\n`;

/** The rules under test; everything else in the fixture is noise. */
const CHECKED =
  /^(pptx\(|eslint\((no-restricted-imports|no-restricted-globals|max-lines|no-eval|no-console)\)|typescript\(no-explicit-any\))/;

function reports(): string[] {
  const result = spawnSync(
    join(root, 'node_modules', '.bin', 'oxlint'),
    ['-f', 'json', '-c', join(workspace, '.oxlintrc.json'), 'packages', 'scripts', 'examples'],
    { cwd: workspace, encoding: 'utf8' }
  );
  if (result.error) throw result.error;
  // Exit 1 is expected: the fixture has errors. Anything on stderr (a plugin that failed to
  // load, an invalid config) is the real failure.
  if (result.stderr.trim()) throw new Error(`oxlint failed:\n${result.stderr}`);
  const { diagnostics } = JSON.parse(result.stdout) as {
    diagnostics: { code: string; filename: string; labels: { span: { line: number } }[] }[];
  };
  const found = diagnostics
    .filter((d) => CHECKED.test(d.code ?? ''))
    .map((d) => `${d.filename}:${d.labels[0]?.span.line} ${d.code}`);
  return [...new Set(found)].sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
}

beforeAll(() => {
  // Real path: override globs match paths relative to the config, and on macOS the temporary
  // directory is a symlink that oxlint resolves.
  workspace = realpathSync(mkdtempSync(join(tmpdir(), 'pptx-oxlint-rules-')));
  const parsed = ts.parseConfigFileTextToJson(
    '.oxlintrc.json',
    readFileSync(join(root, '.oxlintrc.json'), 'utf8')
  );
  if (parsed.error) {
    throw new Error(ts.flattenDiagnosticMessageText(parsed.error.messageText, '\n'));
  }
  const config = parsed.config as { jsPlugins: string[]; $schema?: string };
  config.jsPlugins = config.jsPlugins.map((path) => join(root, path));
  delete config.$schema;
  writeFileSync(join(workspace, '.oxlintrc.json'), JSON.stringify(config));
  for (const [path, text] of Object.entries(FILES)) {
    mkdirSync(join(workspace, dirname(path)), { recursive: true });
    writeFileSync(join(workspace, path), text);
  }
});

afterAll(() => rmSync(workspace, { recursive: true, force: true }));

const sinks = (file: string) => [
  ...[6, 7, 8, 9, 10].map((line) => `${file}:${line} pptx(no-html-sinks)`),
  ...[11, 12].map((line) => `${file}:${line} pptx(no-global-keyframes)`),
];
const spreads = (file: string) =>
  [6, 7, 8, 9, 10].map((line) => `${file}:${line} pptx(no-varargs-spread)`);

test('every banned shape is reported where its rule applies, and only there', () => {
  const expected = [
    ...sinks('packages/core/src/editor/sample.ts'),
    'packages/core/src/editor/misc.ts:1 eslint(no-restricted-globals)',
    'packages/core/src/editor/misc.ts:2 typescript(no-explicit-any)',
    'packages/core/src/editor/misc.ts:3 eslint(no-eval)',
    'packages/core/src/editor/misc.ts:4 eslint(no-console)',
    'packages/core/src/store/__tests__/sample.test.ts:18 eslint(no-restricted-globals)',
    'packages/core/src/store/__tests__/sample.test.ts:20 eslint(no-eval)',
    'scripts/sample.ts:13 eslint(no-restricted-globals)',
    'scripts/sample.ts:15 eslint(no-eval)',
    'packages/react/src/sink.tsx:1 pptx(no-html-sinks)',
    'packages/react/src/iso.ts:1 eslint(no-restricted-imports)',
    'packages/vue/src/iso.ts:1 eslint(no-restricted-imports)',
    'packages/vue/src/iso.ts:2 eslint(no-restricted-imports)',
    'packages/vue/src/iso.ts:3 eslint(no-restricted-imports)',
    'examples/vue/src/Sample.vue:3 pptx(no-html-sinks)',
    'packages/core/src/editor/long.ts:1001 eslint(max-lines)',
    'scripts/long.mjs:1001 eslint(max-lines)',
    ...LANES.flatMap((lane) => spreads(`${lane}/spread.ts`)),
  ];
  const actual = reports();
  // max-lines reports at the first line past the cap or at the file start, depending on the
  // version; compare the rule and the file for it.
  const normalize = (entry: string) =>
    entry.replace(/:\d+ eslint\(max-lines\)$/, ' eslint(max-lines)');
  expect(actual.map(normalize).sort()).toEqual(expected.map(normalize).sort());
});
