#!/usr/bin/env node
// Vue composables declare a named `Use<Name>Return` type (AGENTS.md "Public API"), so core's
// internal types do not leak into the Vue API snapshot. Every Vue composable must also have a
// React hook of the same name, or be listed in the parity contract's `vueExclusive`. After
// DOCX Editor's scripts/check-composable-parity.mjs. Reads the committed snapshots.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { composables, exportedNames } from './parity/api-snapshot.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

/** Every composable error for the given snapshots and contract. */
export function composableErrors({ reactApi, vueApi, contract }) {
  const errors = [];
  const reactExports = exportedNames(reactApi);
  const vueExports = exportedNames(vueApi);
  const exclusive = contract.exports?.vueExclusive ?? {};
  for (const { name, returns } of composables(vueApi)) {
    const expected = `U${name.slice(1)}Return`;
    if (returns !== expected)
      errors.push(
        `${name} must declare its return type as ${expected} (found ${returns ?? 'none'}).`
      );
    else if (!vueExports.has(expected))
      errors.push(`${expected} is not exported from the Vue adapter.`);
    if (!reactExports.has(name) && !(name in exclusive))
      errors.push(`${name} has no React hook of the same name and is not vueExclusive.`);
  }
  return errors;
}

function main() {
  const read = (path) => readFileSync(join(root, path), 'utf8');
  const errors = composableErrors({
    reactApi: read('docs/api/react/index.api.md'),
    vueApi: read('docs/api/vue/index.api.md'),
    contract: JSON.parse(read('scripts/parity/parity.contract.json')),
  });
  if (errors.length) {
    console.error('✘ composable parity:\n');
    for (const error of errors) console.error(`  - ${error}`);
    process.exit(1);
  }
  console.log('✓ Vue composables declare Use<Name>Return types and pair with React hooks.');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
