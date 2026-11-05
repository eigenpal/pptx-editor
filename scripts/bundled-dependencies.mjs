// The third-party code a package build inlines into its published files.
//
// A devDependency never reaches a consumer through the package manager, but a dependency that
// the build inlines (tsup `noExternal`) ships inside `dist/`. The shipped advisory scan reads
// this list so that inlined code, such as the HarfBuzz glue and WASM in core, is scanned too.

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * Package names in the `noExternal` list of `<packageDir>/tsup.config.ts`. A package without
 * a tsup config or without `noExternal` inlines nothing. An entry that is not a plain string
 * (a regular expression, a variable) throws, so a scan never runs on a partial list.
 */
export function bundledDependencies(packageDir) {
  const config = join(packageDir, 'tsup.config.ts');
  if (!existsSync(config)) return [];
  const match = /\bnoExternal\s*:\s*\[([^\]]*)\]/.exec(readFileSync(config, 'utf8'));
  if (!match) return [];
  const entries = match[1]
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean);
  return entries.map((entry) => {
    const literal = /^(['"])([^'"]+)\1$/.exec(entry);
    if (!literal) {
      throw new Error(
        `${config}: noExternal entry ${entry} is not a package name string; update scripts/bundled-dependencies.mjs.`
      );
    }
    return literal[2];
  });
}
