// Find the local DOCX Editor checkout that the scripts in this folder read. Both scripts run on
// a developer machine only: CI has no DOCX Editor checkout.

import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '../..');

/** A file that every DOCX Editor checkout holds, so a folder of another repository is refused. */
const MARKER = 'packages/core/src/store/package/zip.ts';

/** The result of a search: a checkout, a named folder that is not one, or nothing found. */
export type DocxCheckout =
  | { readonly kind: 'found'; readonly dir: string }
  | { readonly kind: 'invalid'; readonly message: string }
  | { readonly kind: 'absent' };

/** The PPTX Editor main checkout, also when a script runs in a worktree of it. */
function mainCheckout(): string {
  try {
    const common = execFileSync(
      'git',
      ['-C', ROOT, 'rev-parse', '--path-format=absolute', '--git-common-dir'],
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }
    ).trim();
    return dirname(common);
  } catch {
    return ROOT;
  }
}

/**
 * The DOCX Editor checkout: `named` (from `--docx`), then `$DOCX_EDITOR_DIR`, then
 * `~/Developer/docx-editor`, `~/docx-editor`, and the `docx-editor` folder next to the main
 * PPTX Editor checkout. A folder named by the option or the environment must be a checkout:
 * there is no silent fallback.
 */
export function findDocxCheckout(named: string | undefined): DocxCheckout {
  const isCheckout = (dir: string): boolean => existsSync(join(dir, MARKER));
  for (const [dir, from] of [
    [named, '--docx'],
    [process.env.DOCX_EDITOR_DIR || undefined, 'DOCX_EDITOR_DIR'],
  ] as const) {
    if (dir === undefined) continue;
    if (!isCheckout(dir))
      return { kind: 'invalid', message: `${from}: ${dir} is not a DOCX Editor checkout.` };
    return { kind: 'found', dir: resolve(dir) };
  }
  const candidates = [
    join(homedir(), 'Developer/docx-editor'),
    join(homedir(), 'docx-editor'),
    join(mainCheckout(), '../docx-editor'),
  ];
  for (const dir of candidates) if (isCheckout(dir)) return { kind: 'found', dir: resolve(dir) };
  return { kind: 'absent' };
}
