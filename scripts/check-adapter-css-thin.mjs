#!/usr/bin/env node
// Copied from DOCX Editor (docx-editor/scripts/check-adapter-css-thin.mjs), with the PPTX
// package names. Guards the single-source-of-truth rule for chrome styling: each adapter's
// `src/styles/editor.css` only imports the shared core stylesheet. No tokens and no UI rules
// may fork into one adapter. If this fails, move the style into
// packages/core/src/styles/editor.css.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const CORE_IMPORT = "@import '@pptx-editor.dev/core/styles/editor.css';";

/** What is left of an adapter stylesheet after comments and the core import, trimmed. */
export function residual(css) {
  return css
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/@import\s+['"]@pptx-editor\.dev\/core\/styles\/editor\.css['"];?/g, '')
    .trim();
}

/** Every failure for the given stylesheets, by path. */
export function cssThinErrors(files) {
  const errors = [];
  for (const [file, css] of Object.entries(files)) {
    if (!css.includes(CORE_IMPORT)) errors.push(`${file}: must ${CORE_IMPORT}`);
    const left = residual(css);
    if (left)
      errors.push(`${file}: adapter-local styling must live in the core stylesheet:\n    ${left}`);
  }
  return errors;
}

function main() {
  const paths = ['packages/react/src/styles/editor.css', 'packages/vue/src/styles/editor.css'];
  const errors = cssThinErrors(
    Object.fromEntries(paths.map((p) => [p, readFileSync(join(root, p), 'utf8')]))
  );
  if (errors.length) {
    console.error('✘ adapter editor.css is not thin:\n');
    for (const error of errors) console.error(`  - ${error}`);
    process.exit(1);
  }
  console.log('✓ adapter editor.css files are thin (shared styling lives in core).');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
