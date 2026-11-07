#!/usr/bin/env node
// Mirrors the Vue adapter's source tree against the React adapter's, after DOCX Editor's
// scripts/check-adapter-mirror.mjs: every Vue source file has a React twin at the same relative
// path (`.ts` and `.tsx` count as one), unless the list below names it with a reason. The list
// only shrinks: an entry whose file is gone, or that now has a twin, fails as stale. Names
// compare by exact case, so a check on a case-insensitive file system agrees with CI.
//
// React files with no Vue twin are listed the same way (`reactOnly`). DOCX Editor checks only
// the Vue side; here both sides are small enough to hold to the same rule.

import { readdirSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const REACT_SRC = join(ROOT, 'packages/react/src');
const VUE_SRC = join(ROOT, 'packages/vue/src');
const PRO_REACT_SRC = join(ROOT, 'packages/pro/src/react');
const PRO_VUE_SRC = join(ROOT, 'packages/pro/src/vue');

const SKIP = [/(^|\/)__tests__\//];

/** Vue source files with no React twin at the mirrored path. */
export const VUE_ONLY = {
  'components.ts':
    'PptxEditorRoot, PptxEditorViewport, and PptxEditorContent; React holds them in PptxEditorRoot.tsx.',
  'composables.ts':
    'The Vue composables and the injection keys; React splits them over context.ts and the use*.ts hooks.',
  'form-draft.ts':
    'The draft state of the form dialogs as Vue refs; React keeps it in the dialog hooks of dialog-form.tsx.',
  'loading.ts': 'PptxEditorLoading; the React twin is Loading.tsx.',
  'menu-row.tsx': 'MenuRow and MenuSubmenu; React holds them in menu-rows.tsx.',
  'rail-memo.ts':
    'The rail item vnode cache with the v-memo mark; React uses the core cache in panes.tsx.',
  'ribbon.ts': 'PptxEditorRibbon and RibbonTab; the React twin is Ribbon.tsx.',
  'status-bar.ts': 'PptxEditorStatusBar and its parts; the React twin is StatusBar.tsx.',
  'toolbar-bar.tsx': 'The toolbar bar and its overflow; React holds them in Toolbar.tsx.',
  'toolbar-controls.tsx':
    'The registry controls (ToolbarControl and the shaped parts); React holds them in Toolbar.tsx.',
  'toolbar.ts':
    'ToolbarGroup and ToolbarSlot; React holds them in Toolbar.tsx and toolbar-parts.tsx.',
};

/** React source files with no Vue twin at the mirrored path. */
export const REACT_ONLY = {
  'chrome-actions.tsx':
    'The React chrome action host (dialogs, find bar, link popup); Vue holds it in composables.ts and popups.ts.',
  'context.ts': 'The React contexts and hooks; Vue holds them in composables.ts.',
  'dialog-exports.ts':
    'The dialog re-exports of the React entry, split out for the line cap; Vue exports its dialogs from index.ts.',
  'dialog-sections.ts':
    'A re-export of the core dialog-section requests; the Vue dialogs import them from core directly.',
  'Loading.tsx': 'The Vue twin is loading.ts.',
  'PptxEditorRoot.tsx': 'The Vue twin is components.ts.',
  'Ribbon.tsx': 'The Vue twin is ribbon.ts.',
  'StatusBar.tsx': 'The Vue twin is status-bar.ts.',
  'sugar-parts.ts':
    'Helpers of the React <PptxEditor> sugar component; the Vue sugar component holds them in index.ts.',
  'toolbar-action.tsx': 'ToolbarAction; Vue holds it in toolbar-parts.ts.',
  'toolbar-named-parts.tsx': 'The named toolbar parts; Vue holds them in toolbar-parts.ts.',
  'Toolbar.tsx':
    'The toolbar, its groups, and its controls; Vue splits them over toolbar.ts, toolbar-bar.tsx, and toolbar-controls.tsx.',
  'useBulletFormat.ts': 'A React hook; Vue holds useBulletFormat in composables.ts.',
  'useFontFormat.ts': 'A React hook; Vue holds useFontFormat in composables.ts.',
  'useParagraphFormat.ts': 'A React hook; Vue holds useParagraphFormat in composables.ts.',
};

function walk(dir, base = dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const abs = join(dir, name);
    const rel = relative(base, abs).split('\\').join('/');
    if (statSync(abs).isDirectory()) out.push(...walk(abs, base));
    else if (/\.(ts|tsx)$/.test(name) && !SKIP.some((re) => re.test(rel))) out.push(rel);
  }
  return out;
}

const stem = (rel) => rel.replace(/\.tsx?$/, '');

/** Every mirror error for the two file lists and the two exception lists. */
export function mirrorErrors({ reactFiles, vueFiles, vueOnly = VUE_ONLY, reactOnly = REACT_ONLY }) {
  const errors = [];
  const sides = [
    { label: 'Vue', files: vueFiles, other: 'React', otherFiles: reactFiles, allowed: vueOnly },
    { label: 'React', files: reactFiles, other: 'Vue', otherFiles: vueFiles, allowed: reactOnly },
  ];
  for (const { label, files, other, otherFiles, allowed } of sides) {
    const twins = new Set(otherFiles.map(stem));
    const present = new Set(files);
    for (const rel of files)
      if (!twins.has(stem(rel)) && !(rel in allowed))
        errors.push(`${label} file ${rel} has no ${other} twin and no listed reason.`);
    for (const [rel, reason] of Object.entries(allowed)) {
      if (!present.has(rel))
        errors.push(`${label}-only entry ${rel} names no ${label} file (stale).`);
      else if (twins.has(stem(rel)))
        errors.push(`${label}-only entry ${rel} now has a ${other} twin (stale).`);
      if (typeof reason !== 'string' || !reason.trim())
        errors.push(`${label}-only entry ${rel} needs a reason.`);
    }
  }
  return errors;
}

function main() {
  const reactFiles = walk(REACT_SRC);
  const vueFiles = walk(VUE_SRC);
  // The Pro entries for the adapters mirror each other with no exceptions.
  const proReact = walk(PRO_REACT_SRC);
  const proVue = walk(PRO_VUE_SRC);
  const errors = [
    ...mirrorErrors({ reactFiles, vueFiles }),
    ...mirrorErrors({ reactFiles: proReact, vueFiles: proVue, vueOnly: {}, reactOnly: {} }).map(
      (error) => `pro: ${error}`
    ),
  ];
  if (errors.length) {
    console.error('✘ adapter mirror:\n');
    for (const error of errors) console.error(`  - ${error}`);
    console.error(
      '\nAdd the twin, or list the file in scripts/check-adapter-mirror.mjs with a reason.'
    );
    process.exit(1);
  }
  const vueOnly = Object.keys(VUE_ONLY).length;
  const reactOnly = Object.keys(REACT_ONLY).length;
  console.log(
    `✓ adapter mirror: ${vueFiles.length} Vue and ${reactFiles.length} React files, and ${proVue.length} and ${proReact.length} in Pro (${vueOnly} Vue-only and ${reactOnly} React-only, each with a reason).`
  );
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) main();
