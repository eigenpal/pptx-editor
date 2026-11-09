// The workspace source aliases of the collaboration examples, so `bun install` and a dev server
// run them from a clean clone with no package build step, as the Vite demo does.

import { fileURLToPath } from 'node:url';
import type { Alias } from 'vite';

const pkg = (p: string): string =>
  fileURLToPath(new URL(`../../../packages/${p}`, import.meta.url));

export const sourceAliases: Alias[] = [
  {
    find: /^@pptx-editor\.dev\/core\/styles\/editor\.css$/,
    replacement: pkg('core/src/styles/editor.css'),
  },
  { find: /^@pptx-editor\.dev\/core\/(\w+)$/, replacement: pkg('core/src/$1/index.ts') },
  { find: /^@pptx-editor\.dev\/core$/, replacement: pkg('core/src/index.ts') },
  { find: /^@pptx-editor\.dev\/react$/, replacement: pkg('react/src/index.ts') },
  { find: /^@pptx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
  { find: /^@pptx-editor\.dev\/pro\/drawing$/, replacement: pkg('pro/src/drawing/index.ts') },
  { find: /^@pptx-editor\.dev\/pro$/, replacement: pkg('pro/src/index.ts') },
];

/** The sample decks of the Vite demo, served as the example's public folder. */
export const sharedPublicDir = fileURLToPath(new URL('../../vite/public', import.meta.url));
