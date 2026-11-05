// The published packages and their public entry points. One list for api-extractor,
// notices, and release checks (docx-editor keeps the same list in scripts/lib/packages.mjs).
export const PACKAGES = [
  { name: '@pptx-editor.dev/i18n', root: 'packages/i18n', slug: 'i18n', entries: ['index'] },
  {
    name: '@pptx-editor.dev/core',
    root: 'packages/core',
    slug: 'core',
    entries: ['index', 'contracts', 'store', 'cascade', 'layout', 'automation', 'output', 'editor'],
  },
  { name: '@pptx-editor.dev/react', root: 'packages/react', slug: 'react', entries: ['index'] },
  { name: '@pptx-editor.dev/vue', root: 'packages/vue', slug: 'vue', entries: ['index'] },
  { name: '@pptx-editor.dev/pro', root: 'packages/pro', slug: 'pro', entries: ['index'] },
  {
    name: '@pptx-editor.dev/pptx-to-pdf',
    root: 'packages/pptx-to-pdf',
    slug: 'pptx-to-pdf',
    entries: ['index'],
  },
  {
    name: '@pptx-editor.dev/editor-api',
    root: 'packages/editor-api',
    slug: 'editor-api',
    entries: ['index', 'browser'],
  },
];
