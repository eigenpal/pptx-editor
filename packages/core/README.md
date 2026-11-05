<p align="center">
  <a href="https://pptx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" alt="PPTX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/core"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/core.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/core"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/core.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=C43E1C" alt="license" /></a>
</p>

# @pptx-editor.dev/core

The PPTX Editor engine. It reads a `.pptx` into one document store, resolves PresentationML inheritance (master, layout, and slide), lays out slides without the DOM, paints them as SVG, applies edits as serializable ops, and saves the package.

## Install

```bash
npm install @pptx-editor.dev/core
```

## Entry points

| Entry | What it holds |
| --- | --- |
| `@pptx-editor.dev/core/editor` | `createPptxEditor`, the editor facade, commands, and the chrome registry |
| `@pptx-editor.dev/core/store` | Bounded OPC read and write, the canonical tree, ops, and the store |
| `@pptx-editor.dev/core/cascade` | Master, layout, and slide inheritance, theme colors and fonts |
| `@pptx-editor.dev/core/layout` | DOM-free geometry and text layout |
| `@pptx-editor.dev/core/output` | SVG paint |
| `@pptx-editor.dev/core/contracts` | Shared types |
| `@pptx-editor.dev/core/styles/editor.css` | Chrome tokens and styles |

## Use

```ts
import { createPptxEditor } from '@pptx-editor.dev/core/editor';

const editor = createPptxEditor();
await editor.load(bytes);
editor.exec({ type: 'toggleBold' });
const saved = editor.save(); // Uint8Array, or null when no document is open
```

For a React UI, use [`@pptx-editor.dev/react`](../react/README.md).

## Fonts

`createPptxEditor({ fonts })` takes a `FontConfiguration`, a fragment, or a `FontResolver`, with the names and shapes of DOCX Editor's API: `loadFonts`, `createFontSource`, `customFonts`, `composeFontConfiguration`, `defineFontResolver`, and `deferFonts`. `editor.setFonts()` replaces the host faces in place, and `onFontError` receives each failure. Without `fonts`, the editor uses the deck's embedded fonts and the packaged substitutes, and makes no network request. See [Fonts](../../docs/site/content/customization/fonts.mdx).

## Bundlers

Vite, webpack, and Turbopack emit the text shaper's `harfbuzz.wasm` with no configuration. Bun and esbuild do not. With them, serve `@pptx-editor.dev/core/harfbuzz.wasm` as a static file, and set its URL before you create the first editor:

```ts
import { setHarfBuzzWasmUrl } from '@pptx-editor.dev/core/layout';

setHarfBuzzWasmUrl('/static/harfbuzz.wasm');
```

## License

Apache-2.0. See [LICENSE](LICENSE).
