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
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/i18n"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/i18n.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/i18n"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/i18n.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=C43E1C" alt="license" /></a>
</p>

# @pptx-editor.dev/i18n

Shared user-facing strings for the PPTX Editor chrome. `en.json` is the source of truth. Other locales mirror its shape; a `null` value falls back to English.

## Install

The React and Vue adapters install this package for you. Install it yourself only when you import a catalog or a type from it directly:

```bash
npm install @pptx-editor.dev/i18n
```

## Use

```ts
import { createTranslator } from '@pptx-editor.dev/i18n';

const t = createTranslator();
t('toolbar.bold'); // "Bold (Ctrl+B)"
t('rail.slideLabel', { index: 3 }); // "Slide 3"
```

## Packaged catalogs

`@pptx-editor.dev/i18n/ja` (Japanese), `@pptx-editor.dev/i18n/ko` (Korean), and `@pptx-editor.dev/i18n/zh-CN` (Simplified Chinese) are packaged catalogs. A `null` value shows English.

```ts
import ja from '@pptx-editor.dev/i18n/ja';

const t = createTranslator(ja);
t('toolbar.bold'); // "太字 (Ctrl+B)"
```
