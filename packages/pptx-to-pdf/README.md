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
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/pptx-to-pdf"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/pptx-to-pdf.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/pptx-to-pdf"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/pptx-to-pdf.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/packages/pptx-to-pdf/LICENSE.md"><img src="https://img.shields.io/badge/license-EigenPal_Pro_License-blue.svg?style=flat-square&color=C43E1C" alt="EigenPal Pro License" /></a>
</p>

# @pptx-editor.dev/pptx-to-pdf

Convert PPTX files to searchable PDFs on Node.js using the editor's layout engine. The package uses the [EigenPal Pro License](LICENSE.md). Production use requires a commercial agreement.

## Install

Install the exporter and its engine peer:

```bash
npm install @pptx-editor.dev/pptx-to-pdf @pptx-editor.dev/core
```

Use Node.js 20.16 or later within version 20, or Node.js 22.3 or later. Keep this package out of browser bundles.

## Export a presentation

This example loads open font fallbacks and writes a PDF with conversion diagnostics:

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { exportPdf } from '@pptx-editor.dev/pptx-to-pdf';

const result = await exportPdf(await readFile('deck.pptx'), {
  googleFonts: true,
  fidelityPolicy: 'best-effort',
});
for (const diagnostic of result.diagnostics) {
  if (diagnostic.severity !== 'information') {
    console.warn(diagnostic.code, diagnostic.slideNumber);
  }
}
await writeFile('deck.pdf', result.bytes);
```

`googleFonts` enables network requests through a pinned, hash-verified catalog. It supplies Amiri for Arabic text and regional Noto Sans faces for CJK text. Supplied fonts and usable embedded faces keep precedence. Substitution can change line breaks. Omit this option for offline exports and supply bytes through `fonts`.

The default fidelity policy is `strict`, which throws `PdfFidelityError` for unsupported or approximate content. Use `best-effort` only when you can accept the reported differences. See [PDF fonts](https://pptx-editor.dev/docs/export/pdf/fonts) for offline assets and embedding permissions.

## Reuse a layout

Open one session to export the same presentation with different page options:

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { exportPdfFrom, openDocumentForExport } from '@pptx-editor.dev/pptx-to-pdf';

const opened = await openDocumentForExport(await readFile('deck.pptx'), {
  googleFonts: true,
});
if (!opened.ok) throw new Error(`Cannot open: ${opened.reason}`);
try {
  const result = await exportPdfFrom(opened.session, {
    fidelityPolicy: 'best-effort',
    includeHiddenSlides: true,
  });
  console.log(result.diagnostics);
  await writeFile('all-slides.pdf', result.bytes);
} finally {
  opened.session.dispose();
}
```

A session holds immutable presentation bytes and resolves fonts once. Set font options when opening it. Set output limits and page options on each export. Open another session after editing the presentation.

## Next steps

- [PDF integrations](https://pptx-editor.dev/docs/export/pdf/integrations): return HTTP responses and manage sessions.
- [PDF API](https://pptx-editor.dev/docs/export/pdf/api): options, diagnostics, limits, cancellation, and errors.
- [Feature support](https://pptx-editor.dev/docs/pptx-fidelity): rendering limits and measured fidelity.
- [Third-party notices](THIRD_PARTY_NOTICES.md): dependency licenses and required attribution.
