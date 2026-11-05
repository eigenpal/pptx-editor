# @pptx-editor.dev/pptx-to-pdf

## 0.1.0

### Minor Changes

- 457a754: East Asian text measures and draws with each run's `a:ea` font (complex-script text with `a:cs`, symbol characters with `a:sym`), and a theme with an empty `a:ea` uses its script font for the run's language. `eaLnBrk="0"` now breaks East Asian text between any two characters with no kinsoku rule, `latinLnBrk="1"` breaks inside Latin words, Korean text breaks at spaces, `hangingPunct` lets a comma or full stop hang past the line end, the presentation's `p:kinsoku` lists replace the built-in lists, and justified East Asian lines spread between characters, as the reference PDFs show. A missing East Asian font falls back to a host CJK face for the text's language, such as Noto Sans CJK JP, and a character that the run's face lacks uses another loaded face. The canvas shows the text that an input method composes at the caret, and the key that ends a composition runs no shortcut. `@pptx-editor.dev/i18n` adds Japanese (`/ja`), Korean (`/ko`), and Simplified Chinese (`/zh-CN`) catalogs, and `@pptx-editor.dev/pptx-to-pdf` draws East Asian text with Noto Sans CJK JP when the optional `@docx-editor.dev/fonts-cjk` is installed.
- 412d750: Pictures can be cropped, changed, reset, adjusted, styled, sized, and given alt text through new commands, the toolbar's `image` group, and the `imageAltText` and `imageProperties` popups, and picture borders now draw outside the picture.
- 412d750: `openDocumentForExport` and `exportPdfFrom` export one deck more than once from one layout, `ExportResourceError` reports a disposed or aborted session, each PDF result maps its pages to deck slide numbers in `slideNumbers`, and an engine failure inside an export is reported as `PdfEncodingError`.
- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.

### Patch Changes

- 412d750: Draw shadow blur, glow, and soft edges at their measured size, put an inner shadow on the side its direction names, and keep a shape's outline inside its effects.
- 412d750: An embedded OLE object now shows the PNG, JPEG, GIF, or BMP picture of its `p:oleObj/p:pic` in the frame's box, in place of a labeled placeholder.
- 412d750: The PDF export draws scaled linear gradients, circle path gradients, and two-stop gradient colors with the same rules as the editor.
- 412d750: `pptx-to-pdf` now draws cropped pictures, linear and path gradients (gradient text too), tables, outer and inner shadows, and glow, and reports each approximation with a diagnostic code; the codes `effects`, `picture.crop`, `path.fill-mode`, `fill.linear`, `fill.path`, and `fill.group` are no longer emitted, and `frame.table` only for a table frame without a grid. Many effects or gradients on one page now stay inside the paint budget and `timeoutMs`.
- 412d750: Clarify package documentation and host responsibilities for collaboration access.
