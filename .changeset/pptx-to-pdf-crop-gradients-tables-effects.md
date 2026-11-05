---
'@pptx-editor.dev/pptx-to-pdf': patch
'@pptx-editor.dev/core': patch
---

`pptx-to-pdf` now draws cropped pictures, linear and path gradients (gradient text too), tables, outer and inner shadows, and glow, and reports each approximation with a diagnostic code; the codes `effects`, `picture.crop`, `path.fill-mode`, `fill.linear`, `fill.path`, and `fill.group` are no longer emitted, and `frame.table` only for a table frame without a grid. Many effects or gradients on one page now stay inside the paint budget and `timeoutMs`.
