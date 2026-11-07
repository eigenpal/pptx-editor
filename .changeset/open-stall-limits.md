---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

A presentation that does not arrive no longer keeps the editor on the loading screen.

- `usePptxSource` stops a URL download when no bytes arrive for 10 seconds, and `error` is a `TimeoutError`. A slow download that keeps sending bytes finishes, however long it takes.
- `usePptxSource` stops a URL download that grows past 2 GiB, the largest archive the engine reads, and `error` is a `RangeError`.
- Eager fonts hold the deck back for at most 2 seconds. Fonts that settle later still apply, and the slides lay out again.
- `registerFontFaces` settles `ready` after at most 2 seconds (`waitMs`), so a face whose load never settles cannot hold the open.
