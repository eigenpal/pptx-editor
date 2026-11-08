---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

`usePptxSource` stops stalled downloads with a `TimeoutError` and oversized downloads with a `RangeError`. Fonts hold the open for at most 2 seconds, and `registerFontFaces` takes `waitMs`.
