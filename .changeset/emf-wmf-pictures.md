---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Draw EMF and WMF pictures. Add `session.picturesReady`, `session.pictureStatus`, and the `picturesTimeoutMs` option of PNG and ZIP export. Call `session.picturesReady` before `exportSlideSvg()`. On Node.js, PDF export uses `@napi-rs/canvas` when installed. Open and save damaged files more safely.
