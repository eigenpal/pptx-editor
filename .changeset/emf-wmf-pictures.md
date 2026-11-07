---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

EMF and WMF pictures draw as pictures, not as placeholders. The editor converts each one to a PNG in the background with `emf-converter`, which loads only for a deck that has a metafile. The placeholder shows until the PNG is ready, and then the editor repaints. The picture file stays unchanged.

- `session.picturesReady(index)` waits until the EMF and WMF pictures of a slide, or of every slide, have converted. It resolves `false` when a picture cannot convert or is not ready after 30 seconds. PNG and ZIP export wait for it. Call it before `exportSlideSvg()`.
- An OLE object whose preview is an EMF or WMF picture draws the preview.
- PDF export draws EMF and WMF pictures with the same converter. On Node.js it draws with `@napi-rs/canvas` when that package is installed, and with its own rasterizer otherwise, so the pixels can differ from the browser's.
- A WebP picture must pass the same size limits as other pictures before it draws.
- The natural size of an EMF or WMF picture comes from its picture frame.

Damaged and unusual files open and save more safely:

- A save keeps the grid columns of a table whose rows have more cells than columns. Extra cells with text stay in the file.
- A shape whose picture fill is missing or damaged draws the picture placeholder, not nothing.
- A damaged notes part, or a damaged theme of the notes master, drops the notes and no longer refuses the deck. A damaged theme of the handout master no longer refuses the deck.
- Edits, renames, deletes, and layout changes find every copy of a shape in nested `mc:AlternateContent`.
