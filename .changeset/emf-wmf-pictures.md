---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

EMF and WMF pictures draw as pictures, not as placeholders. The editor converts each one to a PNG in the background with `emf-converter`, which loads only for a deck that has a metafile. The placeholder shows until the PNG is ready, and then the editor repaints. The picture file stays unchanged.

- `session.picturesReady()` resolves when every EMF and WMF picture has converted. PNG and ZIP export wait for it. Call it before `exportSlideSvg()`.
- PDF export draws EMF and WMF pictures with the same conversion.
- A WebP picture must pass the same size limits as other pictures before it draws.
- The natural size of an EMF or WMF picture comes from its picture frame.

Damaged and unusual files open and save more safely:

- A save keeps the grid columns of a table whose rows have more cells than columns. Extra cells with text stay in the file.
- A shape whose picture fill is missing or damaged draws the picture placeholder, not nothing.
- A damaged notes part, or a damaged theme of the notes master or handout master, drops the notes and no longer refuses the deck.
- Edits, renames, deletes, and layout changes find every copy of a shape in nested `mc:AlternateContent`.
