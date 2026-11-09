---
'@pptx-editor.dev/pro': patch
---

Make a collaborative undo or redo change only what its own edit wrote. An undo of a format or an Enter no longer brings back text that a peer deleted, before or at the same time, and an undo no longer removes text that a peer typed into the edit's new runs or paragraphs. An undo of typing also deletes a peer's copy of that text made at the same time.
