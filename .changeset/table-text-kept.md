---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Fix collaborative table edits that lost text no peer removed. Merged text now survives concurrent row and column deletes, undo of an insert, and merges that overlap, and a covered cell shows text that a peer typed into it. A table command no longer adds a text body to a cell that has none, so its undo keeps a peer's typing.
