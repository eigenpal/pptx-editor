---
'@pptx-editor.dev/pro': minor
---

The Yjs session keeps more text that peers write at the same time.

- Text that a peer types into a paragraph that another peer removes at the same time, with a delete across paragraphs or a join, stays. The paragraph shows again at its old place with only that text, in shapes and in table cells. An undo of the removal puts the rest of the paragraph back.
- A run that a peer deletes keeps the text that another peer types into it at the same time, also after an Enter of a third peer moved the run.
- An undo of one of two joins of neighboring paragraphs puts the text back in its paragraphs, also after a third peer typed into the joined paragraph.
- A format and an Enter of two peers inside the same text keep the text in order, and an undo of either one keeps it in order.

The room format is now 6. A replica of an earlier format refuses the room with `schema-version-mismatch` and keeps its own deck, so update every peer of a room together.
