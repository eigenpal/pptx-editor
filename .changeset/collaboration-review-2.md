---
'@pptx-editor.dev/pro': patch
---

The Yjs session bounds the read of removed runs and paragraphs per slide, and refuses a room format that arrives late.

- One read of a slide reads the fields of each record once and takes one step limit, so a hostile room cannot stall a peer with many parents that name one large record.
- A session refuses the room with `schema-version-mismatch` also when the room format arrives after attach, for example after a late provider sync.
- An undo checks every history step that it can apply for runs that a peer moved on, and a very long list never leaves such a run listed nowhere.
- Every peer reads the same restore marks when a room holds more of them than the read limit.
