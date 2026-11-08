---
'@pptx-editor.dev/pro': patch
---

The Yjs session reads removed runs and paragraphs the same way on every peer.

- Every peer puts back the same removed paragraphs with a peer's text, also after more than 1,024 joins or deletes across paragraphs in one text body.
- One read of the removed runs and paragraphs of a text body or paragraph has a step limit, so a hostile room cannot stall a peer.
- An undo of an Enter that a peer's format followed leaves the formatted text where the peer moved it with another Enter, so each character shows once.
