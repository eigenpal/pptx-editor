---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

A click in a text placeholder that has no `p:txBody` places the caret, and the placeholder shows its prompt. The first edit that adds text adds the text body, which inherits the body properties of the layout and master placeholders. A click, a read, or a delete key leaves the file unchanged. In a collaboration room, two peers that type into such a placeholder at once both keep their text.
