---
'@pptx-editor.dev/core': patch
---

Keep shapes that a cut cannot copy. Ctrl/Cmd+X deleted every selected shape even when the clipboard could not carry one of them, such as a chart whose parts cannot travel. A cut now deletes only a selection that its copy carries in full; otherwise it is refused with a reason, and the clipboard stays as it was.
