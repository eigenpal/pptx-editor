---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/i18n': minor
---

Collaboration sessions can own undo and redo, report a status snapshot with typed failure codes, and exchange presence and selections. The canvas shows collaboration peers' selected shapes and text carets. Custom `EditorCollaborationSession` implementations must now provide `statusSnapshot()`, and `subscribeStatus` listeners receive the reason code and detail.
