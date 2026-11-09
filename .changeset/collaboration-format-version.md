---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Add `readCollaborationFormatVersion()` and `COLLABORATION_FORMAT_VERSION` to `@pptx-editor.dev/pro`, so a participant who joins can wait until the Y.Doc holds the room and never seeds a second copy. The new `not-initialized` failure code means that a Y.Doc holds no room yet.
