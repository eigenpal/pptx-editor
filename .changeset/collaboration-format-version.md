---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Add `readCollaborationFormatVersion()` and `COLLABORATION_FORMAT_VERSION` to `@pptx-editor.dev/pro`, so a participant who joins can wait until the Y.Doc holds the room and never seeds a second copy. `CollaborationFailureCode` gains `not-initialized`, for a Y.Doc that holds no room yet; an exhaustive switch over the codes needs a case for it.
