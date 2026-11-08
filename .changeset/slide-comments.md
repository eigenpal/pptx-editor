---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/editor-api': minor
---

Add, reply to, edit, delete, resolve, and reopen slide comments with the new comment commands or in `PptxEditorReview`, the comments pane of `@pptx-editor.dev/pro/react` and `@pptx-editor.dev/pro/vue`. Give it to the side panel with the `panes` prop in React or the `#comments` slot in Vue. The side panel tabs show icons, and the selected tab also shows its label. `useReview`, `useReviewOf`, `useReviewItem`, and `useReviewAuthor` read the comments for your own UI. The `pane` option of `reviewModule` picks the marker style (`icon` or `initials`) and whether a marker click opens the pane. The `author` prop signs new comments, markers on the canvas select their threads, and `readCommentThreads` reads the threads of any slide. Comments keep their format: threaded or legacy. Duplicated slides keep their threaded comments, and two collaborators who add comments at once keep both.

The collaboration room format is now 7, because the room keeps comment lists as XML trees and a save joins the lists of two peers. A replica of an earlier format refuses the room with `schema-version-mismatch`, so update every peer of a room together.
