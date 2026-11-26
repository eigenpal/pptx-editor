---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Slide.comments` to the editing API, an SDK extension. It reads, adds, replies to, edits, resolves, and deletes slide comments in both formats, through `Comment`, `CommentReply`, and their collections. The new `author` option of `createServer` and `createBrowser` signs new comments. The automation host adds the comment operations, `BrowserAutomationOptions`, and `commentAuthorName`.
