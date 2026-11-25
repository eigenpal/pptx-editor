---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Slide.comments` to the editing API, an SDK extension: read, add, reply to, edit, resolve, and delete slide comments in both formats through `Comment`, `CommentReply`, and their collections. The new `author` option of `createServer` and `createBrowser` signs new comments. The automation host gains the comment operations and `BrowserAutomationOptions`.
