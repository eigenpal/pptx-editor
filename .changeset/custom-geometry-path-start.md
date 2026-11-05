---
'@pptx-editor.dev/core': patch
---

A custom geometry path draws nothing before its first `a:moveTo`, so a path that starts with `a:lnTo` or `a:arcTo` no longer draws from the top-left corner.
