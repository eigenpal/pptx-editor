---
'@pptx-editor.dev/core': minor
---

The document store now refuses a transaction that breaks a table grid, a merge, a shape id, the slide list, a relationship target, or a content type, with a typed `tree-invariant` or `package-invariant` rejection that names the part and the rule.
