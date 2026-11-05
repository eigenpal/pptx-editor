---
'@pptx-editor.dev/editor-api': patch
---

Proxies from a refused sync stay usable, as in Office.js: only the refused sync's calls are dropped, and a retry with the same objects works while they exist.
