---
'@pptx-editor.dev/core': patch
---

Selecting a large table and setting its borders take less time: a border fragment is parsed once, not once for each cell edge.
