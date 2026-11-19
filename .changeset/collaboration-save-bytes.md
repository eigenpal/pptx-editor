---
'@pptx-editor.dev/pro': patch
---

Two replicas of one package in a Yjs collaboration room save the same bytes. The room keeps the attribute order of each element when that order is not sorted, so every peer saves the order of the file. A replica of an earlier build reads the sorted order, and an element that it writes loses the kept order.
