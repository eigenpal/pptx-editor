---
'@pptx-editor.dev/core': patch
---

A deck that inflates past 8 MiB at more than 200 times its file size is now refused, as in DOCX Editor (the threshold was 64 MiB), and a save stays within that cap, so a deck that opened opens again.
