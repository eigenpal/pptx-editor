---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Improve rendering fidelity. Glow text stays selectable in printed PDFs. Pattern fills use measured tiles and stay square to the slide. Compound lines use measured bands. Arabic lines sit at the correct baseline. Vertical and stacked text is placed correctly. Text warps keep one size across their pairs. A path that starts without a move draws nothing. The PDF export draws pattern fills and compound lines, and draws enlarged pictures smoothly.
