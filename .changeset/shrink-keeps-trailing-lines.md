---
'@pptx-editor.dev/core': patch
---

Shrink on overflow no longer shrinks text for empty paragraphs at the end of a text box, so a new line or a small edit in a box with room keeps the text size. An 85 % font scale now tries the 10 % line spacing reduction before the 20 % one.
