---
'@pptx-editor.dev/core': patch
---

Shrink on overflow now ignores empty paragraphs at the end of the text. A new empty line at the end no longer makes the text smaller. At an 85 % font scale, shrink on overflow tries 10 % less line spacing before 20 % less.
