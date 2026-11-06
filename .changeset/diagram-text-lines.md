---
'@pptx-editor.dev/core': patch
---

Place diagram text lines where the reference PDFs put them: diagram text spaces its lines on the font's ascent plus descent, not on 1.2 times the font size. Kerning (`kern`) now uses only the `kern` table of the font that a deck names, so Carlito text draws with its plain advances, as in the reference PDFs. A substitute font keeps its own kerning pairs.
