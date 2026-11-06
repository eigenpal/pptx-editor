---
'@pptx-editor.dev/core': patch
---

Place diagram text lines where the reference PDFs put them: diagram text spaces its lines on the font's ascent plus descent, not on 1.2 times the font size. A font that a deck names now kerns (`kern`) only when it has a `kern` table, so Carlito text draws with its plain advances, as in the reference PDFs. A substitute font, also one from a host substitution, keeps its own kerning pairs.
