---
'@pptx-editor.dev/core': minor
---

Text lines now follow the measured line spacing, paragraph spacing, and stored-autofit rules (`a:lnSpc`, `a:spcBef`, `a:spcAft`, `a:normAutofit`), and shrink on overflow reduces line spacing before font size. `baselineOffsetPt` takes no spacing argument now; the four-argument form is deprecated.
