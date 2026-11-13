---
'@pptx-editor.dev/core': patch
---

Improve the horizontal spacing of equations. Scripts use the advances of the script-style forms of a math font and the space after a script of STIX Two Math. `TextMeasurer.advances` takes a new `mathScript` option for those advances, and `ResolvedFace` has a new `mathTable` flag for faces with a `MATH` table. Radicals widen with their height, end with their radicand, and place a degree with the kerns of the font. Scripts of scripts draw at 60 % of the size, to the nearest half point.
