---
'@pptx-editor.dev/core': minor
---

Lay out right-to-left and mixed-direction text with the Unicode Bidirectional Algorithm (UAX #9). A paragraph with `rtl="1"` reads from the right, and its margins, indent, and bullet start at the right edge. Each line shows Arabic and Hebrew right to left, with Latin words and numbers left to right inside them. The SVG paint draws each right-to-left word as one shaped text element, so Arabic letters join. `GlyphRun.rtl` marks a run that reads right to left. `a:pPr/@algn` names physical edges, so `l` stays left and `r` stays right in a right-to-left paragraph.
