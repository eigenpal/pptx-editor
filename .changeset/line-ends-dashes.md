---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': minor
---

Draw dashes and line ends as the reference PDFs show them, in the editor and in PDF export. Line ends on a line thinner than 2 pt take the size of a 2 pt line. Triangles, stealth arrows, open arrows, ovals, and diamonds take their measured sizes, and the line stops under each head at its measured point. A preset dash with square caps keeps its flat length, and only the ends of the line get the cap. The PDF export now draws dashes and line ends, and no longer reports `line.dash` or `line.ends`. The `arrowhead-joins` approximation is gone. A custom dash keeps at most 64 dash and space pairs, a negative length counts as 0, and a pattern with no length draws a solid line.
