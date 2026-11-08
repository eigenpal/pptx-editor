---
'@pptx-editor.dev/core': minor
---

Move the caret and draw the selection on screen in right-to-left and mixed-direction text. Left and Right step one place on screen, and Alt or Ctrl with them steps one word on screen. Home and End go to the edges of the line in the paragraph's reading direction. A selection draws one highlight for each part of the line that it covers. `TextPoint.affinity` tells which side of a direction boundary the caret shows on; clicks, the arrow keys, and typing set it, and edits and collaboration read only the offset. `Line.bidi` gives the trailing edge and the direction of each unit of a line in visual order.
