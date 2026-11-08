---
'@pptx-editor.dev/core': minor
---

Move the caret and draw the selection on screen in right-to-left and mixed-direction text. Left and Right step one place on screen, and Alt or Ctrl with them steps one word on screen. Cmd with Left or Right goes to the left or the right edge of the line, and Home and End go to the start and the end of the line in reading order. In a table that runs right to left, Left and Right at a cell's edge move to the cell on that side. A selection draws one highlight for each part of the line that it covers. `TextPoint.affinity` tells which side of a direction boundary the caret shows on; clicks, the arrow keys, and typing set it, and edits and collaboration read only the offset. `Line.bidi` gives the trailing edge and the direction of each unit of a line in visual order.
