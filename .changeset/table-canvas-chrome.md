---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

Tables now show editing chrome on the canvas. A handle at the top-left corner selects and moves the table, grid lines drag, and "+" buttons insert rows and columns. Cell ranges copy and paste as tab-separated text, Shift+Arrow extends into a cell range, and the cell menu groups Insert, Delete, and Select in submenus. The Format panel sets row height, column width, and table style options, and the Borders menu draws with a pen color, style, and weight. `insertColumn` now keeps the table width, and the columns narrow in proportion.
