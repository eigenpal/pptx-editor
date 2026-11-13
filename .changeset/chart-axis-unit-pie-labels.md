---
'@pptx-editor.dev/pro': patch
---

Fix chart value axes and best-fit pie labels. An automatic major unit now keeps one label line between ticks, so a small chart with large text shows every step. The top value label lowers the plot less at large text sizes. A best-fit pie label sits at equal distance from both edges of its slice, with its farthest corner 4 pt inside the circle.
