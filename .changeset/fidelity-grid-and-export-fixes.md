---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Improve rendering fidelity. Underlines and superscripts sit on the whole-point baseline grid, equation arrays center their baselines on the math axis, and superscripts in equations rise 0.365 em. The PDF export slants synthetic italic text and draws tiled pictures without seams between the tiles.
