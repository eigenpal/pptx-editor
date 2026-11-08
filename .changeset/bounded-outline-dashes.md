---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Bound native dash expansion on shape and picture outlines and table borders. Outlines and borders that could require more than 100,000 dash segments draw solid in both SVG and PDF. Save preserves the original pattern.
