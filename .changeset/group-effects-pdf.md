---
'@pptx-editor.dev/pptx-to-pdf': patch
---

Draw group effects in exported PDFs: a group's shadow, glow, reflection, and soft edge draw once for the whole group, with sharp text over a soft edge. A group's blur applies to each shape. The `effect.group-silhouette`, `effect.group-soft-edge`, `effect.group-soft-edge-effects`, and `effect.group-body` diagnostics report the approximations.
