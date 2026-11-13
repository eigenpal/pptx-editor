---
'@pptx-editor.dev/pptx-to-pdf': patch
---

In exported PDFs, a group's shadow, glow, reflection, and soft edge draw once for the whole group, with sharp text over a soft edge. A group's blur applies to each shape. Diagnostics `effect.group-silhouette`, `effect.group-soft-edge`, `effect.group-soft-edge-effects`, `effect.group-body`, and `effect.group-text-order` report the approximations.
