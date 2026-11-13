---
'@pptx-editor.dev/pptx-to-pdf': patch
---

Draw group effects in exported PDFs. A group's shadow, glow, and reflection draw once for the whole group. A group's soft edge, inner shadow, and blur apply to each shape in the group. The `effect.group-silhouette` and `effect.group-body` diagnostics report these approximations.
