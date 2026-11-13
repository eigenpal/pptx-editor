---
'@pptx-editor.dev/pptx-to-pdf': patch
---

Draw group effects in exported PDFs. A group's shadow, glow, reflection, and soft edge draw once for the whole group, and the text of a soft-edged group's shapes paints over the faded shapes; a table's text fades with them. A group's blur applies to each shape in the group, and a group's inner shadow draws nothing. The `effect.group-silhouette`, `effect.group-soft-edge`, `effect.group-soft-edge-effects`, and `effect.group-body` diagnostics report these approximations.
