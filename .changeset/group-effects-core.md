---
'@pptx-editor.dev/core': minor
---

Draw group effects (`p:grpSpPr/a:effectLst`) on the whole group: one outer shadow for all its shapes, and with the Pro drawing renderers, glow, soft edge, reflection, and blur. Text under a group soft edge stays sharp. A group's inner shadow draws nothing. `ResolvedGroup` has a new `effects` field.
