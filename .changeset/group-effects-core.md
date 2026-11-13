---
'@pptx-editor.dev/core': minor
---

Draw group effects (`p:grpSpPr/a:effectLst`) on the whole group: one outer shadow, and with the Pro drawing renderers, glow, soft edge, reflection, and blur. Text under a group soft edge paints once, sharp. Inner shadows, effects nested inside two effect groups, and effects of groups over 1,000 elements do not draw. `ResolvedGroup` has a new `effects` field.
