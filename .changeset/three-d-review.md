---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Fix 3-D shapes. A zero-height shape with a bevel no longer writes invalid colors into the SVG or the PDF. The outer shadow and the glow of a turned face follow the face in the editor and in PDF export, and the glow lies under the extrusion. A face turned edge-on draws only its extrusion. Long outlines paint far fewer paths. Setting a bevel keeps a camera that the shape inherits from its layout or theme.
