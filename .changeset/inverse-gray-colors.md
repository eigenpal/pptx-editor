---
'@pptx-editor.dev/core': patch
---

Draw colors with the `a:inv` and `a:gray` transforms in the correct shade: `a:inv` inverts linear RGB components, and `a:gray` weights the sRGB components by the Rec. 709 primaries.
