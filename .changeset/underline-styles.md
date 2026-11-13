---
'@pptx-editor.dev/core': patch
---

Draw every underline style, including dotted, dashed, and wavy underlines and their heavy and double forms. An underline outline (`a:uLn`, `a:uLnTx`) strokes the underline, and the new `RunProps.underlineLine` field holds it. Underlined runs on one line share one underline, and small caps in a superscript or subscript keep the script size.
