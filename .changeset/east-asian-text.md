---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

East Asian text measures and draws with each run's `a:ea` font (complex-script text with `a:cs`, symbol characters with `a:sym`), and a theme with an empty `a:ea` uses its script font for the run's language. `eaLnBrk="0"` now breaks East Asian text between any two characters with no kinsoku rule, `latinLnBrk="1"` breaks inside Latin words, Korean text breaks at spaces, `hangingPunct` lets a comma or full stop hang past the line end, the presentation's `p:kinsoku` lists replace the built-in lists, and justified East Asian lines spread between characters, as the reference PDFs show. A missing East Asian font falls back to a host CJK face for the text's language, such as Noto Sans CJK JP, and a character that the run's face lacks uses another loaded face. The canvas shows the text that an input method composes at the caret, and the key that ends a composition runs no shortcut. `@pptx-editor.dev/i18n` adds Japanese (`/ja`), Korean (`/ko`), and Simplified Chinese (`/zh-CN`) catalogs, and `@pptx-editor.dev/pptx-to-pdf` draws East Asian text with Noto Sans CJK JP when the optional `@docx-editor.dev/fonts-cjk` is installed.
