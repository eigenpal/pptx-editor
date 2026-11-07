---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

A deck no longer stays on the loading spinner. The open waits at most 2 seconds for host fonts, HarfBuzz, and the packaged substitutes. Then it paints with fixed metrics and the faces it has. Late fonts install when they arrive, and the editor reflows. A deck that opens in a hidden tab no longer waits for the tab to show.

- `editor.fontMeasurement()` tells whether fonts are still resolving. The font substitution list stays empty until they finish.
- An error during the open shows the open-failure message with the code `failed`, in place of the spinner. Zip and XML limits have their own codes, for example `xml-depth-limit`. A DTD shows the new `blocked` reason: add it to an exhaustive map of `DocumentOpenFailureReason` values.
- The open-failure message has a **Copy details** button (`showCopy`). `openFailureReport()` gives the same text: code, reason, editor version, browser, and time. When the browser blocks the clipboard, the message shows the text to copy by hand.
- Large decks open and edit in linear time. This covers many slide IDs, `mc:Requires` prefixes, unmatched placeholders, nested namespace declarations, toolbar state on long text, font changes on many shapes, Replace All in a long run, and AutoCorrect after a long word.
- The font wait and the text engine wait run at the same time, under one 2 second limit.
- The shapes whose text you edit before the text engine starts fit again when it starts. Other shapes keep the fit that the file has. The new fit is one undo step, `fit text`.
- A load that fails during its font step keeps the fonts and the font waits of the deck that stays open.
- Typing repaints only the shapes that change. Unchanged shapes keep their resolved style, layout, and SVG nodes. A shrink-on-overflow fit stops measuring at the first paragraph that overflows. A fit value that the file already has writes no change.
- Repeated opens no longer grow memory: the text engine reads font tables without a copy of the whole font.
- With no deck open, the slide canvas is named "No presentation open". **Close** on the open-failure message moves focus to it.
- In development, the editor warns when the `document` or `fonts` prop changes more than 10 times in 1 second.
