---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the SDK extension `TextRange.search(searchText, options?)`, in the shape of `Range.search` of the DOCX Editor API. It returns a `TextRangeCollection` of the matches inside the range, with `items`, `getCount`, `getItemAt`, `getFirst`, `getFirstOrNullObject`, and `getLast`. It uses the rules of `replaceAll`. The automation protocol adds the `findText` operation.
