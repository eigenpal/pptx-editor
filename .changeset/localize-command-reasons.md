---
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Breaking: `useEditorCommand` and `useEditorValueCommand` return `disabledReason` as readable text in the root locale, not a catalog key. The text follows the `locale` prop, and the engine's own text stays unchanged. Remove `translateKeyOrText` around the reason, and do not compare it with a key.
