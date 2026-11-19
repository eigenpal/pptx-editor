---
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

`useEditorCommand` and `useEditorValueCommand` return `disabledReason` as readable text in the root locale, in React and Vue. The hooks translate catalog keys, keep the engine's own text, and update the text when the `locale` prop changes. You no longer need `translateKeyOrText` to show the reason.
