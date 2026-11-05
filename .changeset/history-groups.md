---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Group undo steps with `Editor.beginHistoryGroup()` and `exec(command, { historyGroup })`, restore the selection after a step on redo, and commit `projection` or `awareness` transactions that add no undo step; `runToolbarCommand` and `useEditorValueCommand().execute` take `{ historyGroup }` last; `historyGroup` replaces the store's `mergeKey` and `setTableGrid.gesture`, `closeHistoryGroup()` replaces `sealHistory()`, and `redo()` returns the selection after the step.
