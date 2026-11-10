---
'@pptx-editor.dev/pro': patch
---

Keep peer text during collaborative undo and redo. Undoing formatting or paragraph splits preserves peer insertions and deletions. Undoing typing removes its concurrent copies. Redoing deletion also removes peer copies, including copies of restored text.
