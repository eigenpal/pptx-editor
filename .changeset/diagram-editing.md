---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/editor-api': minor
---

Edit diagrams. Click a shape of a selected diagram, or press Enter on it, to type in the shape: the edit writes the drawing shape and its node in the data model, and keeps run formatting. Tab moves to the next shape. The new `addDiagramShape` and `deleteDiagramShape` commands add a shape after, before, or below a node, and delete a node with its children. They write the data model and lay the saved drawing out again, for diagrams whose drawing is one line of shapes with up to one level of children. `EditorSnapshot.diagram` reports the selected diagram. The **Diagram** toolbar group, the **Diagram Design** ribbon tab, and the diagram context menu show the new commands.
