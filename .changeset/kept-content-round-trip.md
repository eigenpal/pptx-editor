---
'@pptx-editor.dev/core': minor
---

Copy and paste keeps more content. Charts, diagrams, and embedded objects now copy with the parts that they own, and a paste gives them new parts in the same deck or another deck. An object with a copy in each branch of `mc:AlternateContent`, such as an equation, copies with all of its copies. Web links stay on paste after the link check, and a link to a slide stays in its own deck. `ClipboardRelationship` has the new optional fields `target` and `relationships`.

A paragraph animation range (`p:pRg`) now follows its paragraphs through text edits, and an effect whose paragraphs are all deleted goes with them. A move, resize, rotation, or flip of an object in `mc:AlternateContent` now changes its copies in the other branches too, so they stay.
