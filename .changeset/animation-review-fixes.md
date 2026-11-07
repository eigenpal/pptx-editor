---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/i18n': patch
---

Make animation and transition editing safer and keep more of the file. `setTransition` refuses a kind that is not a transition kind, and names from the file that match object keys read as unknown values. A reorder or a new start keeps the click step and group wrappers with their markup and ids, and a removal leaves the other effects where they are. In a collaboration session, removals of different effects all stay, a removal wins over a move of the same effect, and two effects added to one shape at the same time take two build groups. A save of a changed timing keeps one build entry for each shape and group, and a save repairs only the transitions that the session wrote. A duration between the speeds of a plain transition keeps its precision. The **Animations** tab labels read-only effects by their trigger, shows an error for a time that the editor refuses, names the advance controls differently, and clears the chosen effect on another slide. `animationStartLabel` gives the start label of an effect.
