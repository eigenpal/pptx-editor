// Which part of a dialog a menu row opens on: Character Spacing > More Spacing…
// opens the Font dialog on its Character Spacing tab. The requests live in core
// (`toolbar-controls.ts`), so the React and Vue adapters share them.

export { requestDialogSection, takeDialogSection } from '@pptx-editor.dev/core/editor';
