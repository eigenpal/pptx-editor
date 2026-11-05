<script setup lang="ts">
// Insert > Duplicate slide: the Vue twin of `DuplicateSlideRow` (a `CommandRow`) in
// examples/vite/src/menus.tsx. The row duplicates the current slide, and it is disabled with the
// engine's reason when `Editor.can` refuses the command.
import { computed } from 'vue';
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText, type EditorCommand } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const slide = useEditorState((s) => s.currentSlide);
const command = computed<EditorCommand>(() => ({ type: 'duplicateSlide', index: slide.value }));
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  const can = editor.value.can({ type: 'duplicateSlide', index: s.currentSlide });
  return can.ok ? null : can.reason;
});
</script>

<template>
  <PptxEditorMenu.Row
    :label="t('rail.duplicateSlide')"
    icon="content_copy"
    :disabled="!!reason"
    :reason="reason ? translateKeyOrText(t, reason) : null"
    @select="editor?.exec(command)"
  />
</template>
