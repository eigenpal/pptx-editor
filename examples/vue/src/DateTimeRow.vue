<script setup lang="ts">
// Insert > Date and time: the Vue twin of `DateTimeRow` in examples/vite/src/menus.tsx. It turns
// the date on or off for every slide, through each layout's date placeholder (`p:ph type="dt"`).
// The date is a field that shows the date when the file opens.
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const on = useEditorState((s) => s.dateTime === true);
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  const can = editor.value.can({ type: 'setDateTime', show: s.dateTime !== true });
  return can.ok ? null : can.reason;
});
</script>

<template>
  <PptxEditorMenu.Row
    :label="t('menu.dateTime')"
    icon="calendar_today"
    :checked="on"
    :disabled="!!reason"
    :reason="reason ? translateKeyOrText(t, reason) : null"
    @select="editor?.exec({ type: 'setDateTime', show: !on })"
  />
</template>
