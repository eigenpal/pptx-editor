<script setup lang="ts">
// Insert > Date and time: the Vue twin of `DateTimeRow` in examples/vite/src/menus.tsx. It turns
// the date on or off for every slide, through each layout's date placeholder (`p:ph type="dt"`).
// On keeps the dates that slides have; a deck without one gets today's date as a field. The row
// waits until all slides are open, so its check reads every slide.
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const on = useEditorState((s) => (s.dateTime ?? null) !== null);
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  if ((s.pendingSlides ?? 0) > 0) return 'disabled.slidesOpening';
  const can = editor.value.can({ type: 'setDateTime', show: (s.dateTime ?? null) === null });
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
