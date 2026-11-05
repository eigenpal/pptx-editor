<script setup lang="ts">
// Insert > Slide numbers: the Vue twin of `SlideNumbersRow` in examples/vite/src/menus.tsx. It
// turns slide numbers on or off for every slide, through each layout's slide-number placeholder
// (`p:ph type="sldNum"`). The check shows whether the current slide has its number.
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const on = useEditorState((s) => s.slideNumbers);
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  const can = editor.value.can({ type: 'setSlideNumbers', show: !s.slideNumbers });
  return can.ok ? null : can.reason;
});
</script>

<template>
  <PptxEditorMenu.Row
    :label="t('menu.slideNumbers')"
    icon="format_list_numbered"
    :checked="on"
    :disabled="!!reason"
    :reason="reason ? translateKeyOrText(t, reason) : null"
    @select="editor?.exec({ type: 'setSlideNumbers', show: !on })"
  />
</template>
