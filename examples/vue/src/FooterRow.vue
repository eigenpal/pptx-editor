<script setup lang="ts">
// Insert > Footer: the Vue twin of `FooterRow` in examples/vite/src/menus.tsx. It turns the
// footer on or off for every slide, through each layout's footer placeholder
// (`p:ph type="ftr"`). The check shows whether any slide has a footer, so the row never writes
// over existing footers. A new footer says "Footer"; a host passes its own text.
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const on = useEditorState((s) => (s.footer ?? null) !== null);
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  const text = (s.footer ?? null) !== null ? null : t('menu.footer');
  const can = editor.value.can({ type: 'setFooter', text });
  return can.ok ? null : can.reason;
});
</script>

<template>
  <PptxEditorMenu.Row
    :label="t('menu.footer')"
    icon="call_to_action"
    :checked="on"
    :disabled="!!reason"
    :reason="reason ? translateKeyOrText(t, reason) : null"
    @select="editor?.exec({ type: 'setFooter', text: on ? null : t('menu.footer') })"
  />
</template>
