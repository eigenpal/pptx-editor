<script setup lang="ts">
// Insert > Footer: the Vue twin of `FooterRow` in examples/vite/src/menus.tsx. It turns the
// footer on or off for every slide, through each layout's footer placeholder
// (`p:ph type="ftr"`). On keeps the footers that slides have. A deck without one gets the text
// "Footer text"; a host passes its own `text`. The row waits until all slides are open, so its
// check reads every slide.
import {
  PptxEditorMenu,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from '@pptx-editor.dev/vue';
import { translateKeyOrText, type EditorCommand } from '@pptx-editor.dev/core/editor';

const { t } = useTranslation();
const editor = usePptxEditor();
const on = useEditorState((s) => (s.footer ?? null) !== null);
const command = (on: boolean): EditorCommand =>
  on
    ? { type: 'setFooter', show: false }
    : { type: 'setFooter', show: true, text: t('menu.footerText') };
// A string, not the result object, so the subscription sees an unchanged answer as equal.
const reason = useEditorState((s) => {
  if (!editor.value) return 'disabled.noEditor';
  if ((s.pendingSlides ?? 0) > 0) return 'disabled.slidesOpening';
  const can = editor.value.can(command((s.footer ?? null) !== null));
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
    @select="editor?.exec(command(on))"
  />
</template>
