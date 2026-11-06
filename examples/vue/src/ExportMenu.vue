<script setup lang="ts">
// The Export button and its menu (PPTX, PDF, PNG, ZIP, SVG): the Vue twin of `ExportMenu` in
// examples/vite/src/TopBarMenus.tsx. The trigger names its menu; a row's name is its label, its
// hint is its description.
import { computed, onBeforeUnmount, shallowRef, useId } from 'vue';
import { Icon, MenuItem, Popover, keepFocus, useTranslation } from '@pptx-editor.dev/vue';
import { PICTURE_WIDTH, watchPdfExport, type SlideExport } from '../../shared/demo-files.ts';

const emit = defineEmits<{ export: []; exportSlides: [kind: SlideExport] }>();
const { t } = useTranslation();
const anchor = shallowRef<HTMLButtonElement | null>(null);
const open = shallowRef(false);
const id = useId();
// While the server makes the PDF, the button says so and the PDF row waits.
const pdfBusy = shallowRef(false);
onBeforeUnmount(watchPdfExport((busy) => (pdfBusy.value = busy)));
const buttonLabel = computed(() => (pdfBusy.value ? t('menu.exportPdfBusy') : t('menu.export')));
const slideExports = [
  ['pdf', 'PDF', 'menu.exportPdf', 'menu.exportPdfHint'],
  ['png', 'PNG', 'menu.exportSlidePng', 'menu.exportSlidePngHint'],
  ['zip', 'ZIP', 'menu.exportSlidesPng', 'menu.exportSlidesPngHint'],
  ['svg', 'SVG', 'menu.exportSlideSvg', 'menu.exportSlideSvgHint'],
] as const;

function exportDeck(): void {
  emit('export');
  open.value = false;
}

function exportSlides(kind: SlideExport): void {
  emit('exportSlides', kind);
  open.value = false;
}
</script>

<template>
  <div class="demo-menu">
    <button
      ref="anchor"
      type="button"
      class="demo-btn demo-btn--sm demo-btn--primary"
      aria-haspopup="menu"
      :aria-expanded="open"
      :aria-controls="open ? id : undefined"
      :aria-label="buttonLabel"
      :aria-busy="pdfBusy || undefined"
      :data-pptx-tooltip="buttonLabel"
      @mousedown="keepFocus"
      @click="open = !open"
    >
      <Icon name="file_download" className="" /><span class="demo-btn__label" aria-hidden="true">{{
        buttonLabel
      }}</span
      ><Icon name="expand_more" className="demo-btn__chev" />
    </button>
    <Popover
      v-if="open"
      :id="id"
      :anchor="anchor"
      :label="t('menu.export')"
      align="end"
      @close="open = false"
    >
      <MenuItem :description="t('menu.exportPptxHint')" @select="exportDeck">
        <span class="demo-filetype" aria-hidden="true">PPTX</span>
        <span class="demo-menu__text">
          <span>{{ t('menu.exportPptx') }}</span>
          <span class="demo-menu__hint" aria-hidden="true">{{ t('menu.exportPptxHint') }}</span>
        </span>
      </MenuItem>
      <MenuItem
        v-for="[kind, type, label, hint] in slideExports"
        :key="kind"
        :disabled="kind === 'pdf' && pdfBusy"
        :reason="t('menu.exportPdfBusy')"
        :description="t(hint, { width: PICTURE_WIDTH })"
        @select="exportSlides(kind)"
      >
        <span class="demo-filetype" aria-hidden="true">{{ type }}</span>
        <span class="demo-menu__text">
          <span>{{ t(label) }}</span>
          <span class="demo-menu__hint" aria-hidden="true">{{
            t(hint, { width: PICTURE_WIDTH })
          }}</span>
        </span>
      </MenuItem>
    </Popover>
  </div>
</template>
