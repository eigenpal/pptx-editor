<script setup lang="ts">
// The Export button and its menu (PPTX, PNG, ZIP, SVG): the Vue twin of `ExportMenu` in
// examples/vite/src/TopBarMenus.tsx. The trigger names its menu; a row's name is its label, its
// hint is its description.
import { shallowRef, useId } from 'vue';
import { Icon, MenuItem, Popover, keepFocus, useTranslation } from '@pptx-editor.dev/vue';
import { PICTURE_WIDTH, type PictureExport } from '../../shared/demo-files.ts';

const emit = defineEmits<{ export: []; exportPicture: [kind: PictureExport] }>();
const { t } = useTranslation();
const anchor = shallowRef<HTMLButtonElement | null>(null);
const open = shallowRef(false);
const id = useId();
const pictures = [
  ['png', 'PNG', 'menu.exportSlidePng', 'menu.exportSlidePngHint'],
  ['zip', 'ZIP', 'menu.exportSlidesPng', 'menu.exportSlidesPngHint'],
  ['svg', 'SVG', 'menu.exportSlideSvg', 'menu.exportSlideSvgHint'],
] as const;

function exportDeck(): void {
  emit('export');
  open.value = false;
}

function exportPicture(kind: PictureExport): void {
  emit('exportPicture', kind);
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
      :aria-label="t('menu.export')"
      :data-pptx-tooltip="t('menu.export')"
      @mousedown="keepFocus"
      @click="open = !open"
    >
      <Icon name="file_download" className="" /><span class="demo-btn__label" aria-hidden="true">{{
        t('menu.export')
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
        v-for="[kind, type, label, hint] in pictures"
        :key="kind"
        :description="t(hint, { width: PICTURE_WIDTH })"
        @select="exportPicture(kind)"
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
