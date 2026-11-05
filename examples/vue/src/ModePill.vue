<script setup lang="ts">
// The Editing | Viewing pill: the Vue twin of `ModePill` in examples/vite/src/TopBarMenus.tsx.
// The trigger names its menu and the current mode; a row's name is its label, its hint is its
// description.
import { computed, shallowRef, useId } from 'vue';
import { Icon, MenuItem, Popover, keepFocus, useTranslation } from '@pptx-editor.dev/vue';

const mode = defineModel<'edit' | 'view'>({ required: true });
const { t } = useTranslation();
const anchor = shallowRef<HTMLButtonElement | null>(null);
const open = shallowRef(false);
const id = useId();
const modes = ['edit', 'view'] as const;
const current = computed(() => (mode.value === 'edit' ? t('mode.editing') : t('mode.viewing')));

function choose(m: 'edit' | 'view'): void {
  mode.value = m;
  open.value = false;
}
</script>

<template>
  <button
    ref="anchor"
    type="button"
    class="demo-modebar__pill"
    aria-haspopup="menu"
    :aria-expanded="open"
    :aria-controls="open ? id : undefined"
    :aria-label="t('mode.current', { mode: current })"
    :data-pptx-tooltip="t('mode.label')"
    @mousedown="keepFocus"
    @click="open = !open"
  >
    <!-- One line, so no white space text sits between the icons and the label, as in React. -->
    <Icon :name="mode === 'edit' ? 'edit_note' : 'visibility'" className="" /><span
      class="demo-modebar__label"
      >{{ current }}</span
    ><Icon name="arrow_drop_down" className="demo-btn__chev" />
  </button>
  <Popover v-if="open" :id="id" :anchor="anchor" :label="t('mode.label')" @close="open = false">
    <MenuItem
      v-for="m in modes"
      :key="m"
      :checked="mode === m"
      :description="m === 'edit' ? t('mode.editingHint') : t('mode.viewingHint')"
      @select="choose(m)"
    >
      <span class="demo-menu__text">
        <span>{{ m === 'edit' ? t('mode.editing') : t('mode.viewing') }}</span>
        <span class="demo-menu__hint" aria-hidden="true">{{
          m === 'edit' ? t('mode.editingHint') : t('mode.viewingHint')
        }}</span>
      </span>
    </MenuItem>
  </Popover>
</template>
