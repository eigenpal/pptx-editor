<script setup lang="ts">
// The light/dark toggle in the top bar: the Vue twin of examples/vite/src/ThemeToggle.tsx.
// Both use the `demo-theme-toggle` styles in examples/shared/shell.css.
import { radioGroupStep } from '@pptx-editor.dev/core/editor';
import { useTranslation } from '@pptx-editor.dev/vue';

const value = defineModel<'light' | 'dark'>({ required: true });
const { t } = useTranslation();

const SUN =
  'M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4';
const MOON = 'M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z';
const modes = ['light', 'dark'] as const;

// One Tab stop: the checked option; the arrow keys move the choice, as a radio group does.
function onKeydown(e: KeyboardEvent): void {
  const next = radioGroupStep(modes, value.value, e.key);
  if (!next) return;
  e.preventDefault();
  value.value = next;
  (e.currentTarget as HTMLElement).querySelector<HTMLElement>(`[data-value="${next}"]`)?.focus();
}
</script>

<template>
  <div
    class="demo-theme-toggle"
    role="radiogroup"
    :aria-label="t('pageSetup.appearance')"
    @keydown="onKeydown"
  >
    <button
      v-for="mode in modes"
      :key="mode"
      type="button"
      role="radio"
      class="demo-theme-toggle__option"
      :data-value="mode"
      :aria-checked="value === mode"
      :tabindex="value === mode ? 0 : -1"
      :aria-label="t(mode === 'light' ? 'pageSetup.light' : 'pageSetup.dark')"
      :title="t(mode === 'light' ? 'pageSetup.light' : 'pageSetup.dark')"
      @mousedown.prevent
      @click="value = mode"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <circle v-if="mode === 'light'" cx="12" cy="12" r="4" />
        <path :d="mode === 'light' ? SUN : MOON" />
      </svg>
    </button>
  </div>
</template>
