<script setup lang="ts">
// The demo's "Discard unsaved changes?" question: the Vue twin of
// examples/vite/src/DiscardDialog.tsx. Asked before New, Open, or a drop replaces a deck with
// edits that were not exported. A native modal with the editor's dialog classes, so it looks
// like every editor dialog. Escape and Keep editing cancel; the safe choice has the focus first.
import { onBeforeUnmount, onMounted, shallowRef } from 'vue';
import { useTranslation } from '@pptx-editor.dev/vue';

const emit = defineEmits<{ discard: []; cancel: [] }>();
const { t } = useTranslation();
const dialog = shallowRef<HTMLDialogElement | null>(null);
const keep = shallowRef<HTMLButtonElement | null>(null);
let opener: HTMLElement | null = null;

onMounted(() => {
  const el = dialog.value;
  if (!el) return;
  opener = el.ownerDocument.activeElement as HTMLElement | null;
  if (typeof el.showModal === 'function' && !el.open) el.showModal();
  else el.setAttribute('open', '');
  keep.value?.focus();
});
onBeforeUnmount(() => {
  const el = dialog.value;
  if (el?.open && typeof el.close === 'function') el.close();
  if (opener?.isConnected) opener.focus({ preventScroll: true });
});

function onCancel(e: Event): void {
  e.preventDefault();
  emit('cancel');
}
</script>

<template>
  <dialog
    ref="dialog"
    class="pptx-dialog"
    data-pptx-dialog="discardChanges"
    aria-labelledby="demo-discard-title"
    aria-describedby="demo-discard-message"
    @cancel="onCancel"
  >
    <div class="pptx-dialog__header">
      <h2 id="demo-discard-title" class="pptx-dialog__title">{{ t('app.discardTitle') }}</h2>
    </div>
    <div class="pptx-dialog__body">
      <p id="demo-discard-message">{{ t('app.discardMessage') }}</p>
    </div>
    <div class="pptx-dialog__footer">
      <button ref="keep" type="button" class="pptx-button" @click="emit('cancel')">
        {{ t('app.keepEditing') }}
      </button>
      <button type="button" class="pptx-button pptx-button--destructive" @click="emit('discard')">
        {{ t('app.discard') }}
      </button>
    </div>
  </dialog>
</template>
