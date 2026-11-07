<script setup lang="ts">
// The PPTX Editor Vue demo: the Vue twin of examples/vite/src/App.tsx. It opens the same decks
// (`?deck=basic` for the basic deck), honors `?theme=dark` and `?chrome=ribbon`, and puts the
// demo's theme switch into File > Page setup. Everything loads and saves in the browser;
// nothing is uploaded.
import { computed, h, onBeforeUnmount, onMounted, shallowRef, watch, watchEffect } from 'vue';
import { PageSetupDialog, PptxEditorRoot, type PptxEditorPopups } from '@pptx-editor.dev/vue';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { drawingModule } from '@pptx-editor.dev/pro/drawing';
import {
  BLANK,
  EXAMPLES,
  fetchDeck,
  refuseDeckFetch,
  fileTitle,
  initialExample,
  initialTheme,
  openDeck,
  rememberTheme,
  showDeckTitle,
  trackUnsavedEdits,
  type UnsavedEdits,
} from '../../shared/demo-files.ts';
import { CATALOG_FAMILIES, useDemoFonts } from './fonts.ts';
import DiscardDialog from './DiscardDialog.vue';
import Shell from './Shell.vue';

// Glow, inner shadow, reflection, and soft edge (EigenPal Pro License).
const MODULES: readonly EditorModule[] = [
  drawingModule({ licenseKey: import.meta.env.VITE_PPTX_LICENSE }),
];

const base = import.meta.env.BASE_URL;
const initial = initialExample(location.search);
const bytes = shallowRef<Uint8Array | undefined>();
const title = shallowRef<string>(EXAMPLES[initial].title);
const mode = shallowRef<'edit' | 'view'>('edit');
const theme = shallowRef<'light' | 'dark'>(initialTheme(location.search));
// A theme picked here is kept; until then the demo starts light.
watch(theme, rememberTheme);

const editor = shallowRef<PptxEditorInstance | null>(null);
// Unsaved edits: New, Open, and a drop ask before they replace them, and so does leaving.
let unsaved: UnsavedEdits | null = null;
onBeforeUnmount(() => unsaved?.dispose());
const discard = shallowRef<(() => void) | null>(null);
function replaceDeck(run: () => void): void {
  if (unsaved?.dirty()) discard.value = run;
  else run();
}
function confirmDiscard(): void {
  const run = discard.value;
  discard.value = null;
  run?.();
}

// The name changes only once the deck opened: a refused file keeps the open deck's name.
async function openBytes(next: Uint8Array, label: string): Promise<void> {
  if (await openDeck(editor.value, next)) title.value = label;
}

async function loadUrl(file: string, label: string): Promise<void> {
  try {
    await openBytes(await fetchDeck(base, file), label);
  } catch (cause) {
    refuseDeckFetch(editor.value, cause);
  }
}

function openFile(file: File): void {
  replaceDeck(
    () => void file.arrayBuffer().then((b) => openBytes(new Uint8Array(b), fileTitle(file)))
  );
}

// The first deck: `initial` comes from the URL at mount.
const deckFailed = shallowRef<unknown>(null);
onMounted(
  () =>
    void fetchDeck(base, EXAMPLES[initial].file).then(
      (b) => (bytes.value = b),
      (cause: unknown) => (deckFailed.value = cause ?? new Error('fetch failed'))
    )
);
watch([deckFailed, editor], ([failed, instance]) => {
  if (failed && instance) refuseDeckFetch(instance, failed);
});
watchEffect(() => {
  document.documentElement.style.colorScheme = theme.value;
});
// The deck name in the browser tab.
watchEffect(() => showDeckTitle(title.value));

const fonts = useDemoFonts();
// The registry mounts Page setup; this render callback adds the demo's theme switch to it.
const popups = computed<PptxEditorPopups>(() => ({
  pageSetup: (props) =>
    h(PageSetupDialog, {
      ...props,
      theme: theme.value,
      onThemeChange: (next: 'light' | 'dark') => (theme.value = next),
    }),
}));

function onReady(instance: PptxEditorInstance): void {
  // The e2e suite and the verification scripts drive the editor through this handle.
  (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = instance;
  editor.value = instance;
  unsaved?.dispose();
  unsaved = trackUnsavedEdits(instance);
}
</script>

<template>
  <PptxEditorRoot
    :document="bytes"
    :mode="mode"
    :color-mode="theme"
    :popups="popups"
    :modules="MODULES"
    :fonts="fonts"
    :font-families="CATALOG_FAMILIES"
    @ready="onReady"
  >
    <Shell
      v-model:title="title"
      v-model:mode="mode"
      v-model:theme="theme"
      @new="replaceDeck(() => loadUrl(BLANK.file, BLANK.title))"
      @example="replaceDeck(() => loadUrl(EXAMPLES.showcase.file, EXAMPLES.showcase.title))"
      @open-file="openFile"
      @saved="unsaved?.markSaved()"
      :deck-failed="deckFailed !== null"
    >
      <DiscardDialog v-if="discard" @cancel="discard = null" @discard="confirmDiscard" />
    </Shell>
  </PptxEditorRoot>
</template>
