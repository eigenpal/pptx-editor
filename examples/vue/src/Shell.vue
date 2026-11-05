<script setup lang="ts">
// The demo shell: the Vue twin of `Shell` in examples/vite/src/App.tsx. The family top bar
// (brand, file name, menus, Present, mode pill, theme, Open/New/Export, panel toggle) around
// the adapter's toolbar or ribbon, slide rail or sorter, canvas, notes, panel, and zoom pill.
// It reads the editor through composables, so it renders inside `PptxEditorRoot`.
import { onBeforeUnmount, shallowRef, watch } from 'vue';
import {
  Icon,
  PptxEditorContent,
  PptxEditorEditNotice,
  PptxEditorFindBar,
  PptxEditorFontNotice,
  PptxEditorLoading,
  PptxEditorNotesPane,
  PptxEditorOpenFailure,
  PptxEditorRibbon,
  PptxEditorSidePanel,
  PptxEditorSlideRail,
  PptxEditorSlideSorter,
  PptxEditorSlideshow,
  PptxEditorStatusBar,
  PptxEditorToolbar,
  PptxEditorViewport,
  PptxEditorZoomPill,
  keepFocus,
  useEditorState,
  usePptxEditor,
  useSlideshow,
  useTranslation,
} from '@pptx-editor.dev/vue';
import {
  PPTX_MIME,
  dismissPanelOverlay,
  dismissRailOverlay,
  exportDeck as saveDeck,
  exportPicture as savePicture,
  followPanelWidth,
  isPptxFile,
  carriesFiles,
  leftDropZone,
  panelStartsOpen,
  refuseDroppedFile,
  ribbonChrome,
  type PictureExport,
} from '../../shared/demo-files.ts';
import AdapterSwitcher from './AdapterSwitcher.vue';
import BrandLogo from './BrandLogo.vue';
import ExportMenu from './ExportMenu.vue';
import MenuBar from './MenuBar.vue';
import ModePill from './ModePill.vue';
import ThemeToggle from './ThemeToggle.vue';

// Shown in development and in the preview build, as in the React demo.
const showAdapterSwitcher =
  import.meta.env.DEV || import.meta.env.VITE_ENABLE_FRAMEWORK_SWITCHER === 'true';

const title = defineModel<string>('title', { required: true });
const mode = defineModel<'edit' | 'view'>('mode', { required: true });
const theme = defineModel<'light' | 'dark'>('theme', { required: true });
const emit = defineEmits<{ openFile: [file: File]; new: []; example: []; saved: [] }>();

const { t } = useTranslation();
const editor = usePptxEditor();
const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
// `?chrome=ribbon` shows the ribbon and the status bar in place of the flat toolbar.
const ribbon = ribbonChrome(location.search);
const canPresent = useEditorState((s) => s.status === 'ready' && s.slides.length > 0);
const show = useSlideshow();
// Closed on a narrow window, so the slide keeps a usable zoom (demo-files.ts).
const panel = shallowRef(panelStartsOpen());
const stopPanelWidth = followPanelWidth((open) => (panel.value = open));
onBeforeUnmount(stopPanelWidth);
const dragging = shallowRef(false);
// The slide thumbnails over the slide, on a narrow window only.
const rail = shallowRef(false);
const railToggle = shallowRef<HTMLButtonElement | null>(null);
// The overlay rail closes on Escape, on a press outside it, and after a slide is picked.
watch(rail, (open, _was, onCleanup) => {
  if (!open) return;
  onCleanup(
    dismissRailOverlay({
      root: () => railToggle.value?.closest<HTMLElement>('.demo-root') ?? null,
      editor: () => editor.value,
      close: () => (rail.value = false),
    })
  );
});
// The overlay side panel (a narrow window) closes on Escape, as the overlay rail does.
watch(
  panel,
  (open, _was, onCleanup) => {
    if (!open) return;
    onCleanup(
      dismissPanelOverlay({
        root: () => railToggle.value?.closest<HTMLElement>('.demo-root') ?? null,
        close: () => (panel.value = false),
      })
    );
  },
  { immediate: true }
);
// The deck name field, which File > Rename presentation shows and focuses.
const renaming = shallowRef(false);
const titleInput = shallowRef<HTMLInputElement | null>(null);

function rename(): void {
  renaming.value = true;
  requestAnimationFrame(() => (titleInput.value?.focus(), titleInput.value?.select()));
}
const fileInput = shallowRef<HTMLInputElement | null>(null);

function exportDeck(): void {
  saveDeck(editor.value, title.value || t('app.untitled'));
  emit('saved');
}

// Slide pictures come from the editor's own paint; nothing leaves the browser.
function exportPicture(kind: PictureExport): void {
  void savePicture(editor.value, kind, title.value || t('app.untitled'));
}

function onFile(e: Event): void {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  if (file) emit('openFile', file);
  input.value = '';
}

function onDragOver(e: DragEvent): void {
  if (!carriesFiles(e.dataTransfer)) return;
  e.preventDefault();
  dragging.value = true;
}

// A move onto a child of the workspace is not a leave: the hint does not flicker.
function onDragLeave(e: DragEvent): void {
  if (leftDropZone(e.currentTarget, e.relatedTarget)) dragging.value = false;
}

function onDrop(e: DragEvent): void {
  if (!carriesFiles(e.dataTransfer)) return;
  e.preventDefault();
  dragging.value = false;
  const file = e.dataTransfer?.files[0];
  if (isPptxFile(file)) emit('openFile', file);
  else if (file) refuseDroppedFile(editor.value, file);
}
</script>

<template>
  <PptxEditorViewport class="demo-root">
    <header class="demo-topbar" @dragover="onDragOver" @dragleave="onDragLeave" @drop="onDrop">
      <BrandLogo :name="t('app.productName')" />
      <AdapterSwitcher v-if="showAdapterSwitcher" current="vue" />
      <span class="demo-topbar__divider" />
      <!-- File > Rename presentation opens the name over the menu bar on a narrow window. -->
      <div :class="renaming ? 'demo-file demo-file--open' : 'demo-file'">
        <input
          ref="titleInput"
          v-model="title"
          class="demo-file__name"
          :aria-label="t('app.untitled')"
          @blur="renaming = false"
          @keydown.enter="($event.target as HTMLInputElement).blur()"
          @keydown.escape="($event.target as HTMLInputElement).blur()"
        />
        <span class="demo-file__ext">{{ t('app.extension') }}</span>
      </div>
      <MenuBar
        @rename="rename"
        @open="fileInput?.click()"
        @new="emit('new')"
        @example="emit('example')"
        @export="exportDeck"
        @export-picture="exportPicture"
      />
      <div class="demo-topbar__actions">
        <!-- On a narrow window the slide thumbnails open over the slide (shell.css). -->
        <button
          type="button"
          ref="railToggle"
          class="demo-icon-btn demo-rail-toggle"
          :aria-label="t('app.slidesToggle')"
          :data-pptx-tooltip="t('app.slidesToggle')"
          :aria-pressed="rail"
          @mousedown="keepFocus"
          @click="rail = !rail"
        >
          <Icon name="view_agenda" className="" />
        </button>
        <button
          type="button"
          class="demo-btn demo-btn--sm"
          :aria-label="t('slideshow.present')"
          :data-pptx-tooltip="t('slideshow.presentHint')"
          :disabled="!canPresent"
          @mousedown="keepFocus"
          @click="show.start('start')"
        >
          <Icon name="slideshow" className="" /><span class="demo-btn__label">{{
            t('slideshow.present')
          }}</span>
        </button>
        <ModePill v-model="mode" />
        <ThemeToggle v-model="theme" />
        <span class="demo-topbar__divider" />
        <button
          type="button"
          class="demo-btn demo-btn--quiet demo-btn--sm"
          :aria-label="t('menu.open').replace('…', '')"
          :data-pptx-tooltip="t('menu.open').replace('…', '')"
          @mousedown="keepFocus"
          @click="fileInput?.click()"
        >
          <Icon name="folder_open" className="" /><span class="demo-btn__label">{{
            t('menu.open').replace('…', '')
          }}</span>
        </button>
        <button
          type="button"
          class="demo-btn demo-btn--quiet demo-btn--sm"
          :aria-label="t('menu.new')"
          :data-pptx-tooltip="t('menu.new')"
          @mousedown="keepFocus"
          @click="emit('new')"
        >
          <Icon name="note_add" className="" /><span class="demo-btn__label">{{
            t('menu.new').split(' ')[0]
          }}</span>
        </button>
        <ExportMenu @export="exportDeck" @export-picture="exportPicture" />
        <button
          type="button"
          class="demo-icon-btn demo-panel-toggle"
          :aria-label="t('panel.toggle')"
          :data-pptx-tooltip="t('panel.toggle')"
          :aria-pressed="panel"
          @mousedown="keepFocus"
          @click="panel = !panel"
        >
          <Icon name="view_sidebar" className="" />
        </button>
      </div>
      <input ref="fileInput" type="file" :accept="`.pptx,${PPTX_MIME}`" hidden @change="onFile" />
    </header>
    <div class="demo-dock">
      <PptxEditorRibbon v-if="ribbon" />
      <PptxEditorToolbar v-else />
    </div>
    <div
      :class="rail ? 'demo-workspace demo-workspace--rail' : 'demo-workspace'"
      @dragover="onDragOver"
      @dragleave="onDragLeave"
      @drop="onDrop"
    >
      <PptxEditorSlideRail v-if="!sorter" />
      <!-- The slide sorter takes the canvas's place, so the status bar and the notices stay. -->
      <div class="demo-docpane">
        <PptxEditorSlideSorter v-if="sorter" />
        <div class="demo-canvas" :hidden="sorter">
          <PptxEditorContent />
          <PptxEditorFindBar />
          <PptxEditorLoading overlay />
          <PptxEditorFontNotice />
          <PptxEditorZoomPill v-if="!ribbon" />
        </div>
        <!-- The notes pane stays mounted in the slide sorter, so its open state and draft stay. -->
        <PptxEditorNotesPane :style="sorter ? { display: 'none' } : undefined" />
        <PptxEditorEditNotice />
        <PptxEditorOpenFailure />
        <PptxEditorStatusBar v-if="ribbon" />
      </div>
      <PptxEditorSidePanel v-if="panel" :on-close="() => (panel = false)" />
      <div v-if="dragging" class="demo-drop">{{ t('app.dropHint') }}</div>
    </div>
    <PptxEditorSlideshow />
    <!-- Dialogs of the demo, inside the editor scope so the editor's tokens apply. -->
    <slot />
  </PptxEditorViewport>
</template>
