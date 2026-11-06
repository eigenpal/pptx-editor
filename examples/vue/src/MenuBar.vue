<script setup lang="ts">
// The demo's menu bar: the Vue twin of examples/vite/src/menus.tsx, built on the library's
// `PptxEditorMenu`. The built-in menus (File, Edit, View, Insert, Format, Arrange, Slide Show)
// come from the chrome registry, and the demo adds its own rows: opening and exporting files,
// slide structure, slide numbers, and Help. Every built-in row runs a registered slot, so its
// state and action are the toolbar's.
import { PptxEditorMenu, useTranslation } from '@pptx-editor.dev/vue';
import type { SlideExport } from '../../shared/demo-files.ts';
import DuplicateSlideRow from './DuplicateSlideRow.vue';
import SlideNumbersRow from './SlideNumbersRow.vue';

const emit = defineEmits<{
  rename: [];
  open: [];
  new: [];
  example: [];
  export: [];
  exportSlides: [kind: SlideExport];
}>();
const { t } = useTranslation();
</script>

<template>
  <PptxEditorMenu class="demo-menubar">
    <!-- File: the demo's own file rows around the built-in Page setup. -->
    <PptxEditorMenu.File :preset="false">
      <PptxEditorMenu.Row :label="t('menu.open')" icon="folder_open" @select="emit('open')" />
      <PptxEditorMenu.Row :label="t('menu.new')" icon="note_add" @select="emit('new')" />
      <PptxEditorMenu.Row
        :label="t('menu.openSample')"
        icon="slideshow"
        @select="emit('example')"
      />
      <PptxEditorMenu.Row
        :label="t('menu.renameDeck')"
        icon="text_fields"
        @select="emit('rename')"
      />
      <PptxEditorMenu.Separator />
      <PptxEditorMenu.Item slot-id="file.pageSetup" />
      <PptxEditorMenu.Separator />
      <PptxEditorMenu.Group :label="t('menu.export')">
        <PptxEditorMenu.Row
          :label="t('menu.exportPptx')"
          icon="file_download"
          @select="emit('export')"
        />
        <PptxEditorMenu.Row
          :label="t('menu.exportPdf')"
          icon="description"
          @select="emit('exportSlides', 'pdf')"
        />
        <PptxEditorMenu.Row
          :label="t('menu.exportSlidePng')"
          icon="image"
          @select="emit('exportSlides', 'png')"
        />
        <PptxEditorMenu.Row
          :label="t('menu.exportSlidesPng')"
          icon="folder_zip"
          @select="emit('exportSlides', 'zip')"
        />
        <PptxEditorMenu.Row
          :label="t('menu.exportSlideSvg')"
          icon="photo_library"
          @select="emit('exportSlides', 'svg')"
        />
      </PptxEditorMenu.Group>
    </PptxEditorMenu.File>
    <!-- The slide rows go in Insert (host rows follow the registry rows here); delete and move
         live in the slide rail's menu. No extra Slide menu, so the bar fits a laptop window. -->
    <PptxEditorMenu.Insert>
      <PptxEditorMenu.Separator />
      <DuplicateSlideRow />
      <!-- In a group, the `slide.add` row is a second row next to the built-in New slide, not a
           replacement of it. The galleries draw each layout with the engine. -->
      <PptxEditorMenu.Group>
        <PptxEditorMenu.Item
          slot-id="slide.add"
          :direct="false"
          label-key="rail.chooseLayout"
          icon="dashboard"
        />
        <PptxEditorMenu.Item slot-id="slide.layout" />
        <PptxEditorMenu.Item slot-id="slide.reset" />
      </PptxEditorMenu.Group>
      <SlideNumbersRow />
    </PptxEditorMenu.Insert>
    <PptxEditorMenu.Entry id="help" :label="t('menu.help')">
      <PptxEditorMenu.Row
        :label="t('menu.about')"
        icon="info"
        href="https://github.com/eigenpal/pptx-editor"
      />
      <PptxEditorMenu.Row
        :label="t('menu.reportIssue')"
        icon="bug_report"
        href="https://github.com/eigenpal/pptx-editor/issues"
      />
    </PptxEditorMenu.Entry>
  </PptxEditorMenu>
</template>
