<p align="center">
  <a href="https://pptx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/pptx-editor/main/.github/assets/readme-logo-light.svg" alt="PPTX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/vue"><img src="https://img.shields.io/npm/v/@pptx-editor.dev/vue.svg?style=flat-square&color=C43E1C" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@pptx-editor.dev/vue"><img src="https://img.shields.io/npm/dm/@pptx-editor.dev/vue.svg?style=flat-square&color=C43E1C" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/pptx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=C43E1C" alt="license" /></a>
</p>

# @pptx-editor.dev/vue

The Vue 3 adapter for PPTX Editor. It is the twin of `@pptx-editor.dev/react`, in the shape of DOCX Editor's Vue adapter. `PptxEditorRoot` owns the engine instance, `PptxEditorViewport` scopes the styles, and `PptxEditorContent` attaches the slide surface. The toolbar, ribbon, menus, slide rail, notes, side panel, dialogs, find bar, and slide show read the engine through composables and hold no editing state.

The Vue parts have the same props, markup, and behavior as the React parts. A part on React's `PptxEditor.*` namespace is `PptxEditor<Part>` in Vue, for example `PptxEditorToolbar`. `scripts/parity/parity.contract.json` maps each React prop and export to its Vue twin, and `bun run check:parity` keeps the two adapters in step. The feature matrix on the docs site lists what the editor supports.

## Install

```bash
npm install @pptx-editor.dev/core @pptx-editor.dev/vue
```

## Use

Import the stylesheet once, and give the editor a container with a defined height:

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditor } from '@pptx-editor.dev/vue';
import '@pptx-editor.dev/core/styles/editor.css';

const bytes = shallowRef<Uint8Array>();
fetch('/deck.pptx')
  .then((response) => response.arrayBuffer())
  .then((buffer) => (bytes.value = new Uint8Array(buffer)));

function onReady(editor: PptxEditorInstance) {
  console.info('Editor ready', editor.snapshot().status);
}
</script>

<template>
  <div style="height: 100vh">
    <PptxEditor class="my-editor" :document="bytes" @ready="onReady" />
  </div>
</template>
```

## Compose your own layout

Compose the root, viewport, and content, and put your own controls inside the root. Composables must run in a descendant of `PptxEditorRoot`:

```vue
<script setup lang="ts">
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorToolbar,
  PptxEditorViewport,
  ToolbarGroup,
  ToolbarSlot,
} from '@pptx-editor.dev/vue';

defineProps<{ bytes?: Uint8Array }>();
</script>

<template>
  <PptxEditorRoot :document="bytes">
    <PptxEditorViewport>
      <PptxEditorToolbar>
        <ToolbarSlot slot-id="text.strike" hidden />
        <ToolbarGroup id="acme" label="Acme">
          <button type="button">Ask legal</button>
        </ToolbarGroup>
      </PptxEditorToolbar>
      <PptxEditorContent />
    </PptxEditorViewport>
  </PptxEditorRoot>
</template>
```

Without single-file components, use render functions:

```ts
import { createApp, h } from 'vue';
import { PptxEditor } from '@pptx-editor.dev/vue';
import '@pptx-editor.dev/core/styles/editor.css';

const bytes = new Uint8Array(await (await fetch('/deck.pptx')).arrayBuffer());
createApp({
  render: () => h(PptxEditor, { document: bytes, onReady: (editor) => console.info(editor) }),
}).mount('#app');
```

## Composables

| Composable | Returns |
| --- | --- |
| `usePptxEditor()` | `UsePptxEditorReturn`: the editor, or null before it mounts |
| `useEditorState(selector, isEqual?)` | `UseEditorStateReturn<T>`: the selected slice, updated only when it changes |
| `useEditorCommand(slotId)` | `UseEditorCommandReturn`: `execute` (it returns `true` when the engine ran the command, `false` on a refusal), and `isEnabled`, `isActive`, `disabledReason`, `value` as computed refs. `disabledReason` is a catalog key or the engine's own text: show it with `translateKeyOrText(t, reason)` from `@pptx-editor.dev/core/editor`. `slotId` may be a ref or a getter |
| `useTranslation()` | `UseTranslationReturn`: `t`, with shortcut text for the platform |
| `useFontFormat()` | `UseFontFormatReturn`: `format` (the `FontFormatRead` of the selected text or of the selected shapes' runs, with `disagrees`), `isEnabled`, and `apply(update)`, which sends one `setFontFormat` and returns whether the engine accepted it |
| `useBulletFormat()` | `UseBulletFormatReturn`: `format` (the `BulletFormatRead` of the selected paragraphs: `null` when they have no list, `undefined` when they differ or no text is selected), `isEnabled`, and `apply(format)`, which sends one `setBulletFormat` and returns whether the engine accepted it |
| `useParagraphFormat()` | `UseParagraphFormatReturn`: `format` (the selected paragraphs' `ParagraphFormatRead`, with `disagrees`), `isEnabled`, and `apply(update)`, which sends one `setParagraphFormat` and returns whether the engine accepted it |
| `useSlideLayouts()` | `UseSlideLayoutsReturn`: the deck's slide layouts in gallery order, with their master's name, for a layout picker of your own |

The table lists the composables most hosts need. For every composable and its `Use<Name>Return` type, such as `useEditorValueCommand`, `useEditorDialog`, `useSlideshow`, `useNotesPane`, `useOpenFailure`, and `useFonts`, see the [API report](../../docs/api/vue/index.api.md).

The state these composables report comes from the same core functions as the React hooks (`chromeCommandState`, `EMPTY_EDITOR_SNAPSHOT`), so both adapters show the same enabled state and the same refusal reasons. `PptxEditorRoot` holds one editor subscription and shares it with every composable below it.

## Dialogs

`PptxEditorRoot` mounts the editor's dialogs from one `popups` registry, as React and DOCX Editor do. A render function replaces a packaged dialog, and `false` turns it off; the menu rows that open it show as disabled with the reason. Open a dialog from your own control with `useEditorDialog(key).open()`. See [Dialogs](../../docs/site/content/customization/dialogs.mdx).

## Fonts

Bind font bytes to the `fonts` prop of `PptxEditor` or `PptxEditorRoot`, as in DOCX Editor, and listen to `fontError`. Call `useFonts(...origins)` in `<script setup>` for one stable value. A new `fonts` value replaces the faces in place with `setFonts`. See [Fonts](../../docs/site/content/customization/fonts.mdx).

## Docs and demo

- [Customization](../../docs/site/content/customization/index.mdx): the toolbar, ribbon, menus, dialogs, and layout, with React and Vue examples.
- [Toolbar](../../docs/site/content/customization/toolbar.mdx): the customization ladder (`className`, `icon`, `as-child`, `hidden`, slot overrides).
- Live demo: `bun run dev:vue` in the monorepo (`examples/vue`).

## License

Apache-2.0. See [LICENSE](LICENSE).
