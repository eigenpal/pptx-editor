---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

Add the selection pane: the side panel's new **Selection** tab, and the `SelectionPane` component, list the objects of the current slide as a tree to select, show or hide, rename, and reorder them. `SidePanel` takes `tab` and `onTabChange` (`v-model:tab` in Vue) to open a tab from your own control. The toolbar's Selection pane and Format shape controls call the new `requestPane`, which a side panel and your host answer through `listenPaneRequests`. Both adapters export the two functions. The Vue adapter requires Vue 3.5 or later, as its components use `useId`.
