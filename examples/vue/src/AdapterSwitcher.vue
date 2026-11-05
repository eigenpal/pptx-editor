<script setup lang="ts">
// The React | Vue switcher: the Vue twin of examples/vite/src/AdapterSwitcher.tsx. The preview
// deploy serves the React demo at `/react/` and the Vue demo at `/vue/`; in development the two
// demos run on their own ports (5274 and 5275). A navigation: the other demo is a link that keeps
// the page's query, and the current demo is marked `aria-current="page"`.
import { useTranslation } from '@pptx-editor.dev/vue';
import { adapterHref } from '../../shared/demo-files.ts';

defineProps<{ current: 'react' | 'vue' }>();
const { t } = useTranslation();

const DEV = import.meta.env.DEV;
const adapters = [
  { id: 'react', label: 'React', href: DEV ? 'http://localhost:5274/' : '/react/' },
  { id: 'vue', label: 'Vue', href: DEV ? 'http://localhost:5275/' : '/vue/' },
] as const;
const search = location.search;

// The query at the click, so a theme or deck chosen since the render comes along.
function follow(e: MouseEvent, href: string): void {
  (e.currentTarget as HTMLAnchorElement).setAttribute('href', adapterHref(href, location.search));
}
</script>

<template>
  <nav class="demo-adapter-switcher" :aria-label="t('app.adapter')">
    <template v-for="a in adapters" :key="a.id">
      <span v-if="current === a.id" class="demo-adapter-switcher__option" aria-current="page">{{
        a.label
      }}</span>
      <a
        v-else
        :href="adapterHref(a.href, search)"
        class="demo-adapter-switcher__option"
        @click="follow($event, a.href)"
        >{{ a.label }}</a
      >
    </template>
  </nav>
</template>
