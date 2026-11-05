// The chrome color mode a `PptxEditorRoot` declares (`colorMode`): the Vue twin of the React
// adapter's `color-mode.ts`. The viewport adds `dark` beside `.pptx-editor`, and a dialog
// outside the viewport puts the scope classes on its own element.

import {
  computed,
  inject,
  onMounted,
  onScopeDispose,
  provide,
  shallowRef,
  watch,
  type ComputedRef,
  type InjectionKey,
} from 'vue';
import {
  resolveDarkChrome,
  subscribeSystemDark,
  systemPrefersDark,
  type ColorMode,
  type ResolvedColorMode,
} from '@pptx-editor.dev/core/editor';

const chromeDarkKey: InjectionKey<ComputedRef<boolean>> = Symbol('pptxChromeDark');
const lightChrome = computed(() => false);

/** Resolve and provide the chrome color mode for one root. @internal */
export function provideChromeDark(mode: () => ColorMode | undefined): void {
  const systemDark = shallowRef(false);
  let off: (() => void) | null = null;
  const follow = (next: ColorMode | undefined): void => {
    off?.();
    off = null;
    if (next !== 'system') return;
    systemDark.value = systemPrefersDark();
    off = subscribeSystemDark(() => (systemDark.value = systemPrefersDark()));
  };
  // The operating system is read after mount, as DOCX Editor's Vue adapter does: a server
  // render and the hydrating client then start from the same light chrome.
  onMounted(() => follow(mode()));
  watch(mode, follow);
  onScopeDispose(() => off?.());
  provide(
    chromeDarkKey,
    computed(() => resolveDarkChrome(mode(), systemDark.value))
  );
}

/** Whether the enclosing root resolved a dark chrome. @internal */
export function useChromeDark(): ComputedRef<boolean> {
  return inject(chromeDarkKey, lightChrome);
}

/** What `useEditorColorMode` returns. @public */
export interface UseEditorColorModeReturn {
  /** The resolved theme: `'dark'` or `'light'`. Updates when `colorMode` or the OS setting changes. */
  readonly colorMode: ComputedRef<ResolvedColorMode>;
}

/**
 * The theme the enclosing `PptxEditorRoot` resolved from its `colorMode`, as React's
 * `useEditorColorMode` and DOCX Editor's. Outside a root, or without `colorMode`, it is
 * `'light'`. Call it in `setup`. @public
 */
export function useEditorColorMode(): UseEditorColorModeReturn {
  const dark = useChromeDark();
  return { colorMode: computed(() => (dark.value ? 'dark' : 'light')) };
}
