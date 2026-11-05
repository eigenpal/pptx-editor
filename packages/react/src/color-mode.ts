// The chrome color mode a `PptxEditor.Root` declares (`colorMode`), after DOCX Editor's
// `packages/react/src/editor/color-mode-context.ts`. The dark palette is scoped to
// `.pptx-editor.dark`. `Root` renders no element, so the viewport adds `dark` beside
// `.pptx-editor`, and a part rendered outside the viewport (a dialog) puts the scope classes
// on its own root element. Only the chrome is themed: the slide keeps the deck's own colors.

import { createContext, useContext, useSyncExternalStore } from 'react';
import {
  resolveDarkChrome,
  subscribeSystemDark,
  systemPrefersDark,
  type ColorMode,
  type ResolvedColorMode,
} from '@pptx-editor.dev/core/editor';

/**
 * True when the enclosing `Root` resolved a dark chrome. A `Root` that resolves one provides
 * it; without a provider the chrome is light.
 */
export const ChromeDarkContext = createContext(false);

const light = (): boolean => false;
const none = (): (() => void) => () => undefined;

/** Whether `mode` is dark now. `'system'` follows the operating system and updates with it. */
export function useResolvedDark(mode: ColorMode | undefined): boolean {
  const follow = mode === 'system';
  const systemDark = useSyncExternalStore(
    follow ? subscribeSystemDark : none,
    follow ? systemPrefersDark : light,
    light
  );
  return resolveDarkChrome(mode, systemDark);
}

/** Whether the enclosing `Root` resolved a dark chrome. */
export function useChromeDark(): boolean {
  return useContext(ChromeDarkContext);
}

/** What `useEditorColorMode` returns. @public */
export interface UseEditorColorModeReturn {
  /** The resolved theme: `'dark'` or `'light'`. Re-renders when `colorMode` or the OS setting changes. */
  readonly colorMode: ResolvedColorMode;
}

// One object per theme, so the answer keeps its identity until the theme changes.
const LIGHT: UseEditorColorModeReturn = Object.freeze({ colorMode: 'light' });
const DARK: UseEditorColorModeReturn = Object.freeze({ colorMode: 'dark' });

/**
 * The theme the enclosing `PptxEditor.Root` resolved from its `colorMode`, so parts you build
 * yourself (a portaled popover) can match the chrome. `'system'` follows the operating system
 * and re-renders when it changes. Outside a root, or without `colorMode`, it is `'light'`. The
 * same hook as DOCX Editor's.
 *
 * @example
 * ```tsx
 * function MyPanel() {
 *   const { colorMode } = useEditorColorMode();
 *   return <aside className={colorMode === 'dark' ? 'panel panel--dark' : 'panel'} />;
 * }
 * ```
 *
 * @public
 */
export function useEditorColorMode(): UseEditorColorModeReturn {
  return useContext(ChromeDarkContext) ? DARK : LIGHT;
}

/** The scope classes a self-scoping part adds to its root element. */
export function useChromeScopeClassName(): 'pptx-editor' | 'pptx-editor dark' {
  return useContext(ChromeDarkContext) ? 'pptx-editor dark' : 'pptx-editor';
}
