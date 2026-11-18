// The Vue binding of core's side-panel field controller (`live-field.ts`). The React twin is
// `packages/react/src/live-field.ts`.
import { onBeforeUnmount, shallowRef, watch } from 'vue';
import { createLiveField, type LiveField, type LiveFieldHost } from '@pptx-editor.dev/core/editor';

/**
 * Bind one side-panel field. `scope` reads the document value and any key that drops the draft
 * when it changes, unless the field's own edit made the change. Reading `draft` or `error` in a
 * render tracks them. @internal
 */
export function useLiveField(host: () => LiveFieldHost, scope: () => unknown): LiveField {
  const version = shallowRef(0);
  const field = createLiveField(host, () => {
    version.value++;
  });
  watch(
    scope,
    () => {
      field.sync();
      version.value++;
    },
    { flush: 'sync' }
  );
  onBeforeUnmount(() => field.dispose());
  return {
    get draft() {
      void version.value;
      return field.draft;
    },
    get error() {
      void version.value;
      return field.error;
    },
    input: (text, composing) => field.input(text, composing),
    commit: () => field.commit(),
    cancel: () => field.cancel(),
    sync: () => field.sync(),
    dispose: () => field.dispose(),
  };
}
