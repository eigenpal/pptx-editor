// The loading screen: the Vue twin of the React adapter's `PptxEditor.Loading`, after DOCX
// Editor's `DocxEditor.Loading`, with the same markup and classes, so the one core stylesheet
// styles both. It reads the same snapshot as every other part (`showsLoading` in core) and
// renders nothing once the deck is ready.
//
// It carries `pptx-editor` itself, as `PptxEditorViewport` does, so the `--doc-*` tokens
// resolve wherever it is composed.

import {
  defineComponent,
  h,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import { loadingSheetSize, sameLoadingSheetSize, showsLoading } from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';
import { useChromeDark } from './color-mode.ts';

const loadingSpinnerProps = {
  /** A class name added to the spinner, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `PptxEditorLoadingSpinner`, the twin of React's `LoadingSpinnerProps`. @public */
export type LoadingSpinnerProps = ExtractPublicPropTypes<typeof loadingSpinnerProps>;

/**
 * The packaged spinner on its own, to compose into the default slot. It is decorative
 * (`aria-hidden`), so give the surrounding screen its own text. @public
 */
export const PptxEditorLoadingSpinner = defineComponent({
  name: 'PptxEditorLoadingSpinner',
  props: loadingSpinnerProps,
  setup(props) {
    return () =>
      h('span', { class: ['pptx-loading__spinner', props.className], 'aria-hidden': 'true' });
  },
});

const loadingProps = {
  /**
   * An extra host condition, OR-ed with the editor's own. Pass it only for work the editor
   * cannot see, such as a host that mounts the root after its own fetch.
   */
  when: { type: Boolean, default: false },
  /**
   * Pin the screen over the nearest positioned ancestor, so it covers the previous deck while
   * the next one opens. A short appearance delay keeps a fast open from flashing it.
   */
  overlay: { type: Boolean, default: false },
  /** A class name added to the loading container, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Inline styles for the loading container, as React's `style`. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `PptxEditorLoading`, the same name as React's `LoadingProps`. @public */
export type LoadingProps = ExtractPublicPropTypes<typeof loadingProps>;

const LoadingImpl = defineComponent({
  name: 'PptxEditorLoading',
  props: loadingProps,
  setup(props, { slots }) {
    const loading = useEditorState(showsLoading);
    const sheet = useEditorState(loadingSheetSize, sameLoadingSheetSize);
    const dark = useChromeDark();
    const { t } = useTranslation();
    return () => {
      if (!props.when && !loading.value) return null;
      const { width, height } = sheet.value;
      const status = () =>
        slots.indicator?.() ?? [
          h(PptxEditorLoadingSpinner),
          h('span', { class: 'pptx-sr-only' }, t('app.loading')),
        ];
      return h(
        'div',
        {
          class: [
            'pptx-editor',
            dark.value ? 'dark' : null,
            'pptx-loading',
            props.overlay ? 'pptx-loading--overlay' : null,
            props.className,
          ],
          style: props.style,
          role: 'status',
          'aria-live': 'polite',
        },
        slots.default?.() ??
          h(
            'div',
            {
              class: 'pptx-loading__sheet',
              style: { width: `${width}px`, aspectRatio: `${width} / ${height}` },
            },
            status()
          )
      );
    };
  },
});

/**
 * Shows a slide-shaped sheet with a spinner while the editor has no deck, and while a deck
 * opens; renders nothing once the deck is ready. A parse failure clears it. The `indicator`
 * slot replaces only the spinner; the default slot replaces the whole screen. The packaged
 * spinner is on it as `PptxEditorLoading.Spinner`, as React's `Loading.Spinner`.
 *
 * ```vue
 * <PptxEditorLoading overlay>
 *   <template #indicator><MySpinner label="Opening presentation" /></template>
 * </PptxEditorLoading>
 * ```
 *
 * @public
 */
export const PptxEditorLoading = Object.assign(LoadingImpl, {
  /** The packaged spinner, the same component as `PptxEditorLoadingSpinner`. */
  Spinner: PptxEditorLoadingSpinner,
});
