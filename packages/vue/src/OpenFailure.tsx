/** @jsxImportSource vue */
// The open-failure notice: the Vue twin of the React adapter's `OpenFailure.tsx`, after DOCX
// Editor's `DocxEditorOpenFailure`, with the same markup and classes. It says why the editor
// refused a file, and when the deck that was open stays open. It reads
// `snapshot().openFailure`; the reason groups and their messages are core's (`loading.ts`).
//
// The default text is generic on purpose: it explains the refusal and shows the stable code,
// and names no contact. A host adds its own support action through the `actions` slot.

import {
  computed,
  defineComponent,
  shallowRef,
  useId,
  type ComputedRef,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type SlotsType,
  type VNodeChild,
} from 'vue';
import {
  OPEN_FAILURE_MESSAGE_KEYS,
  shownOpenFailure,
  type DocumentOpenFailure,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';

/** What `useOpenFailure` returns, the twin of React's `UseOpenFailureReturn`. @public */
export interface UseOpenFailureReturn {
  /** The refused open, or `null` when there is none or the person dismissed it. */
  readonly failure: ComputedRef<DocumentOpenFailure | null>;
  /** The packaged, translated message for `failure`. */
  readonly message: ComputedRef<string | null>;
  /** Whether the deck that was open before stays open. */
  readonly kept: ComputedRef<boolean>;
  /**
   * Hide this refusal. A later refusal shows again, even for the same file. Pass the failure
   * your UI showed, so a late call cannot hide a newer refusal.
   */
  dismiss(failure?: DocumentOpenFailure | null): void;
}

/**
 * Reads the file the editor refused to open, for a host that shows it in its own UI, such as
 * a toast. Pair it with `:open-failure="false"` on `PptxEditor`, so the packaged notice does
 * not show as well. Call it inside a `PptxEditorRoot`. @public
 */
export function useOpenFailure(): UseOpenFailureReturn {
  const current = useEditorState(shownOpenFailure);
  const ready = useEditorState((s) => s.status === 'ready');
  const { t } = useTranslation();
  // Each refusal is a new object, so remembering the dismissed one by reference lets the next
  // refusal of the same file show again.
  const dismissed = shallowRef<DocumentOpenFailure | null>(null);
  const failure = computed(() =>
    current.value !== null && current.value !== dismissed.value ? current.value : null
  );
  return {
    failure,
    message: computed(() =>
      failure.value ? t(OPEN_FAILURE_MESSAGE_KEYS[failure.value.reason]) : null
    ),
    kept: computed(() => ready.value),
    dismiss(shown) {
      if (!shown || shown === current.value) dismissed.value = current.value;
    },
  };
}

/** What the `actions` slot receives, the twin of React's `OpenFailureContext`. @public */
export interface OpenFailureContext {
  /** The refused open: its stable `code` and its `reason`. */
  readonly failure: DocumentOpenFailure;
  /** The packaged, translated message for the failure. */
  readonly message: string;
  /** Close the notice. It emits `dismiss`, as **Close** does. */
  dismiss(): void;
}

const openFailureProps = {
  /** The notice title. Defaults to the translated `openFailure.title`. */
  title: { type: String, default: undefined },
  /**
   * Your own message in place of the packaged one, as content or per failure. The `message`
   * slot wins over it.
   */
  message: {
    type: [String, Number, Object, Array, Function] as PropType<
      VNodeChild | ((failure: DocumentOpenFailure) => VNodeChild)
    >,
    default: undefined,
  },
  /** Show the stable error code under the message. Defaults to `true`. */
  showCode: { type: Boolean, default: true },
  /**
   * Your own buttons, before **Close**: the function form of the `actions` slot, for the
   * `openFailure` object on `PptxEditor`. The slot wins over it. The notice shows no contact of
   * its own.
   */
  actions: {
    type: Function as PropType<(context: OpenFailureContext) => VNodeChild>,
    default: undefined,
  },
  /** The label of the **Close** button. Defaults to the translated `openFailure.close`. */
  closeLabel: { type: String, default: undefined },
  /** Show the packaged **Close** button. Defaults to `true`. */
  showClose: { type: Boolean, default: true },
  /** A class name added to the notice, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Inline styles for the notice, as React's `style`. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/**
 * Props of `PptxEditorOpenFailure`, the same name as React's. `onDismiss` is the listener of
 * the `dismiss` event, so the `openFailure` object on `PptxEditor` can carry it. @public
 */
export type PptxEditorOpenFailureProps = ExtractPublicPropTypes<typeof openFailureProps> & {
  /** Listener for the `dismiss` event: the person closed the notice, by any route. */
  readonly onDismiss?: (failure: DocumentOpenFailure) => void;
};

/**
 * A notice that explains why a file did not open, the twin of React's `PptxEditor.OpenFailure`.
 * It shows nothing until the editor refuses a file, and shows again for each later refusal.
 * `PptxEditor` renders it by default. The `actions` slot (or prop) adds buttons before
 * **Close**, and the `message` slot (or prop) replaces the message. It emits `dismiss` after the
 * person closes it:
 *
 * ```vue
 * <PptxEditorOpenFailure>
 *   <template #actions="{ failure }">
 *     <a class="pptx-open-failure__action" :href="`/support?code=${failure.code}`">Get help</a>
 *   </template>
 * </PptxEditorOpenFailure>
 * ```
 *
 * @public
 */
export const PptxEditorOpenFailure = defineComponent({
  name: 'PptxEditorOpenFailure',
  props: openFailureProps,
  emits: {
    /** The person closed the notice, by any route. */
    dismiss: (_failure: DocumentOpenFailure) => true,
  },
  slots: Object as SlotsType<{
    /** Your own buttons, before **Close**. It wins over the `actions` prop. */
    actions?: (context: OpenFailureContext) => VNodeChild;
    /** Your own message in place of the packaged one. It wins over the `message` prop. */
    message?: (context: OpenFailureContext) => VNodeChild;
  }>,
  setup(props, { emit, slots }) {
    const { t } = useTranslation();
    const { failure, message, kept, dismiss } = useOpenFailure();
    const titleId = useId();
    return () => {
      const shown = failure.value;
      const packaged = message.value;
      if (!shown || packaged === null) return null;
      const close = (): void => {
        dismiss(shown);
        emit('dismiss', shown);
      };
      const context: OpenFailureContext = { failure: shown, message: packaged, dismiss: close };
      return (
        <div
          class={['pptx-open-failure', props.className]}
          style={props.style}
          role="alert"
          aria-labelledby={titleId}
          data-reason={shown.reason}
        >
          <div class="pptx-open-failure__text">
            <p id={titleId} class="pptx-open-failure__title">
              {props.title ?? t('openFailure.title')}
            </p>
            <div class="pptx-open-failure__message">
              {slots.message
                ? slots.message(context)
                : typeof props.message === 'function'
                  ? props.message(shown)
                  : (props.message ?? packaged)}
            </div>
            {kept.value ? <p class="pptx-open-failure__kept">{t('openFailure.kept')}</p> : null}
            {props.showCode ? (
              <p class="pptx-open-failure__code">{t('openFailure.code', { code: shown.code })}</p>
            ) : null}
          </div>
          <div class="pptx-open-failure__actions">
            {slots.actions ? slots.actions(context) : props.actions?.(context)}
            {props.showClose ? (
              <button type="button" class="pptx-open-failure__action" onClick={close}>
                {props.closeLabel ?? t('openFailure.close')}
              </button>
            ) : null}
          </div>
        </div>
      );
    };
  },
});
