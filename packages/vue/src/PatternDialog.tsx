/** @jsxImportSource vue */
// The pattern fill dialog: the 54 preset patterns, the foreground and background colors, and a
// preview. The Vue twin of the React adapter's `PatternDialog.tsx`: the same markup, classes,
// and ARIA, on the same core rules (`pattern-fields.ts`).

import {
  defineComponent,
  onMounted,
  ref,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  dialogInitialFocus,
  PATTERN_CHOICES,
  patternCss,
  patternDialogStart,
  patternLabel,
  type PatternFill,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useTranslation } from './composables.ts';
import { Swatch, SwatchGroup } from './ColorSwatches.tsx';

const patternDialogProps = {
  /** Whether the dialog shows. */
  open: { type: Boolean, required: true as const },
  /** The pattern the dialog starts from, or null for 50% in the theme's first accent. */
  initial: { type: Object as PropType<PatternFill | null>, default: null },
  /** The foreground of a new pattern, as `#RRGGBB`. */
  base: { type: String, default: undefined },
  /** Called when the dialog closes, applied or not. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** Called with the pattern the person applied. */
  onApply: {
    type: Function as PropType<(pattern: PatternFill) => void>,
    required: true as const,
  },
} as const;

/** Props of `PatternDialog`, the twin of React's `PatternDialogProps`. @public */
export type PatternDialogProps = ExtractPublicPropTypes<typeof patternDialogProps>;

/** The pattern fill dialog: a preset, its two colors, and a preview. @public */
export const PatternDialog = defineComponent({
  name: 'PatternDialog',
  props: patternDialogProps,
  setup(props) {
    const { t } = useTranslation();
    const dialogEl = ref<HTMLDialogElement | null>(null);
    const p = shallowRef<PatternFill>(patternDialogStart(props.initial, props.base));
    const onClose = (): void => props.onClose();
    // Open and close the modal with `open`, after the element is in the document.
    const sync = (): void => {
      const el = dialogEl.value;
      if (!el) return;
      if (props.open) {
        p.value = patternDialogStart(props.initial, props.base);
        if (!el.open) el.showModal();
        // The chosen pattern takes focus, as the first field of the dialog.
        const chosen = el.querySelector<HTMLElement>('.pptx-swatch[aria-pressed="true"]');
        (chosen ?? dialogInitialFocus(el))?.focus({ preventScroll: true });
      } else if (el.open) el.close();
    };
    onMounted(sync);
    watch(() => [props.open, props.initial, props.base] as const, sync, { flush: 'post' });
    const color = (side: 'foreground' | 'background') => (
      <label class="pptx-field">
        {t(side === 'foreground' ? 'pattern.foreground' : 'pattern.background')}
        <input
          type="color"
          class="pptx-gradient__color"
          value={p.value[side].toLowerCase()}
          onInput={(e: Event) => {
            const v = (e.target as HTMLInputElement).value.toUpperCase();
            p.value = { ...p.value, [side]: v };
          }}
        />
      </label>
    );
    const body = () => {
      const cur = p.value;
      return [
        <div class="pptx-dialog__header">
          <h2 id="pptx-pattern-title" class="pptx-dialog__title">
            {t('pattern.title')}
          </h2>
          <ToolButton label={t('pattern.close')} icon="close" onClick={onClose} />
        </div>,
        <div class="pptx-dialog__body">
          <SwatchGroup label={t('pattern.patterns')} columns={9} roving>
            {{
              default: () =>
                PATTERN_CHOICES.map((preset) => (
                  <Swatch
                    key={preset}
                    color={cur.foreground}
                    background={patternCss({ ...cur, preset })}
                    name={patternLabel(t, preset)}
                    menu={false}
                    pressed={cur.preset === preset}
                    onPick={() => (p.value = { ...p.value, preset })}
                  />
                )),
            }}
          </SwatchGroup>
          <div class="pptx-pattern__row">
            {color('foreground')}
            {color('background')}
            <div class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('pattern.preview')}</h3>
              <div
                class="pptx-pattern__preview"
                role="img"
                aria-label={patternLabel(t, cur.preset)}
                style={{ background: patternCss(cur, 12) }}
              />
            </div>
          </div>
        </div>,
        <div class="pptx-dialog__footer">
          <button type="button" class="pptx-button" onClick={onClose}>
            {t('dialog.cancel')}
          </button>
          <button
            type="button"
            class="pptx-button pptx-button--primary"
            onClick={() => props.onApply(p.value)}
          >
            {t('dialog.ok')}
          </button>
        </div>,
      ];
    };
    return () => (
      <dialog
        ref={dialogEl}
        class="pptx-dialog pptx-pattern"
        data-pptx-dialog="pattern"
        aria-labelledby="pptx-pattern-title"
        onCancel={(e: Event) => {
          e.preventDefault();
          onClose();
        }}
        onClick={(e: MouseEvent) => {
          if (e.target === e.currentTarget) onClose();
        }}
      >
        {props.open ? body() : null}
      </dialog>
    );
  },
});
