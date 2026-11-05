/** @jsxImportSource vue */
// The Picture Format controls whose look is more than a list: the Corrections gallery, a grid
// with a row per contrast and a column per brightness, each cell a sample picture with that
// correction. The Vue twin of the React adapter's `PictureControls.tsx`, on the same core rules
// (`chrome-picture.ts`, `picture-corrections.ts`).

import { defineComponent, ref, type PropType } from 'vue';
import {
  chromeOptionChecked,
  correctionGridRows,
  correctionPercent,
  correctionSwatch,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** One Corrections sample: the sample picture with a brightness and a contrast. */
const CorrectionSwatch = defineComponent({
  name: 'PptxCorrectionSwatch',
  props: {
    brightness: { type: Number, required: true as const },
    contrast: { type: Number, required: true as const },
  },
  setup(props) {
    return () => (
      <svg class="pptx-corrections-grid__swatch" viewBox="0 0 40 28" aria-hidden="true">
        {correctionSwatch(props.brightness, props.contrast).map((p) => (
          <path key={p.d} d={p.d} fill={p.fill} />
        ))}
      </svg>
    );
  },
});

/** The toolbar control of `image.corrections`: a trigger and the grid in a popover. @internal */
export const CorrectionsControl = defineComponent({
  name: 'PptxCorrectionsControl',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    control: { type: Object as PropType<ChromeControl>, required: true as const },
  },
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    // A disabled picker stays in the arrow order and says why.
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const { t } = useTranslation();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    const trigger = useTriggerAttributes(
      props.slot,
      'pptx-toolbar__button pptx-toolbar__button--wide',
      () => ({ disabled: !cmd.isEnabled.value })
    );
    const pick = (value: string): void => {
      cmd.execute(value);
      open.value = false;
    };
    return () => {
      const rows = correctionGridRows(props.control.options ?? []);
      return [
        <button
          ref={anchor}
          type="button"
          {...trigger()}
          data-slot={props.slot}
          aria-label={label.value}
          aria-haspopup="menu"
          aria-expanded={open.value}
          {...tip.value}
          {...off.attrs.value}
          onMousedown={keepFocus}
          onClick={() => (open.value = !open.value)}
        >
          <Icon paths={props.control.paths} />
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>,
        off.note(),
        open.value ? (
          <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
            {{
              default: () => (
                <div
                  class="pptx-style-gallery pptx-corrections-grid"
                  role="group"
                  aria-label={t('pictureFormat.brightnessContrast')}
                  // A grid for the arrow keys: Up and Down move by a row of brightness values.
                  data-columns={rows[0]?.cells.length}
                >
                  {/* The axis labels: each cell names its values, so they are hidden from assistive technology. */}
                  <div class="pptx-corrections-grid__title" aria-hidden="true">
                    {t('pictureFormat.brightness')}
                  </div>
                  <div class="pptx-corrections-grid__head" aria-hidden="true">
                    {rows[0]?.cells.map(({ brightness }) => (
                      <span key={brightness} class="pptx-corrections-grid__label">
                        {t('pictureFormat.percent', { value: correctionPercent(brightness) })}
                      </span>
                    ))}
                  </div>
                  <div class="pptx-corrections-grid__body" role="presentation">
                    <div class="pptx-corrections-grid__side" aria-hidden="true">
                      {t('pictureFormat.contrast')}
                    </div>
                    <div class="pptx-corrections-grid__rows" role="presentation">
                      {rows.map(({ contrast, cells }) => (
                        <div key={contrast} class="pptx-style-gallery__row" role="presentation">
                          <span
                            class="pptx-corrections-grid__label pptx-corrections-grid__label--row"
                            aria-hidden="true"
                          >
                            {t('pictureFormat.percent', { value: correctionPercent(contrast) })}
                          </span>
                          {cells.map(({ option: o, brightness, contrast: c }) => {
                            const name = t(o.labelKey as TranslationKey, o.labelParams);
                            return (
                              <button
                                key={o.value}
                                type="button"
                                role="menuitemradio"
                                aria-checked={chromeOptionChecked(cmd.value.value, o.value)}
                                tabindex={-1}
                                class="pptx-menu__item pptx-style-gallery__cell"
                                aria-label={name}
                                data-pptx-tooltip={name}
                                data-slot={props.slot}
                                data-value={o.value}
                                onClick={() => pick(o.value)}
                              >
                                <CorrectionSwatch brightness={brightness} contrast={c} />
                              </button>
                            );
                          })}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              ),
            }}
          </Popover>
        ) : null,
      ];
    };
  },
});
