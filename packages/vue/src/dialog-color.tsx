/** @jsxImportSource vue */
// The color field of the form dialogs (Font color, Underline color, the bullet color): the Vue
// twin of the React adapter's `dialog-color.tsx`. A list of the inherited color, "no change",
// and the theme colors, and beside it a swatch button that shows the current color and opens
// the color grid. The picks, labels, and the color OK sends live in core (`color-fields.ts`,
// `chrome-forms.ts`).

import { defineComponent, ref, watch, type PropType } from 'vue';
import {
  colorChoiceOf,
  colorPickLabel,
  colorPickOf,
  colorPicks,
  renderedColorOf,
  type ColorPick,
} from '@pptx-editor.dev/core/editor';
import { ColorSwatches } from './ColorSwatches.tsx';
import { useEditorState, useTranslation } from './composables.ts';
import { SelectField, dialogControl } from './dialog-form.tsx';

/**
 * A color field: the list of picks and the swatch grid beside it. Other attributes (`class`,
 * `style`, `data-*`) reach the outer element. @internal
 */
export const ColorField = dialogControl(
  defineComponent({
    name: 'PptxColorField',
    props: {
      name: { type: String, required: true as const },
      label: { type: String, required: true as const },
      value: { type: String as PropType<ColorPick>, required: true as const },
      /** The value the field opened with: `keep` and `mixed` are offered only then. */
      seed: { type: String as PropType<ColorPick>, default: undefined },
      /** The color the selection has now (`#RRGGBB`): the swatch of "Current color" shows it. */
      current: { type: String as PropType<string | null>, default: null },
      disabled: { type: Boolean, default: false },
      className: { type: String, default: undefined },
      onChange: { type: Function as PropType<(pick: ColorPick) => void>, required: true as const },
    },
    setup(props) {
      const { t } = useTranslation();
      const rows = useEditorState((s) => s.themeColors);
      const open = ref(false);
      const panel = ref<HTMLDivElement | null>(null);
      const button = ref<HTMLButtonElement | null>(null);
      // The grid closes on a pick, on Escape (which does not close the dialog), and on a press
      // outside it.
      const away = (e: PointerEvent): void => {
        const target = e.target as Node;
        if (!panel.value?.contains(target) && !button.value?.contains(target)) open.value = false;
      };
      watch(
        open,
        (now, _old, onCleanup) => {
          if (!now) return;
          document.addEventListener('pointerdown', away, true);
          panel.value?.querySelector<HTMLElement>('button')?.focus();
          onCleanup(() => document.removeEventListener('pointerdown', away, true));
        },
        { flush: 'post' }
      );
      const closeGrid = (): void => {
        open.value = false;
        button.value?.focus();
      };
      return () => {
        const rendered = renderedColorOf(props.value, rows.value, props.current);
        const moreLabel = t('colors.moreColors', { field: props.label });
        return (
          <div class={['pptx-color-field', props.className].filter(Boolean).join(' ')}>
            <SelectField
              name={props.name}
              label={props.label}
              disabled={props.disabled}
              value={props.value}
              options={colorPicks(props.seed, props.value).map(
                (p) => [p, colorPickLabel(t, p)] as const
              )}
              onChange={props.onChange}
            />
            <button
              ref={button}
              type="button"
              class="pptx-swatch pptx-color-field__button"
              aria-label={moreLabel}
              aria-haspopup="true"
              aria-expanded={open.value}
              disabled={props.disabled}
              data-color={rendered ?? undefined}
              style={rendered ? { background: rendered } : undefined}
              onClick={() => {
                open.value = !open.value;
              }}
            />
            {open.value ? (
              <div
                ref={panel}
                role="group"
                aria-label={moreLabel}
                class="pptx-color-field__panel"
                onKeydown={(e: KeyboardEvent) => {
                  if (e.key !== 'Escape') return;
                  e.preventDefault();
                  e.stopPropagation();
                  closeGrid();
                }}
              >
                <ColorSwatches
                  menu={false}
                  selected={colorChoiceOf(props.value)}
                  onPick={(choice) => {
                    props.onChange(colorPickOf(choice));
                    closeGrid();
                  }}
                />
              </div>
            ) : null}
          </div>
        );
      };
    },
  })
);
