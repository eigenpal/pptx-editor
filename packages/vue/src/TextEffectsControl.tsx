/** @jsxImportSource vue */
// Text Effects (`shape.textEffects`): a Transform submenu with No Transform and the text warp
// gallery, one section for the warps that follow a path and one for the warps between paths.
// The Vue twin of the React adapter's `TextEffectsControl.tsx`, on the same core rules
// (`chrome-text-effects.ts`): each cell strokes the warp's own paths (`textWarpPreview`).

import { defineComponent, ref, type PropType } from 'vue';
import {
  TEXT_WARP_GALLERY,
  chromeOptionChecked,
  textWarpPreview,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { MenuRow, MenuSubmenu } from './menu-row.tsx';
import { useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** The Text Effects trigger: a letter over an arched baseline, drawn, not an icon path. */
const textEffectsGlyph = () => (
  <svg class="pptx-icon pptx-text-effects-glyph" viewBox="0 0 24 24" aria-hidden="true">
    <path d="M7 15 11 4h2l4 11h-2.2l-.9-2.6H10L9.2 15Zm3.6-4.4h2.8L12 6.5Z" />
    <path d="M3 21q9-6 18 0" class="pptx-text-effects-glyph__arc" />
  </svg>
);

/** One gallery cell: the warp's paths, stroked. */
const warpSwatch = (preset: string) => (
  <svg class="pptx-style-gallery__swatch" viewBox="0 0 40 28" aria-hidden="true">
    {textWarpPreview(preset).map((d, i) => (
      <path key={i} d={d} class="pptx-warp-gallery__path" />
    ))}
  </svg>
);

/**
 * The Transform gallery: a heading and a grid of cells for each section. Up and Down move by a
 * row of cells (`data-columns`).
 */
const WarpGallery = defineComponent({
  name: 'PptxWarpGallery',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    value: { type: [String, Number] as PropType<string | number>, default: undefined },
  },
  emits: { pick: (_value: string) => true },
  setup(props, { emit }) {
    const { t } = useTranslation();
    const columns = 5;
    return () =>
      TEXT_WARP_GALLERY.map(({ section, labelKey, warps }) => (
        <div
          key={section}
          class="pptx-style-gallery pptx-warp-gallery"
          role="group"
          aria-label={t(labelKey)}
          data-columns={columns}
        >
          <div class="pptx-menu__heading" aria-hidden="true">
            {t(labelKey)}
          </div>
          {Array.from({ length: Math.ceil(warps.length / columns) }, (_, r) => (
            <div key={r} class="pptx-style-gallery__row" role="presentation">
              {warps.slice(r * columns, (r + 1) * columns).map((w) => {
                const name = t(`toolbar.textWarps.${w}` as TranslationKey);
                const option = `transform.${w}`;
                return (
                  <button
                    key={w}
                    type="button"
                    role="menuitemradio"
                    aria-checked={chromeOptionChecked(props.value, option)}
                    tabindex={-1}
                    class="pptx-menu__item pptx-style-gallery__cell"
                    aria-label={name}
                    data-pptx-tooltip={name}
                    data-slot={props.slot}
                    data-value={option}
                    onClick={() => emit('pick', option)}
                  >
                    {warpSwatch(w)}
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      ));
  },
});

/**
 * The Text Effects menu (`shape.textEffects`): the Transform submenu, with the selection's warp
 * checked. @internal
 */
export const TextEffectsMenuControl = defineComponent({
  name: 'PptxTextEffectsMenuControl',
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
    const close = (): void => {
      open.value = false;
    };
    const pick = (value: string): void => {
      cmd.execute(value);
      close();
    };
    return () => [
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
        {textEffectsGlyph()}
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>,
      off.note(),
      open.value ? (
        <Popover anchor={anchor.value} onClose={close}>
          {{
            default: () => (
              <>
                {props.control.optionGroups?.map((group) => (
                  <MenuSubmenu key={group.id} label={t(group.labelKey as TranslationKey)}>
                    {{
                      default: () => [
                        <MenuRow
                          slot={props.slot}
                          label={t('toolbar.noTransform')}
                          checked={chromeOptionChecked(cmd.value.value, 'transform.none')}
                          radio
                          onSelect={() => pick('transform.none')}
                        />,
                        <WarpGallery
                          slot={props.slot}
                          value={cmd.value.value as string | undefined}
                          onPick={pick}
                        />,
                      ],
                    }}
                  </MenuSubmenu>
                ))}
              </>
            ),
          }}
        </Popover>
      ) : null,
    ];
  },
});
