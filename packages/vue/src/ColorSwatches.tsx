/** @jsxImportSource vue */
// The swatch sections of every color menu and the background dialog: the deck's theme colors
// (ten columns, each with its lighter and darker variants), the Standard Colors row, and a
// custom color (a hex field and the browser's color selector). The Vue twin of the React
// adapter's `ColorSwatches.tsx`, on the same core rules (`chrome-forms.ts`, `color-fields.ts`).

import {
  computed,
  defineComponent,
  onMounted,
  onUpdated,
  ref,
  shallowRef,
  type ComputedRef,
  type PropType,
} from 'vue';
import {
  STANDARD_COLORS,
  customHexColor,
  customHexCurrent,
  customHexDraft,
  customHexSeed,
  swatchGridKeyDown,
  swatchName,
  swatchSelection,
  syncSwatchTabStops,
  themeBaseSwatches,
  themeSwatchChoice,
  type ColorChoice,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './composables.ts';

/** One named swatch. In a menu it is a `menuitem`; in a dialog, a button that shows its state. @internal */
export const Swatch = defineComponent({
  name: 'PptxSwatch',
  props: {
    color: { type: String, required: true as const },
    name: { type: String, required: true as const },
    menu: { type: Boolean, required: true as const },
    pressed: { type: Boolean, default: undefined },
    /** CSS background when it is not the color itself (a gradient preview). */
    background: { type: String, default: undefined },
    onPick: { type: Function as PropType<() => void>, required: true as const },
  },
  setup(props) {
    return () => (
      <button
        type="button"
        class="pptx-swatch"
        style={{ background: props.background ?? props.color }}
        aria-label={props.name}
        data-pptx-tooltip={props.name}
        // In a menu the current color is marked with `aria-current`; in a dialog it is pressed.
        {...(props.menu
          ? { role: 'menuitem', tabindex: -1, ...(props.pressed ? { 'aria-current': true } : {}) }
          : { 'aria-pressed': !!props.pressed })}
        onClick={props.onPick}
      />
    );
  },
});

/**
 * A labelled swatch section; `columns` lets the menu's arrow keys move by row. With `roving`
 * (a dialog), the section is one Tab stop, on the chosen swatch or the first, and the arrow
 * keys move by column and row, as a WAI-ARIA grid of buttons does. @internal
 */
export const SwatchGroup = defineComponent({
  name: 'PptxSwatchGroup',
  props: {
    label: { type: String, required: true as const },
    columns: { type: Number, default: 10 },
    className: { type: String, default: undefined },
    roving: { type: Boolean, default: false },
  },
  setup(props, { slots }) {
    const grid = ref<HTMLDivElement | null>(null);
    // After each render: the focused swatch keeps the stop, else the chosen one, else the first.
    const sync = (): void => {
      if (props.roving) syncSwatchTabStops(grid.value);
    };
    onMounted(sync);
    onUpdated(sync);
    return () => (
      <div role="group" aria-label={props.label} class="pptx-swatches__section">
        <p class="pptx-swatches__title" aria-hidden="true">
          {props.label}
        </p>
        <div
          ref={grid}
          class={`pptx-swatches${props.className ? ` ${props.className}` : ''}`}
          data-columns={props.columns}
          onKeydown={(e: KeyboardEvent) => {
            if (props.roving) swatchGridKeyDown(grid.value, e, props.columns);
          }}
        >
          {slots.default?.()}
        </div>
      </div>
    );
  },
});

/**
 * The Custom Color section, as React's `CustomColor`: a hex field with Apply, and the browser's
 * color selector, which applies when it closes. In a menu, the fields join the arrow keys
 * (`data-menu-field`) and Apply is a row. @internal
 */
export const CustomColor = defineComponent({
  name: 'PptxCustomColor',
  props: {
    onPick: { type: Function as PropType<(color: string) => void>, required: true as const },
    menu: { type: Boolean, required: true as const },
    selected: {
      type: [String, Object] as PropType<ColorChoice | null>,
      default: undefined,
    },
  },
  setup(props) {
    const { t } = useTranslation();
    const draft = shallowRef(customHexSeed(props.selected));
    // An empty field shows the current color: as its placeholder and in the selector.
    const rows = useEditorState((s) => s.themeColors);
    const apply = (): void => {
      const color = customHexColor(draft.value);
      if (color) props.onPick(color);
    };
    return () => {
      const color = customHexColor(draft.value);
      const current = customHexCurrent(props.selected, rows.value);
      const label = t('colors.customColor');
      return (
        <div role="group" aria-label={label} class="pptx-swatches__section">
          <p class="pptx-swatches__title" aria-hidden="true">
            {label}
          </p>
          <div class="pptx-custom-color">
            <input
              type="color"
              class="pptx-custom-color__picker"
              aria-label={t('colors.customPicker')}
              data-menu-field=""
              tabindex={props.menu ? -1 : undefined}
              value={(color ?? `#${current}`).toLowerCase()}
              // `input` follows the selector while it moves; `change` applies the color it closed on.
              onInput={(e: Event) => {
                draft.value = customHexDraft((e.target as HTMLInputElement).value);
              }}
              onChange={(e: Event) => {
                const picked = customHexColor((e.target as HTMLInputElement).value);
                if (picked) props.onPick(picked);
              }}
            />
            <span class="pptx-custom-color__hash" aria-hidden="true">
              #
            </span>
            <input
              type="text"
              class="pptx-custom-color__hex"
              aria-label={t('colors.customHex')}
              data-menu-field=""
              tabindex={props.menu ? -1 : undefined}
              placeholder={current}
              maxlength={6}
              spellcheck={false}
              autocomplete="off"
              value={draft.value}
              onInput={(e: Event) => {
                const el = e.target as HTMLInputElement;
                draft.value = customHexDraft(el.value);
                // Keep the field to what the draft holds, as React's controlled field does.
                el.value = draft.value;
              }}
              onKeydown={(e: KeyboardEvent) => {
                if (e.key !== 'Enter') return;
                // The Enter that applies is this field's: a dialog does not apply on it too.
                e.preventDefault();
                e.stopPropagation();
                apply();
              }}
            />
            <button
              type="button"
              class="pptx-button pptx-custom-color__apply"
              {...(props.menu ? { role: 'menuitem', tabindex: -1 } : {})}
              disabled={!color}
              onClick={apply}
            >
              {t('colors.customApply')}
            </button>
          </div>
        </div>
      );
    };
  },
});

/**
 * The theme colors and the standard colors. A theme swatch picks a theme color, which the
 * engine writes as `a:schemeClr`; a standard swatch and a custom color pick their `#RRGGBB`.
 * `onPick` also gets the color as rendered. `selected` marks the current color. @internal
 */
export const ColorSwatches = defineComponent({
  name: 'PptxColorSwatches',
  props: {
    onPick: {
      type: Function as PropType<(choice: ColorChoice, color: string) => void>,
      required: true as const,
    },
    menu: { type: Boolean, default: true },
    selected: {
      type: [String, Object] as PropType<ColorChoice | null>,
      default: undefined,
    },
  },
  setup(props) {
    const { t } = useTranslation();
    const rows = useEditorState((s) => s.themeColors);
    return () => {
      const isSelected = swatchSelection(props.selected, rows.value);
      return [
        rows.value.length ? (
          <SwatchGroup
            label={t('colors.themeColors')}
            className="pptx-swatches--theme"
            roving={!props.menu}
          >
            {{
              default: () =>
                rows.value.flatMap((row, r) =>
                  row.map((s, c) => (
                    <Swatch
                      key={`${r}-${c}`}
                      color={s.color}
                      name={swatchName(t, s)}
                      menu={props.menu}
                      pressed={isSelected(s.color, s.slot, s.shift)}
                      onPick={() => props.onPick(themeSwatchChoice(s), s.color)}
                    />
                  ))
                ),
            }}
          </SwatchGroup>
        ) : null,
        <SwatchGroup label={t('colors.standardColors')} roving={!props.menu}>
          {{
            default: () =>
              STANDARD_COLORS.map((s) => (
                <Swatch
                  key={s.name}
                  color={s.color}
                  name={t(`colors.standard.${s.name}`)}
                  menu={props.menu}
                  pressed={isSelected(s.color)}
                  onPick={() => props.onPick(s.color, s.color)}
                />
              )),
          }}
        </SwatchGroup>,
        <CustomColor
          menu={props.menu}
          selected={props.selected}
          onPick={(color: string) => props.onPick(color, color)}
        />,
      ];
    };
  },
});

/** What `useThemeBases` returns: the ten theme colors with their names. @internal */
export type UseThemeBasesReturn = ComputedRef<readonly { color: string; name: string }[]>;

/** The ten theme colors (the grid's first row), for gradient presets. @internal */
export function useThemeBases(): UseThemeBasesReturn {
  const { t } = useTranslation();
  const first = useEditorState((s) => s.themeColors[0]);
  return computed(() => themeBaseSwatches(t, first.value));
}
