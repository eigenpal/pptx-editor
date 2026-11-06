/** @jsxImportSource vue */
// The Shape Format controls whose look is more than an icon: the Shape Styles gallery, which
// draws each theme style with the slide's theme colors, and the Shadow, Reflection, Glow, and
// Soft Edges menus, which name their current choice. The Vue twin of the React adapter's
// `ShapeFormatControls.tsx`, on the same core rules (`chrome-shapes.ts`, `insert-controls.ts`).
// The fill control lives with the other shape controls (`ShapeControls.tsx`) and is re-exported
// here.

import { defineComponent, ref, type ExtractPublicPropTypes, type PropType } from 'vue';
import {
  GLOW_ACCENTS,
  SHAPE_STYLE_GALLERY,
  chromeOptionChecked,
  effectSeparatorBefore,
  glowGridRows,
  shapeStyleChoice,
  shapeStylePreview,
  type ChromeControl,
  type ChromeSlotId,
  type ShapeStyleChoice,
  type ShapeStyleValue,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { MenuRow, MenuSeparator, MenuSubmenu } from './menu-row.tsx';
import { useEditorState, useOfferedControl, useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

export { FillControl } from './ShapeControls.tsx';

const cellKey = (c: ShapeStyleChoice): string => `${c.row}.${c.color}`;

/** The cell the trigger shows while the selection has no gallery style: Colored Fill, Accent 1. */
const DEFAULT_CELL: ShapeStyleChoice = { row: 'coloredFill', color: 'accent1' };

const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** One gallery cell: a small rounded box in the style's fill, outline, and text colors. */
const StyleSwatch = defineComponent({
  name: 'PptxStyleSwatch',
  props: {
    cell: { type: Object as PropType<ShapeStyleChoice>, required: true as const },
    className: { type: String, default: 'pptx-style-gallery__swatch' },
  },
  setup(props) {
    const { t } = useTranslation();
    const themeColors = useEditorState((s) => s.themeColors);
    return () => {
      const p = shapeStylePreview(props.cell, themeColors.value);
      return (
        <svg class={props.className} viewBox="0 0 40 28" aria-hidden="true">
          {p.effect > 0 ? (
            <rect
              x="4"
              y="5"
              width="34"
              height="22"
              rx="3"
              class="pptx-style-gallery__shadow"
              opacity={0.12 * p.effect}
            />
          ) : null}
          <rect
            x="2"
            y="2"
            width="34"
            height="22"
            rx="3"
            fill={p.fill ?? 'none'}
            stroke={p.line ?? 'none'}
            stroke-width={p.line ? 1.5 : 0}
          />
          <text x="19" y="17" text-anchor="middle" font-size="9" fill={p.text}>
            {t('toolbar.shapeStyleSample')}
          </text>
        </svg>
      );
    };
  },
});

const shapeStyleGalleryProps = {
  /**
   * The cell to check: a cell, or the slot value (`row.color`, as
   * `useEditorValueCommand('shape.styles').value` gives it), or null for none.
   */
  value: {
    type: [Object, String] as PropType<ShapeStyleChoice | ShapeStyleValue | null>,
    default: undefined,
  },
  /** Called with the picked cell, as React's `onPick`. */
  onPick: { type: Function as PropType<(cell: ShapeStyleChoice) => void>, required: true as const },
  /** A class name added to the gallery, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `ShapeStyleGallery`, the twin of React's `ShapeStyleGalleryProps`. @public */
export type ShapeStyleGalleryProps = ExtractPublicPropTypes<typeof shapeStyleGalleryProps>;

/**
 * The Quick Styles gallery: six theme style rows of seven colors, drawn with the current
 * slide's theme. Put it in a popover or a menu. A pick calls `onPick` with the cell;
 * `className` (or a host `class`) goes on the gallery. @public
 */
export const ShapeStyleGallery = defineComponent({
  name: 'ShapeStyleGallery',
  props: shapeStyleGalleryProps,
  setup(props) {
    const { t } = useTranslation();
    return () => {
      const picked = shapeStyleChoice(props.value);
      const checked = picked ? cellKey(picked) : undefined;
      return (
        <div
          class={['pptx-style-gallery', props.className]}
          role="group"
          aria-label={t('toolbar.shapeStyles')}
          // A grid for the arrow keys: Up and Down move by a row of colors.
          data-columns={SHAPE_STYLE_GALLERY[0]?.cells.length}
        >
          {SHAPE_STYLE_GALLERY.map(({ row, cells }) => (
            <div key={row} class="pptx-style-gallery__row" role="presentation">
              {cells.map((cell) => {
                const name = t('toolbar.shapeStyleCell', {
                  row: t(`toolbar.shapeStyleRows.${cell.row}` as TranslationKey),
                  color: t(`toolbar.shapeStyleColors.${cell.color}` as TranslationKey),
                });
                return (
                  <button
                    key={cellKey(cell)}
                    type="button"
                    role="menuitemradio"
                    aria-checked={checked === cellKey(cell)}
                    tabindex={-1}
                    class="pptx-menu__item pptx-style-gallery__cell"
                    aria-label={name}
                    data-pptx-tooltip={name}
                    data-style={cellKey(cell)}
                    onClick={() => props.onPick(cell)}
                  >
                    <StyleSwatch cell={cell} />
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      );
    };
  },
});

/** The toolbar control of `shape.styles`: a trigger and the gallery in a popover. @internal */
export const ShapeStyleGalleryControl = defineComponent({
  name: 'PptxShapeStyleGalleryControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    // a disabled picker stays in the arrow order and says why.
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    const trigger = useTriggerAttributes(
      props.slot,
      'pptx-toolbar__button pptx-toolbar__button--wide',
      () => ({ disabled: !cmd.isEnabled.value })
    );
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
        <StyleSwatch
          cell={shapeStyleChoice(cmd.value.value) ?? DEFAULT_CELL}
          className="pptx-style-gallery__trigger"
        />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>,
      off.note(),
      open.value ? (
        <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
          {{
            default: () => (
              <>
                <ShapeStyleGallery
                  value={shapeStyleChoice(cmd.value.value)}
                  onPick={(cell: ShapeStyleChoice) => {
                    cmd.execute(cell);
                    open.value = false;
                  }}
                />
              </>
            ),
          }}
        </Popover>
      ) : null,
    ];
  },
});

/** The Shape Effects trigger: a rounded box with an offset shadow, drawn, not an icon path. */
function effectsGlyph() {
  return (
    <svg class="pptx-icon pptx-effects-glyph" viewBox="0 0 24 24" aria-hidden="true">
      <rect x="7" y="7" width="13" height="13" rx="2.5" class="pptx-effects-glyph__shadow" />
      <rect x="4" y="4" width="13" height="13" rx="2.5" class="pptx-effects-glyph__box" />
    </svg>
  );
}

type EffectOption = NonNullable<ChromeControl['options']>[number];

/**
 * Glow Variations: a grid with a row per size and a column per accent, each cell a box with a
 * glow in that accent. Up and Down move by a row (`data-columns`).
 */
const GlowGrid = defineComponent({
  name: 'PptxGlowGrid',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    options: { type: Array as PropType<readonly EffectOption[]>, required: true as const },
    value: { type: [String, Number] as PropType<string | number>, default: undefined },
  },
  emits: { pick: (_value: string) => true },
  setup(props, { emit }) {
    const { t } = useTranslation();
    const themeColors = useEditorState((s) => s.themeColors);
    const swatch = (accent: string): string =>
      themeColors.value[0]?.find((s) => s.slot === accent)?.color ?? '#808080';
    return () => (
      <div
        class="pptx-style-gallery pptx-glow-grid"
        role="group"
        aria-label={t('toolbar.glowVariationsLabel')}
        data-columns={GLOW_ACCENTS.length}
      >
        {glowGridRows(props.options).map(({ size, cells }) => (
          <div key={size} class="pptx-style-gallery__row" role="presentation">
            {cells.map(({ option: o, accent }) => {
              const name = t(o.labelKey as TranslationKey, o.labelParams);
              return (
                <button
                  key={o.value}
                  type="button"
                  role="menuitemradio"
                  aria-checked={chromeOptionChecked(props.value, o.value)}
                  tabindex={-1}
                  class="pptx-menu__item pptx-style-gallery__cell"
                  aria-label={name}
                  data-pptx-tooltip={name}
                  data-slot={props.slot}
                  data-value={o.value}
                  onClick={() => emit('pick', o.value)}
                >
                  <svg class="pptx-style-gallery__swatch" viewBox="0 0 40 28" aria-hidden="true">
                    <rect
                      x="9"
                      y="6"
                      width="22"
                      height="16"
                      rx="2"
                      fill="none"
                      stroke={swatch(accent)}
                      stroke-opacity="0.4"
                      // The glow width follows the size: 18 pt draws twice as wide as 8 pt.
                      stroke-width={Number(size) / 2}
                    />
                    <rect x="9" y="6" width="22" height="16" rx="2" class="pptx-glow-grid__box" />
                  </svg>
                </button>
              );
            })}
          </div>
        ))}
      </div>
    );
  },
});

/**
 * The Shape Effects menu (`shape.effects`): one submenu per effect (Shadow, Reflection, Glow,
 * Soft Edges, the option groups of the registry), with each effect's current choice checked.
 * @internal
 */
export const EffectMenuControl = defineComponent({
  name: 'PptxEffectMenuControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    // a disabled picker stays in the arrow order and says why.
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
    // The choices of this build and selection: they follow each change while the menu is open.
    const shown = useOfferedControl(
      () => props.slot,
      () => props.control
    );
    const offered = (): ChromeControl => shown.value ?? props.control;
    const row = (o: EffectOption) => (
      <MenuRow
        key={o.value}
        slot={props.slot}
        label={t(o.labelKey as TranslationKey, o.labelParams)}
        checked={chromeOptionChecked(cmd.value.value, o.value)}
        radio
        onSelect={() => pick(o.value)}
      />
    );
    const groupRows = (group: string, options: readonly EffectOption[]) => {
      const [first, ...rest] = options;
      if (group === 'glow' && first)
        return [
          row(first),
          // Without the drawing module, No Glow can be the only row.
          ...(rest.length
            ? [<GlowGrid slot={props.slot} options={rest} value={cmd.value.value} onPick={pick} />]
            : []),
        ];
      // Shadows: a line between the outer and the inner presets (the Outer and Inner sections).
      return options.flatMap((o, i) =>
        effectSeparatorBefore(group, options, i)
          ? [<MenuSeparator key={`sep-${o.value}`} />, row(o)]
          : [row(o)]
      );
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
        {effectsGlyph()}
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>,
      off.note(),
      open.value ? (
        <Popover anchor={anchor.value} onClose={close}>
          {{
            // The effects this build paints: without the drawing module, only the outer shadows.
            default: () => (
              <>
                {offered().optionGroups?.map((group) => (
                  <MenuSubmenu key={group.id} label={t(group.labelKey as TranslationKey)}>
                    {{
                      default: () =>
                        groupRows(
                          group.id,
                          offered().options?.filter((o) => o.group === group.id) ?? []
                        ),
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
