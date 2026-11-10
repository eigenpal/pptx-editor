/** @jsxImportSource vue */
// Shape controls: the insert-shape gallery (previews drawn by the engine's own geometry), the
// fill menu (theme colors, gradients, a pattern, and a picture), the custom gradient dialog, and
// the line menus. The Vue
// twin of the React adapter's `ShapeControls.tsx`: the same markup, classes, ARIA, and keys, on
// the same core rules (`insert-controls.ts`, `shape-gallery.ts`, `palette.ts`).

import {
  computed,
  defineComponent,
  inject,
  onMounted,
  ref,
  shallowRef,
  useId,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  formatSelectionHasLineEnds,
  dialogInitialFocus,
  isDialogBackdropPoint,
  colorMenuSelection,
  SHAPE_GALLERY,
  CROP_SHAPE_GALLERY,
  LINE_GLYPH_STROKE,
  defaultFillChoice,
  gradientAddStop,
  gradientAngleFrom,
  gradientCss,
  gradientDialogStart,
  gradientHandleKey,
  gradientPosAt,
  gradientRemoveStop,
  gradientSetStop,
  lineGlyphParts,
  lineGlyphTransform,
  lineSlotEnd,
  lineSlotShown,
  presetGradient,
  shapePreview,
  toolbarDropdownRow,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
  type ColorChoice,
  type GradientFill,
  type PatternFill,
  pickPictureFill,
} from '@pptx-editor.dev/core/editor';
import { PatternDialog } from './PatternDialog.tsx';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import {
  Icon,
  MenuItem,
  Popover,
  ToolButton,
  focusText,
  keepFocus,
  useAriaDisabled,
} from './chrome.tsx';
import {
  chromeActionHostKey,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';
import { ColorSwatches, Swatch, SwatchGroup, useThemeBases } from './ColorSwatches.tsx';

/** The props every slot renderer takes: the slot id and its registry entry. */
const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

const shapeIconProps = { preset: { type: String, required: true as const } } as const;

/** The props of `ShapeIcon`, the twin of React's `ShapeIconProps`. @public */
export type ShapeIconProps = ExtractPublicPropTypes<typeof shapeIconProps>;

/** A preset's preview, drawn from its spec geometry. @public */
export const ShapeIcon = defineComponent({
  name: 'ShapeIcon',
  props: shapeIconProps,
  setup(props) {
    return () => {
      const preview = shapePreview(props.preset);
      if (!preview) return null;
      return (
        <svg viewBox={preview.viewBox} class="pptx-shape-icon" aria-hidden="true">
          {preview.paths.map((p, i) => (
            <path
              key={i}
              d={p.d}
              class={p.fill ? 'pptx-shape-icon__fill' : undefined}
              fill={p.fill ? undefined : 'none'}
              stroke={p.stroke ? 'currentColor' : 'none'}
              vector-effect="non-scaling-stroke"
            />
          ))}
        </svg>
      );
    };
  },
});

const shapeGridProps = {
  presets: { type: Array as PropType<readonly string[]>, required: true as const },
  /** Called with the picked preset, as React's `onPick`. */
  onPick: { type: Function as PropType<(preset: string) => void>, required: true as const },
} as const;

/** The props of `ShapeGrid`, the twin of React's `ShapeGridProps`. @public */
export type ShapeGridProps = ExtractPublicPropTypes<typeof shapeGridProps>;

/**
 * One gallery category's presets as a grid of previews drawn from their spec geometry. Each
 * preview is a `menuitem`; the arrow keys move by grid rows (`data-columns`). A pick calls
 * `onPick` with the preset. @public
 */
export const ShapeGrid = defineComponent({
  name: 'ShapeGrid',
  props: shapeGridProps,
  setup(props) {
    const { t } = useTranslation();
    return () => (
      <div class="pptx-shape-gallery__grid" data-columns={8}>
        {props.presets.map((preset) => {
          const name = translateKeyOrText(t, `shapes.${preset}`);
          return (
            <button
              key={preset}
              type="button"
              role="menuitem"
              tabindex={-1}
              class="pptx-shape-gallery__item"
              aria-label={name}
              data-pptx-tooltip={name}
              onClick={() => props.onPick(preset)}
            >
              <ShapeIcon preset={preset} />
            </button>
          );
        })}
      </div>
    );
  },
});

/** The props of `ShapeGalleryControl`, the twin of React's `ShapeGalleryControlProps`. @public */
export type ShapeGalleryControlProps = ExtractPublicPropTypes<typeof slotPartProps>;

/** The insert-shape gallery: every category, every preset as its own preview. @public */
export const ShapeGalleryControl = defineComponent({
  name: 'ShapeGalleryControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const { t } = useTranslation();
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
        <Icon paths={props.control.paths} />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>,
      off.note(),
      open.value ? (
        <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
          {{
            default: () => (
              <>
                <div class="pptx-shape-gallery">
                  {/* A picture is cropped to shapes with an inside: lines are left out (core). */}
                  {(props.slot === 'image.cropToShape' ? CROP_SHAPE_GALLERY : SHAPE_GALLERY).map(
                    (category) => (
                      <section
                        key={category.id}
                        class="pptx-shape-gallery__section"
                        role="group"
                        aria-label={t(category.labelKey as TranslationKey)}
                      >
                        <h3 class="pptx-shape-gallery__title" aria-hidden="true">
                          {t(category.labelKey as TranslationKey)}
                        </h3>
                        <ShapeGrid
                          presets={category.presets}
                          onPick={(preset: string) => {
                            cmd.execute(preset);
                            open.value = false;
                          }}
                        />
                      </section>
                    )
                  )}
                </div>
              </>
            ),
          }}
        </Popover>
      ) : null,
    ];
  },
});

/** The props of `FillControl`, the twin of React's `FillControlProps`. @public */
export type FillControlProps = ExtractPublicPropTypes<typeof slotPartProps>;

/**
 * Fill color: one menu with the theme and standard colors, Transparent, gradient presets from
 * the theme colors, a custom gradient dialog, a pattern dialog, and a picture. @public
 */
export const FillControl = defineComponent({
  name: 'FillControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const current = useEditorState((s) => s.geometry?.gradient);
    // The solid fill of the selection, marked in the menu.
    const selected = useEditorState((s) => colorMenuSelection(props.slot, s));
    const bases = useThemeBases();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    const picked = shallowRef<{ choice: ColorChoice; color: string } | null>(null);
    // Before a pick, the button applies the theme's Accent 1, the default fill of a new shape.
    const last = computed<{ choice: ColorChoice; color: string }>(
      () => picked.value ?? defaultFillChoice(bases.value[4]?.color)
    );
    const dialog = shallowRef<GradientFill | null>(null);
    const applyGradient = (g: GradientFill): void => {
      editor.value?.exec({ type: 'setShapeGradient', gradient: g });
    };
    const currentPattern = useEditorState((s) => s.geometry?.pattern);
    const pattern = shallowRef<PatternFill | null | false>(false);
    const done = (): void => {
      focusText(editor.value, anchor.value);
      open.value = false;
    };
    return () => (
      <span class="pptx-toolbar__group">
        <ToolButton
          slot={props.slot}
          label={label.value}
          paths={props.control.paths}
          disabled={!cmd.isEnabled.value}
          disabledReason={reason.value}
          onClick={() => cmd.execute(last.value.choice)}
        >
          {{
            default: () => (
              <>
                <span class="pptx-toolbar__swatch" style={{ background: last.value.color }} />
              </>
            ),
          }}
        </ToolButton>
        <button
          ref={anchor}
          type="button"
          class="pptx-toolbar__button pptx-toolbar__step"
          aria-label={label.value}
          aria-haspopup="menu"
          aria-expanded={open.value}
          {...tip.value}
          {...off.attrs.value}
          onMousedown={keepFocus}
          onClick={() => (open.value = !open.value)}
        >
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>
        {off.note()}
        {open.value ? (
          <Popover anchor={anchor.value} label={label.value} onClose={() => (open.value = false)}>
            {{
              default: () => (
                <>
                  <ColorSwatches
                    selected={selected.value}
                    onPick={(choice: ColorChoice, color: string) => {
                      picked.value = { choice, color };
                      cmd.execute(choice);
                      done();
                    }}
                  />
                  <MenuItem onSelect={() => (cmd.execute(null), done())}>
                    {{ default: () => t('toolbar.transparent') }}
                  </MenuItem>
                  <div class="pptx-menu__sep" />
                  {bases.value.length ? (
                    <SwatchGroup label={t('toolbar.gradients')}>
                      {{
                        default: () => (
                          <>
                            {bases.value.map((b) => {
                              const g = presetGradient(b.color, 0.35, 0.35);
                              return (
                                <Swatch
                                  key={b.name}
                                  color={b.color}
                                  background={gradientCss(g)}
                                  name={t('toolbar.gradientPreset', { color: b.name })}
                                  menu
                                  onPick={() => (applyGradient(g), done())}
                                />
                              );
                            })}
                          </>
                        ),
                      }}
                    </SwatchGroup>
                  ) : null}
                  <MenuItem
                    onSelect={() => {
                      dialog.value = current.value ?? presetGradient(last.value.color, 0.35, 0.35);
                      open.value = false;
                    }}
                  >
                    {{
                      default: () => (
                        <>
                          <Icon name="add" />
                          {t('toolbar.gradientCustom')}
                        </>
                      ),
                    }}
                  </MenuItem>
                  <div class="pptx-menu__sep" />
                  <MenuItem
                    onSelect={() => {
                      pattern.value = currentPattern.value ?? null;
                      open.value = false;
                    }}
                  >
                    {{ default: () => t('toolbar.pattern') }}
                  </MenuItem>
                  <MenuItem
                    onSelect={() => {
                      if (editor.value) pickPictureFill(editor.value, anchor.value);
                      open.value = false;
                    }}
                  >
                    {{ default: () => t('toolbar.pictureFill') }}
                  </MenuItem>
                </>
              ),
            }}
          </Popover>
        ) : null}
        <PatternDialog
          open={pattern.value !== false}
          initial={pattern.value || null}
          base={bases.value[4]?.color}
          onClose={() => {
            pattern.value = false;
            // Focus goes back where the fill menu was opened, not to the page.
            focusText(editor.value, anchor.value);
          }}
          onApply={(p: PatternFill) => {
            editor.value?.exec({ type: 'setShapePattern', pattern: p });
            pattern.value = false;
            focusText(editor.value, anchor.value);
          }}
        />
        <GradientDialog
          open={dialog.value !== null}
          initial={dialog.value}
          onClose={() => {
            dialog.value = null;
            focusText(editor.value, anchor.value);
          }}
          onApply={(g: GradientFill) => {
            applyGradient(g);
            dialog.value = null;
            focusText(editor.value, anchor.value);
          }}
        />
      </span>
    );
  },
});

const gradientDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** The gradient the dialog starts from. */
  initial: { type: Object as PropType<GradientFill | null>, default: null },
  /** Called on Cancel, Escape, and a click on the backdrop. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** Called by OK with the gradient. */
  onApply: {
    type: Function as PropType<(gradient: GradientFill) => void>,
    required: true as const,
  },
} as const;

/** Props of `GradientDialog`, the twin of React's `GradientDialogProps`. @public */
export type GradientDialogProps = ExtractPublicPropTypes<typeof gradientDialogProps>;

/**
 * The custom gradient dialog: type, angle, stops, rotation, and a preview. Calls `onClose`,
 * and `onApply` with the gradient. @public
 */
export const GradientDialog = defineComponent({
  name: 'GradientDialog',
  props: gradientDialogProps,
  setup(props) {
    const { t } = useTranslation();
    const dialogEl = ref<HTMLDialogElement | null>(null);
    const titleId = useId();
    let pressedOutside = false;
    const bar = ref<HTMLDivElement | null>(null);
    const g = shallowRef<GradientFill>(gradientDialogStart(props.initial));
    const selected = ref(0);
    let dragging: number | null = null;
    const onClose = (): void => props.onClose();
    // Open and close the modal with `open`, after the element is in the document.
    const sync = (): void => {
      const el = dialogEl.value;
      if (!el) return;
      if (props.open) {
        g.value = gradientDialogStart(props.initial);
        selected.value = 0;
        if (!el.open) el.showModal();
        // The first field takes focus, as in every dialog, not the Close button.
        dialogInitialFocus(el)?.focus({ preventScroll: true });
      } else if (el.open) el.close();
    };
    onMounted(sync);
    watch(() => [props.open, props.initial] as const, sync, { flush: 'post' });
    const setStop = (i: number, patch: Partial<{ pos: number; color: string }>): void => {
      g.value = gradientSetStop(g.value, i, patch);
    };
    const add = (): void => {
      const next = gradientAddStop(g.value, selected.value);
      g.value = next.gradient;
      selected.value = next.selected;
    };
    const remove = (): void => {
      const next = gradientRemoveStop(g.value, selected.value);
      g.value = next.gradient;
      selected.value = next.selected;
    };
    const body = () => {
      const cur = g.value;
      const stop = cur.stops[Math.min(selected.value, cur.stops.length - 1)]!;
      return [
        <div class="pptx-dialog__header">
          <h2 id={titleId} class="pptx-dialog__title">
            {t('gradient.title')}
          </h2>
          <ToolButton label={t('gradient.close')} icon="close" onClick={onClose} />
        </div>,
        <div class="pptx-dialog__body">
          <div class="pptx-gradient__top">
            <div class="pptx-gradient__fields">
              <div class="pptx-gradient__row">
                <label class="pptx-field">
                  {t('gradient.type')}
                  <select
                    value={cur.kind}
                    onChange={(e: Event) => {
                      const v = (e.target as HTMLSelectElement).value;
                      g.value = { ...g.value, kind: v === 'radial' ? 'radial' : 'linear' };
                    }}
                  >
                    <option value="linear">{t('gradient.linear')}</option>
                    <option value="radial">{t('gradient.radial')}</option>
                  </select>
                </label>
                <label class="pptx-field">
                  {t('gradient.angle')}
                  <span class="pptx-field__control">
                    <input
                      type="number"
                      min={0}
                      max={359}
                      step={1}
                      value={Math.round(cur.angle)}
                      disabled={cur.kind === 'radial'}
                      onInput={(e: Event) => {
                        const angle = gradientAngleFrom((e.target as HTMLInputElement).value);
                        if (angle !== null) g.value = { ...g.value, angle };
                      }}
                    />
                    <span class="pptx-field__unit">°</span>
                  </span>
                </label>
              </div>
              <div class="pptx-dialog__section">
                <h3 class="pptx-dialog__label">{t('gradient.stops')}</h3>
                <div class="pptx-gradient__row">
                  <button type="button" class="pptx-button" onClick={add}>
                    {t('gradient.add')}
                  </button>
                  <button
                    type="button"
                    class="pptx-button"
                    onClick={remove}
                    disabled={cur.stops.length <= 2}
                  >
                    {t('gradient.remove')}
                  </button>
                  <input
                    type="color"
                    class="pptx-gradient__color"
                    aria-label={t('gradient.stopColor')}
                    value={stop.color.toLowerCase()}
                    onInput={(e: Event) =>
                      setStop(selected.value, {
                        color: (e.target as HTMLInputElement).value.toUpperCase(),
                      })
                    }
                  />
                </div>
              </div>
              <label class="pptx-radio">
                <input
                  type="checkbox"
                  checked={cur.rotWithShape}
                  onChange={(e: Event) => {
                    g.value = { ...g.value, rotWithShape: (e.target as HTMLInputElement).checked };
                  }}
                />
                {t('gradient.rotate')}
              </label>
            </div>
            <div class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('gradient.preview')}</h3>
              <div class="pptx-gradient__preview" style={{ background: gradientCss(cur) }} />
            </div>
          </div>
          <div
            ref={bar}
            class="pptx-gradient__bar"
            style={{ background: gradientCss({ ...cur, kind: 'linear', angle: 0 }) }}
          >
            {cur.stops.map((s, i) => (
              <button
                key={i}
                type="button"
                class="pptx-gradient__handle"
                aria-label={t('gradient.stopHandle', { percent: Math.round(s.pos * 100) })}
                aria-pressed={i === selected.value}
                style={{ left: `${s.pos * 100}%`, background: s.color }}
                onPointerdown={(e: PointerEvent) => {
                  selected.value = i;
                  dragging = i;
                  (e.currentTarget as Element).setPointerCapture(e.pointerId);
                }}
                onPointermove={(e: PointerEvent) => {
                  if (dragging !== null)
                    setStop(dragging, {
                      pos: gradientPosAt(bar.value?.getBoundingClientRect(), e.clientX),
                    });
                }}
                onPointerup={() => (dragging = null)}
                onKeydown={(e: KeyboardEvent) => {
                  const pos = gradientHandleKey(e.key, e.shiftKey, s.pos);
                  if (pos !== null) setStop(i, { pos });
                }}
              />
            ))}
          </div>
        </div>,
        <div class="pptx-dialog__footer">
          <button type="button" class="pptx-button" onClick={onClose}>
            {t('gradient.cancel')}
          </button>
          <button
            type="button"
            class="pptx-button pptx-button--primary"
            onClick={() => props.onApply(g.value)}
          >
            {t('gradient.ok')}
          </button>
        </div>,
      ];
    };
    return () => (
      <dialog
        ref={dialogEl}
        class="pptx-dialog pptx-gradient"
        data-pptx-dialog="gradient"
        aria-labelledby={titleId}
        onCancel={(e: Event) => {
          e.preventDefault();
          onClose();
        }}
        onMousedown={(e: MouseEvent) => {
          pressedOutside = isDialogBackdropPoint(e.currentTarget as HTMLElement, e);
        }}
        onClick={(e: MouseEvent) => {
          // A press and a release on the backdrop close it; a drag out of a field does not.
          if (pressedOutside && isDialogBackdropPoint(e.currentTarget as HTMLElement, e)) onClose();
          pressedOutside = false;
        }}
      >
        {props.open ? body() : null}
      </dialog>
    );
  },
});

const lineGlyphProps = {
  value: { type: String, required: true as const },
  start: { type: Boolean, default: undefined },
} as const;

/** The props of `LineGlyph`, the twin of React's `LineGlyphProps`. @public */
export type LineGlyphProps = ExtractPublicPropTypes<typeof lineGlyphProps>;

/** Glyphs for line kinds, line ends, and preset dashes, drawn in a 48×24 box. @public */
export const LineGlyph = defineComponent({
  name: 'LineGlyph',
  props: lineGlyphProps,
  setup(props) {
    const stroke = {
      fill: LINE_GLYPH_STROKE.fill,
      stroke: LINE_GLYPH_STROKE.stroke,
      'stroke-width': LINE_GLYPH_STROKE.strokeWidth,
      'stroke-linecap': LINE_GLYPH_STROKE.strokeLinecap,
    };
    const solid = { fill: 'currentColor', stroke: 'none' };
    return () => (
      <svg class="pptx-line-glyph" viewBox="0 0 48 24" aria-hidden="true">
        <g transform={lineGlyphTransform(props.start)}>
          {lineGlyphParts(props.value, props.start).map((p, i) =>
            p.kind === 'dot' ? (
              <circle key={i} cx={p.cx} cy={p.cy} r={p.r} {...solid} />
            ) : p.kind === 'solid' ? (
              <path key={i} d={p.d} {...solid} />
            ) : (
              <path
                key={i}
                d={p.d}
                {...stroke}
                {...(p.join ? { 'stroke-linejoin': 'round' } : {})}
                {...(p.dash ? { 'stroke-dasharray': p.dash } : {})}
              />
            )
          )}
        </g>
      </svg>
    );
  },
});

/** The props of `LineMenuControl`, the twin of React's `LineMenuControlProps`. @public */
export type LineMenuControlProps = ExtractPublicPropTypes<typeof slotPartProps>;

/**
 * The line tool, the line start and end menus, and the dash menus, with each option drawn.
 * @public
 */
export const LineMenuControl = defineComponent({
  name: 'LineMenuControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const editor = usePptxEditor();
    const host = inject(chromeActionHostKey, null);
    const { t } = useTranslation();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    const isLine = useEditorState((s) => formatSelectionHasLineEnds(s));
    const trigger = useTriggerAttributes(
      props.slot,
      'pptx-toolbar__button pptx-toolbar__button--wide',
      () => ({ disabled: !cmd.isEnabled.value })
    );
    return () => {
      // Line ends draw on the side they belong to; the line tool and dash menus show their
      // options.
      const start = lineSlotEnd(props.slot);
      const shown = lineSlotShown(props.slot, cmd.value.value);
      // Line ends are contextual: they show only while a line is selected.
      if (start !== undefined && !isLine.value) return null;
      const end = start !== undefined ? { start } : {};
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
          <LineGlyph value={shown} {...end} />
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>,
        off.note(),
        open.value ? (
          <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
            {{
              default: () => (
                <>
                  {props.control.options?.map((o) => {
                    // Line ends and dashes are a choice with a current value; the line tool
                    // inserts a kind. Core's rule says which rows are checked or off, with the
                    // reason.
                    const row = toolbarDropdownRow(
                      editor.value,
                      host,
                      props.slot,
                      props.control,
                      o,
                      cmd.value.value
                    );
                    return (
                      <MenuItem
                        key={o.value}
                        {...(row.checked !== undefined ? { checked: row.checked } : {})}
                        {...(row.refused
                          ? {
                              disabled: true,
                              ...(row.reason ? { reason: translateKeyOrText(t, row.reason) } : {}),
                            }
                          : {})}
                        onSelect={() => {
                          cmd.execute(o.value);
                          focusText(editor.value, anchor.value);
                          open.value = false;
                        }}
                      >
                        {{
                          default: () => (
                            <>
                              <LineGlyph value={o.value} {...end} />
                              {t(o.labelKey as TranslationKey)}
                            </>
                          ),
                        }}
                      </MenuItem>
                    );
                  })}
                </>
              ),
            }}
          </Popover>
        ) : null,
      ];
    };
  },
});
