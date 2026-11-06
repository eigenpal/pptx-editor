/** @jsxImportSource vue */
// The Format Background dialog (Format → Slide Background…): the Vue twin of the React
// adapter's `BackgroundDialog`. A color, a gradient, a pattern, or a picture for the current slide or
// every slide, or back to the layout's background. Built on the dialog parts
// (`dialog-parts.ts`): the choices write a draft, OK applies it as one undo step, and Cancel
// discards it. The draft rules live in core (`dialog-fields.ts`, `form-dialog-support.ts`), so
// both adapters apply the same command. The host owns visibility (`open`, `onClose`).

import {
  computed,
  defineComponent,
  h,
  ref,
  shallowRef,
  useId,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
import {
  BACKGROUND_GRADIENT_FALLBACK,
  backgroundDraftColor,
  backgroundDraftCommand,
  backgroundDraftFrom,
  backgroundPresetGradient,
  backgroundStatusKey,
  commitDialogCommand,
  gradientCss,
  patternCss,
  patternDialogStart,
  sameBackground,
  translateDialogErrors,
  validateBackgroundDraft,
  type GradientFill,
  type PatternFill,
  type SlideBackground,
} from '@pptx-editor.dev/core/editor';
import { PatternDialog } from './PatternDialog.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { ToolButton } from './chrome.tsx';
import { ColorSwatches, Swatch, SwatchGroup, useThemeBases } from './ColorSwatches.tsx';
import { GradientDialog } from './ShapeControls.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type UseDialogReturn,
} from './dialog-parts.ts';

/**
 * The draft of the Background dialog. `background` is the fill OK applies: a fill, `null` for
 * the layout's background, or `undefined` to keep the slide's own fill. @public
 */
export interface BackgroundDialogFields {
  /** The fill OK applies: a fill, `null` for the layout's background, `undefined` to keep the slide's own. */
  readonly background: SlideBackground | null | undefined;
  /** Apply the fill to every slide. */
  readonly allSlides: boolean;
}

/** The Background draft and its actions. @public */
export interface UseBackgroundDialogReturn extends UseDialogReturn<BackgroundDialogFields> {
  /** Pick the layout's background in the draft. */
  reset(): void;
}

/**
 * The field names of the Background dialog, for `BackgroundDialog.Field`: one per fill kind
 * (`SlideBackground['kind']`), and Apply to all slides. @public
 */
export type BackgroundDialogFieldName = 'solid' | 'gradient' | 'pattern' | 'picture' | 'allSlides';

const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif'];

const parts = createDialogParts<UseBackgroundDialogReturn, BackgroundDialogFieldName>(
  'BackgroundDialog'
);

const backgroundDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/** Props of `BackgroundDialog`, the twin of React's `BackgroundDialogProps`. @public */
export type BackgroundDialogProps = ExtractPublicPropTypes<typeof backgroundDialogProps>;

const BackgroundDialogOpen = defineComponent({
  name: 'BackgroundDialogOpen',
  inheritAttrs: false,
  props: backgroundDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const file = ref<HTMLInputElement | null>(null);
    const current = useEditorState((s) => s.slideBackground);
    // Why OK cannot apply now, or null when it can.
    const refusal = useEditorState((s) =>
      s.status !== 'ready' ? 'disabled.noEditor' : s.mode !== 'edit' ? 'disabled.viewing' : null
    );
    const initial = backgroundDraftFrom(current.value);
    const custom = shallowRef<GradientFill | null>(null);
    const pattern = shallowRef<PatternFill | null>(null);
    // Tile a chosen picture at its own size; a change re-picks the chosen picture.
    const tile = ref(false);
    // Every pick and picture read counts up, so a picture that finishes reading after a later
    // pick is dropped. `picked` turns on with a real pick (Reset too), so OK applies even a value
    // equal to the opening one; setting the opening value itself back (by identity) turns it off.
    let picks = 0;
    let picked = false;
    const draft: ReturnType<typeof useDialogDraft<BackgroundDialogFields>> = useDialogDraft({
      initial,
      validate: (values) => validateBackgroundDraft(values, initial, picked),
      resetValues: (): BackgroundDialogFields => ({
        background: null,
        allSlides: draft.values.value.allSlides,
      }),
      commit: (values) =>
        commitDialogCommand(editor.value, backgroundDraftCommand(values, initial, picked)),
    });
    const close = (): void => props.onClose();
    const state: UseBackgroundDialogReturn = {
      values: draft.values,
      errors: computed(() => translateDialogErrors(t, draft.errors.value)),
      isEnabled: computed(() => refusal.value === null),
      setValue: (name, value) => {
        // Setting the value it already has is no pick, and going back to the opening value
        // undoes the pick: OK must not write the opening snapshot.
        if (name === 'background' && !Object.is(draft.values.value.background, value)) {
          picks += 1;
          picked = !Object.is(value, initial.background);
        }
        draft.setValue(name, value);
      },
      apply: () => {
        if (draft.apply()) close();
      },
      cancel: close,
      reset: () => {
        picks += 1;
        picked = true;
        draft.reset();
      },
    };
    const render = parts.provideParts(state);
    const pick = (background: SlideBackground): void => state.setValue('background', background);
    const onFile = async (f: File | undefined): Promise<void> => {
      if (!f || !IMAGE_TYPES.includes(f.type)) return;
      const at = ++picks;
      let bytes: Uint8Array;
      try {
        bytes = new Uint8Array(await f.arrayBuffer());
      } catch {
        return; // An unreadable file picks nothing.
      }
      if (at === picks) pick({ kind: 'picture', bytes, mime: f.type, tile: tile.value });
    };
    const bases = useThemeBases();
    const errorId = useId();
    useCloseOnNewDocument(close);
    return () => {
      const values = draft.values.value;
      const errors = state.errors.value;
      const enabled = state.isEnabled.value;
      const why = refusal.value ? t(refusal.value) : undefined;
      const chosen = values.background;
      const error = errors.allSlides ?? errors.background ?? errors.form;
      const presets = bases.value.map((b) => ({
        base: b,
        gradient: backgroundPresetGradient(b.color),
      }));
      // What the draft holds, for the status line a screen reader announces. A swatch pick
      // shows as its pressed swatch instead.
      const statusKey = backgroundStatusKey(
        chosen,
        initial.background,
        presets.map((p) => p.gradient)
      );
      const defaults = (
        <>
          <div data-pptx-part="header" class="pptx-dialog__header">
            <h2 data-pptx-part="title" class="pptx-dialog__title">
              {t('background.title')}
            </h2>
            <ToolButton label={t('background.close')} icon="close" onClick={close} />
          </div>
          <div data-pptx-part="body" class="pptx-dialog__body">
            <section data-pptx-part="field" data-pptx-field="solid" class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('background.color')}</h3>
              <ColorSwatches
                menu={false}
                selected={backgroundDraftColor(values)}
                onPick={(color) => pick({ kind: 'solid', color })}
              />
            </section>
            <section data-pptx-part="field" data-pptx-field="gradient" class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('background.gradient')}</h3>
              {presets.length ? (
                <SwatchGroup label={t('toolbar.gradients')} roving>
                  {{
                    default: () => (
                      <>
                        {presets.map(({ base: b, gradient: g }) => (
                          <Swatch
                            key={b.name}
                            color={b.color}
                            background={gradientCss(g)}
                            name={t('toolbar.gradientPreset', { color: b.name })}
                            menu={false}
                            pressed={sameBackground(chosen, { kind: 'gradient', gradient: g })}
                            onPick={() => pick({ kind: 'gradient', gradient: g })}
                          />
                        ))}
                      </>
                    ),
                  }}
                </SwatchGroup>
              ) : null}
              <div class="pptx-gradient__row">
                <button
                  type="button"
                  class="pptx-button"
                  onClick={() => {
                    custom.value =
                      chosen?.kind === 'gradient'
                        ? chosen.gradient
                        : (current.value.gradient ??
                          backgroundPresetGradient(
                            bases.value[4]?.color ?? BACKGROUND_GRADIENT_FALLBACK
                          ));
                  }}
                >
                  {t('toolbar.gradientCustom')}
                </button>
              </div>
            </section>
            <section data-pptx-part="field" data-pptx-field="pattern" class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('background.pattern')}</h3>
              <div class="pptx-gradient__row">
                {chosen?.kind === 'pattern' ? (
                  <span
                    class="pptx-pattern__chip"
                    aria-hidden="true"
                    style={{ background: patternCss(chosen.pattern) }}
                  />
                ) : null}
                <button
                  type="button"
                  class="pptx-button"
                  onClick={() => {
                    pattern.value =
                      chosen?.kind === 'pattern'
                        ? chosen.pattern
                        : patternDialogStart(current.value.pattern, bases.value[4]?.color);
                  }}
                >
                  {t('background.choosePattern')}
                </button>
              </div>
            </section>
            <section data-pptx-part="field" data-pptx-field="picture" class="pptx-dialog__section">
              <h3 class="pptx-dialog__label">{t('background.image')}</h3>
              <div class="pptx-gradient__row">
                <button type="button" class="pptx-button" onClick={() => file.value?.click()}>
                  {t('background.chooseImage')}
                </button>
                <input
                  ref={file}
                  type="file"
                  accept={IMAGE_TYPES.join(',')}
                  hidden
                  onChange={(e: Event) => {
                    const input = e.target as HTMLInputElement;
                    void onFile(input.files?.[0]);
                    input.value = '';
                  }}
                />
              </div>
              <label class="pptx-radio">
                <input
                  type="checkbox"
                  checked={tile.value}
                  onChange={(e: Event) => {
                    const on = (e.target as HTMLInputElement).checked;
                    tile.value = on;
                    if (chosen?.kind === 'picture') pick({ ...chosen, tile: on });
                  }}
                />
                {t('background.tile')}
              </label>
            </section>
            <p role="status" class="pptx-dialog__muted pptx-dialog__status">
              {statusKey ? t(statusKey) : ''}
            </p>
            <label data-pptx-part="field" data-pptx-field="allSlides" class="pptx-radio">
              <input
                type="checkbox"
                aria-invalid={errors.allSlides ? true : undefined}
                aria-describedby={errors.allSlides ? errorId : undefined}
                checked={values.allSlides}
                onChange={(e: Event) =>
                  draft.setValue('allSlides', (e.target as HTMLInputElement).checked)
                }
              />
              {t('background.allSlides')}
            </label>
          </div>
          <div data-pptx-part="footer" class="pptx-dialog__footer">
            <span id={errorId} data-pptx-part="error" role="alert" class="pptx-dialog__error">
              {error ?? null}
            </span>
            <button
              type="button"
              data-pptx-part="reset"
              class="pptx-button pptx-dialog__reset"
              disabled={!enabled}
              title={why}
              onClick={state.reset}
            >
              {t('background.reset')}
            </button>
            <button
              type="button"
              data-pptx-part="cancel"
              class="pptx-button"
              onClick={state.cancel}
            >
              {t('dialog.cancel')}
            </button>
            <button
              type="button"
              data-pptx-part="apply"
              class="pptx-button pptx-button--primary"
              disabled={!enabled}
              title={why}
              onClick={state.apply}
            >
              {t('dialog.ok')}
            </button>
          </div>
        </>
      ) as VNode;
      return (
        <DialogFrame
          {...attrs}
          class={[attrs.class, props.className]}
          style={[attrs.style as CSSProperties, props.style]}
          kind="background"
          label={t('background.title')}
          onClose={close}
          onKeydown={(event: KeyboardEvent) => applyOnEnter(event, enabled, state.apply)}
        >
          {{
            default: () => [
              ...render(defaults, slots.default?.(), props.preset),
              <PatternDialog
                open={pattern.value !== null}
                initial={pattern.value}
                onClose={() => {
                  pattern.value = null;
                }}
                onApply={(p: PatternFill) => {
                  pick({ kind: 'pattern', pattern: p });
                  pattern.value = null;
                }}
              />,
              <GradientDialog
                open={custom.value !== null}
                initial={custom.value}
                onClose={() => {
                  custom.value = null;
                }}
                onApply={(g: GradientFill) => {
                  pick({ kind: 'gradient', gradient: g });
                  custom.value = null;
                }}
              />,
            ],
          }}
        </DialogFrame>
      );
    };
  },
});

/**
 * Change the slide background, with replaceable parts (`Header`, `Title`, `Body`, `Footer`,
 * `Field`, `Error`, `Reset`, `Apply`, `Cancel`), as React's `BackgroundDialog`. Renders nothing
 * while `open` is false. @public
 */
export const BackgroundDialog = Object.assign(
  defineComponent({
    name: 'BackgroundDialog',
    inheritAttrs: false,
    props: backgroundDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(BackgroundDialogOpen, { ...attrs, ...props }, slots) : null);
    },
  }),
  {
    Header: parts.Header,
    Title: parts.Title,
    Body: parts.Body,
    Footer: parts.Footer,
    Field: parts.Field,
    Error: parts.Error,
    Reset: parts.Reset,
    Apply: parts.Apply,
    Cancel: parts.Cancel,
  }
);

/**
 * Read the draft of the enclosing Background dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('background').open()` does. @public
 */
export function useBackgroundDialog(): UseBackgroundDialogReturn {
  return parts.useState();
}
