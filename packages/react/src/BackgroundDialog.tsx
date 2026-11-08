/** @jsxImportSource react */
// The Format Background dialog (Format → Slide Background…): a color, a gradient, a pattern, or a picture
// for the current slide or every slide, or back to the layout's background. Built on the dialog
// parts (`dialog-parts.tsx`): the choices write a draft, OK applies it as one undo step, and
// Cancel discards it. The draft rules live in core
// (`dialog-fields.ts`), so the Vue twin applies the same command.

import { useId, useMemo, useRef, useState, type ReactNode } from 'react';
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
  type DialogDraft,
  type GradientFill,
  type PatternFill,
  type SlideBackground,
} from '@pptx-editor.dev/core/editor';
import { PatternDialog } from './PatternDialog.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { ToolButton } from './chrome.tsx';
import { ColorSwatches, Swatch, SwatchGroup, useThemeBases } from './ColorSwatches.tsx';
import { GradientDialog } from './ShapeControls.tsx';
import {
  DialogFrame,
  applyOnEnter,
  createDialogParts,
  useCloseOnNewDocument,
  useDialogDraft,
  type DialogCustomizationProps,
  type UseDialogReturn,
} from './dialog-parts.tsx';

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

/** Props for `PptxEditor.BackgroundDialog`. @public */
export interface BackgroundDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
}

const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif'];

const parts = createDialogParts<BackgroundDialogFieldName, UseBackgroundDialogReturn>(
  'BackgroundDialog'
);

function BackgroundDialogRoot(props: BackgroundDialogProps): ReactNode {
  return props.open ? <BackgroundDialogOpen {...props} /> : null;
}

function BackgroundDialogOpen({
  onClose,
  className,
  style,
  preset = true,
  children,
}: BackgroundDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const file = useRef<HTMLInputElement>(null);
  const current = useEditorState((s) => s.slideBackground);
  // The slide master view sets the background of the layout or master on the canvas only.
  const template = useEditorState((s) => s.presentationView === 'slideMaster');
  // Why OK cannot apply now, or null when it can.
  const refusal = useEditorState((s) =>
    s.status !== 'ready' ? 'disabled.noEditor' : s.mode !== 'edit' ? 'disabled.viewing' : null
  );
  const enabled = refusal === null;
  const [initial] = useState(() => backgroundDraftFrom(current));
  const [custom, setCustom] = useState<GradientFill | null>(null);
  const [pattern, setPattern] = useState<PatternFill | null>(null);
  // Tile a chosen picture at its own size; a change re-picks the chosen picture.
  const [tile, setTile] = useState(() => current?.kind === 'picture' && current.tile === true);
  // Every pick and picture read counts up, so a picture that finishes reading after a later
  // pick is dropped. `picked` turns on with a real pick (Reset too), so OK applies even a value
  // equal to the opening one; setting the opening value itself back (by identity) turns it off.
  const picks = useRef(0);
  const picked = useRef(false);
  const draft: DialogDraft<BackgroundDialogFields> = useDialogDraft(() => ({
    initial,
    validate: (values) => validateBackgroundDraft(values, initial, picked.current),
    resetValues: (): BackgroundDialogFields => ({
      background: null,
      allSlides: draft.values.allSlides,
    }),
    commit: (values) =>
      commitDialogCommand(editor, backgroundDraftCommand(values, initial, picked.current)),
  }));
  // One object per error change, so a host effect on `errors` runs only when they change.
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const state: UseBackgroundDialogReturn = {
    values: draft.values,
    setValue: (name, value) => {
      // Setting the value it already has is no pick, and going back to the opening value
      // undoes the pick: OK must not write the opening snapshot.
      if (name === 'background' && !Object.is(draft.values.background, value)) {
        picks.current += 1;
        picked.current = !Object.is(value, initial.background);
      }
      draft.setValue(name, value);
    },
    errors: translatedErrors,
    isEnabled: enabled,
    apply: () => {
      if (draft.apply()) onClose();
    },
    cancel: onClose,
    reset: () => {
      picks.current += 1;
      picked.current = true;
      draft.reset();
    },
  };
  const pick = (background: SlideBackground): void => state.setValue('background', background);
  const onFile = async (f: File | undefined): Promise<void> => {
    if (!f || !IMAGE_TYPES.includes(f.type)) return;
    const at = ++picks.current;
    let bytes: Uint8Array;
    try {
      bytes = new Uint8Array(await f.arrayBuffer());
    } catch {
      return; // An unreadable file picks nothing.
    }
    if (at === picks.current) pick({ kind: 'picture', bytes, mime: f.type, tile });
  };
  const gradientFrom = backgroundPresetGradient;
  const bases = useThemeBases();
  const chosen = draft.values.background;
  const error = state.errors.allSlides ?? state.errors.background ?? state.errors.form;
  const selectedColor = backgroundDraftColor(draft.values);
  const errorId = useId();
  const presets = bases.map((b) => ({ base: b, gradient: gradientFrom(b.color) }));
  // What the draft holds, for the status line a screen reader announces. A swatch pick shows
  // as its pressed swatch instead.
  const statusKey = backgroundStatusKey(
    chosen,
    initial.background,
    presets.map((p) => p.gradient)
  );
  const status = statusKey ? t(statusKey) : '';
  useCloseOnNewDocument(onClose);
  return (
    <DialogFrame
      kind="background"
      label={t('background.title')}
      className={className}
      style={style}
      onClose={onClose}
      onKeyDown={(event) => applyOnEnter(event, state.isEnabled, state.apply)}
    >
      <parts.Composition
        state={state}
        preset={preset}
        defaults={
          <>
            <div data-pptx-part="header" className="pptx-dialog__header">
              <h2 data-pptx-part="title" className="pptx-dialog__title">
                {t('background.title')}
              </h2>
              <ToolButton label={t('background.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              <section
                data-pptx-part="field"
                data-pptx-field="solid"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('background.color')}</h3>
                <ColorSwatches
                  menu={false}
                  selected={selectedColor}
                  onPick={(color) => pick({ kind: 'solid', color })}
                />
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="gradient"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('background.gradient')}</h3>
                {presets.length ? (
                  <SwatchGroup label={t('toolbar.gradients')} roving>
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
                  </SwatchGroup>
                ) : null}
                <div className="pptx-gradient__row">
                  <button
                    type="button"
                    className="pptx-button"
                    onClick={() =>
                      setCustom(
                        chosen?.kind === 'gradient'
                          ? chosen.gradient
                          : (current.gradient ??
                              gradientFrom(bases[4]?.color ?? BACKGROUND_GRADIENT_FALLBACK))
                      )
                    }
                  >
                    {t('toolbar.gradientCustom')}
                  </button>
                </div>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="pattern"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('background.pattern')}</h3>
                <div className="pptx-gradient__row">
                  {chosen?.kind === 'pattern' ? (
                    <span
                      className="pptx-pattern__chip"
                      aria-hidden="true"
                      style={{ background: patternCss(chosen.pattern) }}
                    />
                  ) : null}
                  <button
                    type="button"
                    className="pptx-button"
                    onClick={() =>
                      setPattern(
                        chosen?.kind === 'pattern'
                          ? chosen.pattern
                          : patternDialogStart(current.pattern, bases[4]?.color)
                      )
                    }
                  >
                    {t('background.choosePattern')}
                  </button>
                </div>
              </section>
              <section
                data-pptx-part="field"
                data-pptx-field="picture"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('background.image')}</h3>
                <div className="pptx-gradient__row">
                  <button
                    type="button"
                    className="pptx-button"
                    onClick={() => file.current?.click()}
                  >
                    {t('background.chooseImage')}
                  </button>
                  <input
                    ref={file}
                    type="file"
                    accept={IMAGE_TYPES.join(',')}
                    hidden
                    onChange={(e) => {
                      void onFile(e.target.files?.[0]);
                      e.target.value = '';
                    }}
                  />
                </div>
                <label className="pptx-radio">
                  <input
                    type="checkbox"
                    checked={tile}
                    onChange={(e) => {
                      setTile(e.target.checked);
                      if (chosen?.kind === 'picture') pick({ ...chosen, tile: e.target.checked });
                    }}
                  />
                  {t('background.tile')}
                </label>
              </section>
              <p role="status" className="pptx-dialog__muted pptx-dialog__status">
                {status}
              </p>
              <label
                data-pptx-part="field"
                data-pptx-field="allSlides"
                className="pptx-radio"
                hidden={template}
              >
                <input
                  type="checkbox"
                  aria-invalid={state.errors.allSlides ? true : undefined}
                  aria-describedby={state.errors.allSlides ? errorId : undefined}
                  checked={draft.values.allSlides}
                  onChange={(e) => draft.setValue('allSlides', e.target.checked)}
                />
                {t('background.allSlides')}
              </label>
            </div>
            <div data-pptx-part="footer" className="pptx-dialog__footer">
              <span id={errorId} data-pptx-part="error" role="alert" className="pptx-dialog__error">
                {error ?? null}
              </span>
              <button
                type="button"
                data-pptx-part="reset"
                className="pptx-button pptx-dialog__reset"
                disabled={!enabled}
                title={refusal ? t(refusal) : undefined}
                onClick={state.reset}
              >
                {t('background.reset')}
              </button>
              <button
                type="button"
                data-pptx-part="cancel"
                className="pptx-button"
                onClick={state.cancel}
              >
                {t('dialog.cancel')}
              </button>
              <button
                type="button"
                data-pptx-part="apply"
                className="pptx-button pptx-button--primary"
                disabled={!enabled}
                title={refusal ? t(refusal) : undefined}
                onClick={state.apply}
              >
                {t('dialog.ok')}
              </button>
            </div>
          </>
        }
      >
        {children}
      </parts.Composition>
      <PatternDialog
        open={pattern !== null}
        initial={pattern}
        onClose={() => setPattern(null)}
        onApply={(p) => {
          pick({ kind: 'pattern', pattern: p });
          setPattern(null);
        }}
      />
      <GradientDialog
        open={custom !== null}
        initial={custom}
        onClose={() => setCustom(null)}
        onApply={(g) => {
          pick({ kind: 'gradient', gradient: g });
          setCustom(null);
        }}
      />
    </DialogFrame>
  );
}

/**
 * Read the draft of the enclosing Background dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('background').open()` does. @public
 */
export function useBackgroundDialog(): UseBackgroundDialogReturn {
  return parts.useState();
}

/**
 * Change the slide background, with replaceable parts (`Header`, `Title`, `Body`, `Footer`,
 * `Field`, `Error`, `Reset`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.BackgroundDialog {...props}>
 *   <PptxEditor.BackgroundDialog.Field name="picture" hidden />
 * </PptxEditor.BackgroundDialog>
 * ```
 *
 * @public
 */
export const BackgroundDialog = Object.assign(BackgroundDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Reset: parts.Reset,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
