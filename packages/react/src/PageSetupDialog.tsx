/** @jsxImportSource react */
// The Page setup dialog (File → Page setup…): slide size, appearance, and fonts. The host owns
// visibility (`open` / `onClose`), as with DOCX Editor's `PageSetupDialog`. Built on the dialog
// parts (`dialog-parts.tsx`): every element is a named part a host can restyle, replace, or
// hide. The slide size and the appearance choice are a draft that OK applies and Cancel
// discards. A new slide size is one `setSlideSize` command, so one undo step restores it.

import { useMemo, useState, type ReactNode } from 'react';
import {
  SLIDE_SIZE_CHOICES,
  commitDialogCommand,
  radioGroupStep,
  seedSlideSizeFields,
  slideSizeCommandOf,
  slideSizeFieldError,
  translateDialogErrors,
  withSlideDimension,
  withSlideOrientation,
  withSlideSizePreset,
  type SlideOrientation,
  type SlideSizeFields,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { CheckField, NumberField, SelectField } from './dialog-form.tsx';
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
 * The draft of the Page setup dialog. `theme` keeps the name of the dialog's `theme` and
 * `onThemeChange` props on `main`. @public
 */
export interface PageSetupDialogFields {
  /** The chrome theme the appearance choice shows. */
  readonly theme: 'light' | 'dark';
  /** The slide size: a preset, the width and height in inches, and whether content scales. */
  readonly slideSize: SlideSizeFields;
}

/** The Page setup draft and its actions. @public */
export type UsePageSetupDialogReturn = UseDialogReturn<PageSetupDialogFields>;

/** The field names of the Page setup dialog, for `PageSetupDialog.Field`. @public */
export type PageSetupDialogFieldName = 'slideSize' | 'theme' | 'fonts' | 'substitutes';

/** Props for `PptxEditor.PageSetupDialog`. @public */
export interface PageSetupDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK. */
  readonly onClose: () => void;
  /** The chrome theme; the appearance choice shows only with `onThemeChange`. */
  readonly theme?: 'light' | 'dark';
  /** Called by OK when the appearance choice changed. */
  readonly onThemeChange?: (theme: 'light' | 'dark') => void;
}

const parts = createDialogParts<PageSetupDialogFieldName, UsePageSetupDialogReturn>(
  'PageSetupDialog'
);
const THEMES = ['light', 'dark'] as const;
const ORIENTATIONS: readonly SlideOrientation[] = ['landscape', 'portrait'];

/** A segmented radio group with the arrow keys of the WAI-ARIA radio group pattern. */
function Segmented<T extends string>(props: {
  readonly label: string;
  readonly choices: readonly T[];
  readonly value: T;
  readonly disabled?: boolean;
  readonly labelOf: (choice: T) => string;
  readonly onChange: (choice: T) => void;
}): ReactNode {
  return (
    <div
      className="pptx-segmented"
      role="radiogroup"
      aria-label={props.label}
      aria-disabled={props.disabled || undefined}
      onKeyDown={(event) => {
        if (props.disabled) return;
        const next = radioGroupStep(props.choices, props.value, event.key);
        if (!next) return;
        event.preventDefault();
        props.onChange(next);
        event.currentTarget.querySelector<HTMLElement>(`[data-value="${next}"]`)?.focus();
      }}
    >
      {props.choices.map((m) => (
        <button
          key={m}
          type="button"
          role="radio"
          data-value={m}
          disabled={props.disabled}
          aria-checked={props.value === m}
          tabIndex={props.value === m && !props.disabled ? 0 : -1}
          onClick={() => props.onChange(m)}
        >
          {props.labelOf(m)}
        </button>
      ))}
    </div>
  );
}

function PageSetupDialogRoot(props: PageSetupDialogProps): ReactNode {
  return props.open ? <PageSetupDialogOpen {...props} /> : null;
}

function PageSetupDialogOpen({
  onClose,
  theme = 'light',
  onThemeChange,
  className,
  style,
  preset = true,
  children,
}: PageSetupDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const width = useEditorState((s) => s.slideWidth);
  const height = useEditorState((s) => s.slideHeight);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const embedded = useEditorState((s) => s.embeddedFontFaces);
  const substituted = useEditorState((s) => s.substitutedFonts);
  // OK reports a change from the opening choice only, as the Vue twin: a host theme change
  // while the dialog is open is not reverted by an untouched OK.
  const [opening] = useState(theme);
  // The slide size the dialog opened on: OK sends a size only when the fields moved from it.
  const [opened] = useState(() => ({ width, height }));
  const [seed] = useState(() => seedSlideSizeFields(width, height));
  const draft = useDialogDraft<PageSetupDialogFields>(() => ({
    initial: { theme: opening, slideSize: seed },
    validate: (values) => {
      const problem = slideSizeFieldError(seed, values.slideSize, opened);
      return problem ? { slideSize: problem.key } : {};
    },
    commit: (values) => {
      const resized = commitDialogCommand(
        editor,
        slideSizeCommandOf(seed, values.slideSize, opened)
      );
      if (!resized.ok) return resized;
      if (values.theme !== opening) onThemeChange?.(values.theme);
      return { ok: true };
    },
  }));
  const size = draft.values.slideSize;
  const setSize = (next: SlideSizeFields): void => draft.setValue('slideSize', next);
  const translatedErrors = useMemo(() => translateDialogErrors(t, draft.errors), [t, draft.errors]);
  const sizeError = translatedErrors.slideSize;
  const state: UsePageSetupDialogReturn = {
    values: draft.values,
    setValue: draft.setValue,
    errors: translatedErrors,
    isEnabled: true,
    apply: () => {
      if (draft.apply()) onClose();
    },
    cancel: onClose,
  };
  useCloseOnNewDocument(onClose);
  return (
    <DialogFrame
      kind="pageSetup"
      label={t('pageSetup.title')}
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
                {t('pageSetup.title')}
              </h2>
              <ToolButton label={t('pageSetup.close')} icon="close" onClick={onClose} />
            </div>
            <div data-pptx-part="body" className="pptx-dialog__body">
              <section
                data-pptx-part="field"
                data-pptx-field="slideSize"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('pageSetup.slideSize')}</h3>
                {editable ? null : (
                  // Viewing mode: the size as text, as the fields cannot change it.
                  <p className="pptx-dialog__value">
                    {t(`pageSetup.slideSizeName.${seed.size}`)}
                    <span className="pptx-dialog__muted">
                      {t('pageSetup.slideSizeValue', { width: seed.width, height: seed.height })}
                    </span>
                  </p>
                )}
                <div className="pptx-dialog__rows" hidden={!editable}>
                  <SelectField
                    name="size"
                    label={t('pageSetup.size')}
                    disabled={!editable}
                    value={size.size}
                    options={SLIDE_SIZE_CHOICES.map(
                      (c) => [c, t(`pageSetup.slideSizeName.${c}`)] as const
                    )}
                    onChange={(c) => setSize(withSlideSizePreset(size, c))}
                  />
                  <div className="pptx-dialog__grid">
                    {(['width', 'height'] as const).map((field) => (
                      <NumberField
                        key={field}
                        name={field}
                        label={t(`pageSetup.${field}`)}
                        unit={t('pageSetup.inches')}
                        step="0.01"
                        min="1"
                        max="56"
                        disabled={!editable}
                        value={size[field]}
                        onChange={(v) => setSize(withSlideDimension(size, field, v))}
                      />
                    ))}
                  </div>
                  <Segmented
                    label={t('pageSetup.orientation')}
                    choices={ORIENTATIONS}
                    value={size.orientation}
                    disabled={!editable}
                    labelOf={(o) => t(`pageSetup.${o}`)}
                    onChange={(o) => setSize(withSlideOrientation(size, o))}
                  />
                  <CheckField
                    name="scaleContent"
                    label={t('pageSetup.scaleContent')}
                    value={size.scaleContent}
                    disabled={!editable}
                    onChange={(on) => setSize({ ...size, scaleContent: on })}
                  />
                  {sizeError ? (
                    <p role="alert" className="pptx-dialog__error">
                      {sizeError}
                    </p>
                  ) : null}
                </div>
              </section>
              {onThemeChange ? (
                <section
                  data-pptx-part="field"
                  data-pptx-field="theme"
                  className="pptx-dialog__section"
                >
                  <h3 className="pptx-dialog__label">{t('pageSetup.appearance')}</h3>
                  <Segmented
                    label={t('pageSetup.appearance')}
                    choices={THEMES}
                    value={draft.values.theme}
                    labelOf={(m) => t(`pageSetup.${m}`)}
                    onChange={(m) => draft.setValue('theme', m)}
                  />
                </section>
              ) : null}
              <section
                data-pptx-part="field"
                data-pptx-field="fonts"
                className="pptx-dialog__section"
              >
                <h3 className="pptx-dialog__label">{t('pageSetup.fonts')}</h3>
                <p className="pptx-dialog__value">
                  {embedded
                    ? t('pageSetup.fontsEmbedded', { count: embedded })
                    : t('pageSetup.fontsNone')}
                </p>
                {/* No substitutes: one line here, and no empty Substituted fonts section. */}
                {substituted.length === 0 ? (
                  <p className="pptx-dialog__value">{t('pageSetup.substitutesNone')}</p>
                ) : null}
              </section>
              {substituted.length > 0 ? (
                <section
                  data-pptx-part="field"
                  data-pptx-field="substitutes"
                  className="pptx-dialog__section"
                >
                  <h3 className="pptx-dialog__label">{t('pageSetup.substitutes')}</h3>
                  <ul className="pptx-dialog__list">
                    {substituted.map((f) => (
                      <li key={f.family}>
                        {t('pageSetup.substitution', {
                          family: f.family,
                          substitute: f.substitute,
                        })}
                        {f.metricCompatible ? (
                          <span className="pptx-chip pptx-chip--quiet">
                            {t('pageSetup.metricCompatible')}
                          </span>
                        ) : null}
                      </li>
                    ))}
                  </ul>
                </section>
              ) : null}
            </div>
            <div data-pptx-part="footer" className="pptx-dialog__footer">
              <span data-pptx-part="error" role="alert" className="pptx-dialog__error">
                {state.errors.form ?? null}
              </span>
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
                disabled={!state.isEnabled}
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
    </DialogFrame>
  );
}

/**
 * Read the draft of the enclosing Page setup dialog, for a host's own field inside it. It does
 * not open the dialog: `useEditorDialog('pageSetup').open()` does. @public
 */
export function usePageSetupDialog(): UsePageSetupDialogReturn {
  return parts.useState();
}

/**
 * Page setup: slide size, appearance, and fonts, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @example
 * ```tsx
 * <PptxEditor.PageSetupDialog {...props}>
 *   <PptxEditor.PageSetupDialog.Field name="substitutes" hidden />
 *   <PptxEditor.PageSetupDialog.Apply className="acme-primary" />
 * </PptxEditor.PageSetupDialog>
 * ```
 *
 * @public
 */
export const PageSetupDialog = Object.assign(PageSetupDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
