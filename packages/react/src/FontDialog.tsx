/** @jsxImportSource react */
// The Font dialog (Format > Font…, the text context menu's Font…, or Character Spacing > More
// Spacing…): the Font tab (Latin font, style, size; font color, underline
// style and underline color; effects with the offset) and the Character Spacing tab (spacing
// and kerning). The fields seed from the selected text when the dialog opens; a field the selected runs disagree about opens blank, and
// a checkbox opens mixed. OK sends only the fields that changed as ONE `setFontFormat` command,
// so the dialog is one undo step. The form model is the engine's (`font-dialog-fields.ts`), so
// the Vue twin can share it. The host owns visibility (`open` / `onClose`).

import { useId, useState, type ReactNode } from 'react';
import {
  UNDERLINE_STYLES,
  fontCommandOf,
  fontDialogError,
  fontDialogFamilies,
  formCommandRejection,
  fontFieldError,
  seedFontFields,
  takeDialogSection,
  withExclusive,
  withScript,
  withSpacing,
  type FontDialogFields,
  type FontDialogTab,
  type FontStyleChoice,
  type SpacingChoice,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { ColorField } from './dialog-color.tsx';
import {
  CheckField,
  FormDialog,
  FormSection,
  FormTabPanel,
  FormTabs,
  NumberField,
  SelectField,
  fieldPart,
  formDialogState,
  useDialogDraft,
  type FormDialogParts,
} from './dialog-form.tsx';
import {
  createDialogParts,
  type DialogCustomizationProps,
  type UseDialogReturn,
} from './dialog-parts.tsx';

/** Props for `PptxEditor.FontDialog`. @public */
export interface FontDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  readonly onClose: () => void;
}

/**
 * The field names of the Font dialog, for `FontDialog.Field`: a section of the Font tab
 * (`allText`, `effects`) or one field, named as its `FontDialogFields` key, on either tab.
 * @public
 */
export type FontDialogFieldName =
  | 'allText'
  | 'effects'
  | 'fontFamily'
  | 'style'
  | 'size'
  | 'color'
  | 'underline'
  | 'underlineColor'
  | 'strike'
  | 'doubleStrike'
  | 'superscript'
  | 'offset'
  | 'subscript'
  | 'smallCaps'
  | 'allCaps'
  | 'spacing'
  | 'by'
  | 'kerning'
  | 'kerningAt';

/** The Font draft and its actions, for a host's own field inside the dialog. @public */
export type UseFontDialogReturn = UseDialogReturn<FontDialogFields>;

const parts: FormDialogParts<FontDialogFieldName, FontDialogFields> = createDialogParts<
  FontDialogFieldName,
  UseFontDialogReturn
>('FontDialog');

const STYLES: readonly FontStyleChoice[] = ['regular', 'italic', 'bold', 'boldItalic'];
const SPACINGS: readonly SpacingChoice[] = ['normal', 'expanded', 'condensed'];

function FontDialogRoot({
  open,
  onClose,
  className,
  style,
  preset,
  children,
}: FontDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const idBase = useId();
  const current = useEditorState((s) => s.format?.font ?? null);
  const textColor = useEditorState((s) => s.format?.color ?? null);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const [tab, setTab] = useState<FontDialogTab>('font');
  const { seed, draft, set, update, error, setError } = useDialogDraft(open, () => {
    setTab(takeDialogSection(editor, 'font') === 'spacing' ? 'spacing' : 'font');
    return seedFontFields(current);
  });
  const apply = (): void => {
    if (!draft || !seed) return;
    const problem = fontFieldError(seed, draft);
    if (problem) {
      // Show the field's tab first, so the error can mark and focus the field.
      setTab(problem.tab);
      return setError(fontDialogError(t, problem));
    }
    const command = fontCommandOf(seed, draft);
    if (!command) return onClose();
    const refused = formCommandRejection(t, editor?.exec({ ...command, type: 'setFontFormat' }));
    if (refused === null) onClose();
    else setError({ message: refused });
  };
  const mixed = t('fontDialog.mixed');
  const pt = t('fontDialog.points');
  const families = draft ? fontDialogFamilies(draft.fontFamily, editor?.fontFamilies() ?? []) : [];
  const blank = (k: 'size' | 'offset' | 'by' | 'kerningAt') =>
    seed?.[k] === '' ? mixed : undefined;
  const state = formDialogState({
    draft,
    set,
    error,
    isEnabled: editable,
    onSubmit: apply,
    onClose,
  });
  return (
    <FormDialog
      name="font"
      className={className}
      style={style}
      preset={preset}
      parts={parts}
      state={state}
      hostChildren={children}
      title={t('fontDialog.title')}
      open={open}
      okLabel={t('fontDialog.ok')}
      cancelLabel={t('fontDialog.cancel')}
      closeLabel={t('fontDialog.close')}
      error={error}
      onClose={onClose}
      onSubmit={apply}
    >
      {draft ? (
        <>
          <FormTabs
            label={t('fontDialog.title')}
            idBase={idBase}
            tabs={(['font', 'spacing'] as const).map((id) => [id, t(`fontDialog.tabs.${id}`)])}
            value={tab}
            onChange={(next) => (setTab(next), setError(null))}
          />
          {/* Both tabs share one cell, so the dialog keeps its size when the tab changes. */}
          <div className="pptx-dialog__panels">
            <FormTabPanel idBase={idBase} tab="font" inactive={tab !== 'font'}>
              <div className="pptx-dialog__grid">
                <SelectField
                  {...fieldPart('fontFamily')}
                  name="fontFamily"
                  label={t('fontDialog.fields.fontFamily')}
                  autoFocus
                  mixed={mixed}
                  value={draft.fontFamily}
                  options={families.map((f) => [f, f])}
                  onChange={(v) => set('fontFamily', v)}
                />
                <SelectField
                  {...fieldPart('style')}
                  name="style"
                  label={t('fontDialog.fields.style')}
                  mixed={mixed}
                  value={draft.style}
                  options={STYLES.map((s) => [s, t(`fontDialog.styles.${s}`)])}
                  onChange={(v) => set('style', v)}
                />
                <NumberField
                  {...fieldPart('size')}
                  name="size"
                  label={t('fontDialog.fields.size')}
                  unit={pt}
                  step="1"
                  min="1"
                  placeholder={blank('size')}
                  value={draft.size}
                  onChange={(v) => set('size', v)}
                />
              </div>
              <FormSection {...fieldPart('allText')} title={t('fontDialog.allText')}>
                <div className="pptx-dialog__grid">
                  <ColorField
                    {...fieldPart('color')}
                    name="color"
                    label={t('fontDialog.fields.color')}
                    value={draft.color}
                    seed={seed?.color}
                    current={textColor}
                    onChange={(v) => set('color', v)}
                  />
                  <SelectField
                    {...fieldPart('underline')}
                    name="underline"
                    label={t('fontDialog.fields.underline')}
                    mixed={mixed}
                    value={draft.underline}
                    options={UNDERLINE_STYLES.map((u) => [u, t(`fontDialog.underlines.${u}`)])}
                    onChange={(v) => set('underline', v)}
                  />
                  <ColorField
                    {...fieldPart('underlineColor')}
                    name="underlineColor"
                    label={t('fontDialog.fields.underlineColor')}
                    value={draft.underline === 'none' ? 'auto' : draft.underlineColor}
                    seed={seed?.underlineColor}
                    disabled={draft.underline === 'none'}
                    onChange={(v) => set('underlineColor', v)}
                  />
                </div>
              </FormSection>
              <FormSection {...fieldPart('effects')} title={t('fontDialog.effects')}>
                {/* Two columns: the strike and script pairs, then the caps pair. */}
                <div className="pptx-dialog__columns">
                  <div className="pptx-dialog__rows">
                    <CheckField
                      {...fieldPart('strike')}
                      name="strike"
                      label={t('fontDialog.strikethrough')}
                      value={draft.strike}
                      onChange={(on) => update((d) => withExclusive(d, 'strike', on))}
                    />
                    <CheckField
                      {...fieldPart('doubleStrike')}
                      name="doubleStrike"
                      label={t('fontDialog.doubleStrikethrough')}
                      value={draft.doubleStrike}
                      onChange={(on) => update((d) => withExclusive(d, 'doubleStrike', on))}
                    />
                    <CheckField
                      {...fieldPart('superscript')}
                      name="superscript"
                      label={t('fontDialog.superscript')}
                      value={draft.superscript}
                      onChange={(on) => update((d) => withScript(d, 'superscript', on))}
                    />
                    <CheckField
                      {...fieldPart('subscript')}
                      name="subscript"
                      label={t('fontDialog.subscript')}
                      value={draft.subscript}
                      onChange={(on) => update((d) => withScript(d, 'subscript', on))}
                    />
                    {/* The offset of either script: after both checkboxes, not between them. */}
                    <NumberField
                      {...fieldPart('offset')}
                      name="offset"
                      label={t('fontDialog.fields.offset')}
                      unit={t('fontDialog.percent')}
                      step="1"
                      min="0"
                      placeholder={blank('offset')}
                      disabled={draft.superscript !== true && draft.subscript !== true}
                      value={draft.offset}
                      onChange={(v) => set('offset', v)}
                    />
                  </div>
                  <div className="pptx-dialog__rows">
                    <CheckField
                      {...fieldPart('smallCaps')}
                      name="smallCaps"
                      label={t('fontDialog.smallCaps')}
                      value={draft.smallCaps}
                      onChange={(on) => update((d) => withExclusive(d, 'smallCaps', on))}
                    />
                    <CheckField
                      {...fieldPart('allCaps')}
                      name="allCaps"
                      label={t('fontDialog.allCaps')}
                      value={draft.allCaps}
                      onChange={(on) => update((d) => withExclusive(d, 'allCaps', on))}
                    />
                  </div>
                </div>
              </FormSection>
            </FormTabPanel>
            <FormTabPanel idBase={idBase} tab="spacing" inactive={tab !== 'spacing'}>
              <div className="pptx-dialog__inline">
                <SelectField
                  {...fieldPart('spacing')}
                  name="spacing"
                  label={t('fontDialog.fields.spacing')}
                  autoFocus
                  mixed={mixed}
                  value={draft.spacing}
                  options={SPACINGS.map((s) => [s, t(`fontDialog.spacings.${s}`)])}
                  onChange={(v) => update((d) => withSpacing(d, v))}
                />
                <NumberField
                  {...fieldPart('by')}
                  name="by"
                  label={t('fontDialog.fields.by')}
                  unit={pt}
                  step="0.1"
                  min="0"
                  placeholder={blank('by')}
                  disabled={draft.spacing === 'normal' || draft.spacing === ''}
                  value={draft.by}
                  onChange={(v) => set('by', v)}
                />
              </div>
              <div className="pptx-dialog__inline">
                <CheckField
                  {...fieldPart('kerning')}
                  name="kerning"
                  label={t('fontDialog.kerning')}
                  value={draft.kerning}
                  onChange={(on) => set('kerning', on)}
                />
                <NumberField
                  {...fieldPart('kerningAt')}
                  name="kerningAt"
                  label={t('fontDialog.fields.kerningAt')}
                  unit={pt}
                  step="1"
                  placeholder={blank('kerningAt')}
                  disabled={draft.kerning !== true}
                  value={draft.kerningAt}
                  onChange={(v) => set('kerningAt', v)}
                />
              </div>
            </FormTabPanel>
          </div>
        </>
      ) : null}
    </FormDialog>
  );
}

/**
 * Read the draft of the enclosing Font dialog, for a host's own field inside it. It does not
 * open the dialog: `useEditorDialog('font').open()` does. @public
 */
export function useFontDialog(): UseFontDialogReturn {
  return parts.useState();
}

/**
 * Format > Font: the font, style, size, underline, effects, character spacing, and kerning of
 * the selected text, applied as one undo step, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @public
 */
export const FontDialog = Object.assign(FontDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
