/** @jsxImportSource react */
// The Bullets and Numbering dialog (Format > Bullets and Numbering…, or the last row of the
// Bullets and Numbering galleries): a Bulleted and a Numbered tab with the
// gallery, the size as a percentage of the text, the color, and the start number. OK sends one
// `setBulletFormat` command with only what changed, so the dialog is one undo step and a size or
// color change keeps each paragraph's own bullet. The Bulleted tab also takes any character in
// any font, through a Customize… symbol picker, and a picture (PNG, JPEG, or GIF) as a picture
// bullet. The form model is the engine's
// (`bullets-dialog-fields.ts`), so the Vue twin can share it. The host owns visibility.

import { useEffect, useId, useRef, useState, type ReactNode } from 'react';
import {
  BULLET_PICTURE_ACCEPT,
  bulletColorEnabled,
  bulletCommandOf,
  bulletFieldsEnabled,
  bulletFontOptions,
  bulletStyleOptions,
  bulletsDialogError,
  bulletsFieldError,
  customBulletPreview,
  BULLET_TEXT_FONT,
  bulletFontMixed,
  chooseBulletPicture,
  editBulletGlyph,
  formCommandRejection,
  pickBulletStyle,
  readBulletPicture,
  seedBulletFields,
  takeDialogSection,
  type BulletFormatRead,
  type BulletsDialogFields,
  type BulletsDialogTab,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { ColorField } from './dialog-color.tsx';
import {
  FormDialog,
  FormTabPanel,
  FormTabs,
  NumberField,
  SelectField,
  TextField,
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

/** Props for `PptxEditor.BulletsAndNumberingDialog`. @public */
export interface BulletsAndNumberingDialogProps extends DialogCustomizationProps {
  /** Whether the dialog is shown. The host owns this state. */
  readonly open: boolean;
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  readonly onClose: () => void;
}

/**
 * The field names of the Bullets and Numbering dialog, for `BulletsAndNumberingDialog.Field`:
 * the fields of a whole tab (`bullet`, `number`), or one field, named as its input: `style`,
 * `char`, `font`, `picture`, `sizePercent`, and `color` on the Bulleted tab; `numberStyle`,
 * `numberSizePercent`, `numberColor`, and `startAt` on the Numbered tab. @public
 */
export type BulletsAndNumberingDialogFieldName =
  | 'bullet'
  | 'number'
  | 'style'
  | 'char'
  | 'font'
  | 'picture'
  | 'sizePercent'
  | 'color'
  | 'numberStyle'
  | 'numberSizePercent'
  | 'numberColor'
  | 'startAt';

/** The Bullets and Numbering draft and its actions, for a host's own field. @public */
export type UseBulletsAndNumberingDialogReturn = UseDialogReturn<BulletsDialogFields>;

const parts: FormDialogParts<BulletsAndNumberingDialogFieldName, BulletsDialogFields> =
  createDialogParts<BulletsAndNumberingDialogFieldName, UseBulletsAndNumberingDialogReturn>(
    'BulletsAndNumberingDialog'
  );

const TABS = ['bullet', 'number'] as const satisfies readonly BulletsDialogTab[];

function BulletsAndNumberingDialogRoot({
  open,
  onClose,
  className,
  style,
  preset,
  children,
}: BulletsAndNumberingDialogProps): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const idBase = useId();
  // Undefined: the selected paragraphs differ; null: they have no list.
  const current = useEditorState((s) => s.format?.bulletFormat);
  const shared = useEditorState((s) => s.format?.bulletFormatMixed);
  const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
  const [opened, setOpened] = useState<BulletFormatRead | null | undefined>(null);
  const file = useRef<HTMLInputElement>(null);
  // Why the Picture field refused the last file the user picked.
  const [pictureProblem, setPictureProblem] = useState<TranslationKey | null>(null);
  // Each opening starts with no refusal, as the Vue twin mounts its form anew.
  useEffect(() => {
    if (open) setPictureProblem(null);
  }, [open]);
  const { seed, draft, set, update, error, setError } = useDialogDraft(open, () => {
    setOpened(current);
    const tab = takeDialogSection(editor, 'bulletsAndNumbering');
    return seedBulletFields(
      current,
      tab === 'bullet' || tab === 'number' ? tab : undefined,
      shared
    );
  });
  const apply = (): void => {
    if (!draft || !seed) return;
    const problem = bulletsFieldError(seed, draft);
    if (problem) return setError(bulletsDialogError(t, problem, draft.kind));
    const format = bulletCommandOf(seed, draft);
    if (format === undefined) return onClose();
    const refused = formCommandRejection(t, editor?.exec({ type: 'setBulletFormat', format }));
    if (refused === null) onClose();
    else setError({ message: refused });
  };
  const mixed = t('bulletsAndNumberingDialog.mixed');
  const styleOptions = (kind: BulletsDialogTab): (readonly [string, string])[] =>
    bulletStyleOptions(t, kind, seed, draft, opened);
  // The Font field: the text font, the bullet fonts, then the deck's and the host's families.
  // Paragraphs whose bullets differ show "Mixed" until the user picks a style, a character, or
  // a font.
  const fonts: (readonly [string, string])[] = draft
    ? bulletFontOptions(t, seed, editor?.fontFamilies() ?? [])
    : [];
  const fontMixed = !!draft && !!seed && bulletFontMixed(seed, draft);
  // While a tab's style is None, its size, color, and start are dimmed.
  const off = (kind: BulletsDialogTab): boolean =>
    !draft || !bulletFieldsEnabled({ ...draft, kind });
  const fields = (kind: BulletsDialogTab) =>
    draft ? (
      <div className="pptx-dialog__grid" {...fieldPart(kind)}>
        <SelectField
          {...fieldPart(kind === 'bullet' ? 'style' : 'numberStyle')}
          name={kind === 'bullet' ? 'style' : 'numberStyle'}
          label={t('bulletsAndNumberingDialog.style')}
          autoFocus
          mixed={mixed}
          value={draft[kind]}
          options={styleOptions(kind)}
          onChange={(v) =>
            kind === 'bullet' ? update((d) => pickBulletStyle(d, v)) : set(kind, v)
          }
        />
        {kind === 'bullet' ? (
          <>
            <TextField
              {...fieldPart('char')}
              name="char"
              label={t('bulletsAndNumberingDialog.character')}
              disabled={off(kind)}
              value={draft.char}
              font={draft.font}
              // A symbol font's character also shows as its Unicode twin.
              preview={
                draft.char && customBulletPreview(draft) !== draft.char
                  ? customBulletPreview(draft)
                  : undefined
              }
              onChange={(v) => update((d) => editBulletGlyph(d, 'char', v, seed ?? undefined))}
            />
            <SelectField
              {...fieldPart('font')}
              name="font"
              label={t('bulletsAndNumberingDialog.font')}
              disabled={off(kind)}
              value={draft.font || (fontMixed ? '' : BULLET_TEXT_FONT)}
              mixed={fontMixed ? mixed : undefined}
              options={fonts}
              onChange={(v) =>
                update((d) =>
                  editBulletGlyph(d, 'font', v === BULLET_TEXT_FONT ? '' : v, seed ?? undefined)
                )
              }
            />
            {/* A picture file sets a picture bullet; the style row then shows Picture. */}
            <div className="pptx-field" {...fieldPart('picture')}>
              <span>{t('bulletsAndNumberingDialog.picture')}</span>
              <span className="pptx-gradient__row">
                <button
                  type="button"
                  name="picture"
                  className="pptx-button"
                  onClick={() => file.current?.click()}
                >
                  {t('bulletsAndNumberingDialog.choosePicture')}
                </button>
                <span role="status" className="pptx-dialog__muted">
                  {pictureProblem
                    ? t(pictureProblem)
                    : draft.picture
                      ? t('bulletsAndNumberingDialog.pictureChosen')
                      : ''}
                </span>
              </span>
              <input
                ref={file}
                type="file"
                accept={BULLET_PICTURE_ACCEPT}
                hidden
                onChange={(e) => {
                  const picked = e.target.files?.[0];
                  e.target.value = '';
                  void readBulletPicture(picked).then((read) => {
                    if (!read) return;
                    setPictureProblem(read.ok ? null : read.reason);
                    if (read.ok) update((d) => chooseBulletPicture(d, read.picture));
                  });
                }}
              />
            </div>
          </>
        ) : null}
        <NumberField
          {...fieldPart(kind === 'bullet' ? 'sizePercent' : 'numberSizePercent')}
          name={kind === 'bullet' ? 'sizePercent' : 'numberSizePercent'}
          label={t('bulletsAndNumberingDialog.size')}
          unit={t('bulletsAndNumberingDialog.ofText')}
          step="5"
          min="25"
          max="400"
          disabled={off(kind)}
          placeholder={seed?.sizePercent === '' ? mixed : undefined}
          value={draft.sizePercent}
          onChange={(v) => set('sizePercent', v)}
        />
        <ColorField
          {...fieldPart(kind === 'bullet' ? 'color' : 'numberColor')}
          name={kind === 'bullet' ? 'color' : 'numberColor'}
          label={t('bulletsAndNumberingDialog.color')}
          // A picture bullet has no color.
          disabled={!draft || !bulletColorEnabled({ ...draft, kind })}
          value={draft.color}
          seed={seed?.color}
          onChange={(v) => set('color', v)}
        />
        {kind === 'number' ? (
          <NumberField
            {...fieldPart('startAt')}
            name="startAt"
            label={t('bulletsAndNumberingDialog.startAt')}
            unit=""
            step="1"
            min="1"
            max="32767"
            disabled={off(kind)}
            placeholder={seed?.startAt === '' ? mixed : undefined}
            value={draft.startAt}
            onChange={(v) => set('startAt', v)}
          />
        ) : null}
      </div>
    ) : null;
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
      name="bulletsAndNumbering"
      className={className}
      style={style}
      preset={preset}
      parts={parts}
      state={state}
      hostChildren={children}
      title={t('bulletsAndNumberingDialog.title')}
      open={open}
      okLabel={t('bulletsAndNumberingDialog.ok')}
      cancelLabel={t('bulletsAndNumberingDialog.cancel')}
      closeLabel={t('bulletsAndNumberingDialog.close')}
      error={error}
      onClose={onClose}
      onSubmit={apply}
    >
      {draft ? (
        <>
          <FormTabs
            label={t('bulletsAndNumberingDialog.title')}
            idBase={idBase}
            tabs={TABS.map((id) => [id, t(`bulletsAndNumberingDialog.tabs.${id}`)])}
            value={draft.kind}
            onChange={(kind) => (set('kind', kind), setError(null))}
          />
          {/* Both tabs share one cell, so the dialog keeps its size when the tab changes. */}
          <div className="pptx-dialog__panels">
            {TABS.map((kind) => (
              <FormTabPanel key={kind} idBase={idBase} tab={kind} inactive={draft.kind !== kind}>
                {fields(kind)}
              </FormTabPanel>
            ))}
          </div>
        </>
      ) : null}
    </FormDialog>
  );
}

/**
 * Read the draft of the enclosing Bullets and Numbering dialog, for a host's own field inside
 * it. It does not open the dialog: `useEditorDialog('bulletsAndNumbering').open()` does.
 * @public
 */
export function useBulletsAndNumberingDialog(): UseBulletsAndNumberingDialogReturn {
  return parts.useState();
}

/**
 * Format > Bullets and Numbering: the bullet or numbering style, size, color, and start number
 * of the selected paragraphs, applied as one undo step, with replaceable parts (`Header`,
 * `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`).
 *
 * @public
 */
export const BulletsAndNumberingDialog = Object.assign(BulletsAndNumberingDialogRoot, {
  Header: parts.Header,
  Title: parts.Title,
  Body: parts.Body,
  Footer: parts.Footer,
  Field: parts.Field,
  Error: parts.Error,
  Apply: parts.Apply,
  Cancel: parts.Cancel,
});
