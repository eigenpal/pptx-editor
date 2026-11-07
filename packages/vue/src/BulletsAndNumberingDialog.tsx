/** @jsxImportSource vue */
// The Bullets and Numbering dialog (Format > Bullets and Numbering…, or the last row of the
// Bullets and Numbering galleries): the Vue twin of the React adapter's
// `BulletsAndNumberingDialog`. A Bulleted and a Numbered tab with the gallery, the size as a
// percentage of the text, the color, and the start number. OK sends one `setBulletFormat`
// command with only what changed, so the dialog is one undo step and a size or color change
// keeps each paragraph's own bullet. The Bulleted tab also takes any character in any font. The
// form model is core's (`bullets-dialog-fields.ts`). The host owns visibility.

import {
  defineComponent,
  h,
  useId,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
import {
  BULLET_TEXT_FONT,
  bulletCommandOf,
  bulletFieldsEnabled,
  bulletFontMixed,
  bulletFontOptions,
  bulletStyleOptions,
  bulletsDialogError,
  bulletsFieldError,
  customBulletPreview,
  editBulletGlyph,
  pickBulletStyle,
  seedBulletFields,
  takeDialogSection,
  type BulletsDialogFields,
  type BulletsDialogTab,
  type ColorPick,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { createDialogParts, type UseDialogReturn } from './dialog-parts.ts';
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
  type FormDialogParts,
} from './dialog-form.tsx';
import { useFormDraft } from './form-draft.ts';

/**
 * The field names of the Bullets and Numbering dialog, for `BulletsAndNumberingDialog.Field`:
 * the fields of a whole tab (`bullet`, `number`), or one field, named as its input: `style`,
 * `char`, `font`, `sizePercent`, and `color` on the Bulleted tab; `numberStyle`,
 * `numberSizePercent`, `numberColor`, and `startAt` on the Numbered tab. @public
 */
export type BulletsAndNumberingDialogFieldName =
  | 'bullet'
  | 'number'
  | 'style'
  | 'char'
  | 'font'
  | 'sizePercent'
  | 'color'
  | 'numberStyle'
  | 'numberSizePercent'
  | 'numberColor'
  | 'startAt';

/** The Bullets and Numbering draft and its actions, for a host's own field. @public */
export type UseBulletsAndNumberingDialogReturn = UseDialogReturn<BulletsDialogFields>;

const parts: FormDialogParts<BulletsAndNumberingDialogFieldName, BulletsDialogFields> =
  createDialogParts<UseBulletsAndNumberingDialogReturn, BulletsAndNumberingDialogFieldName>(
    'BulletsAndNumberingDialog'
  );

const TABS = ['bullet', 'number'] as const satisfies readonly BulletsDialogTab[];

const bulletsAndNumberingDialogProps = {
  /** Whether the dialog is shown. The host owns this state. */
  open: { type: Boolean, required: true as const },
  /** Called on Cancel, Escape, a click on the backdrop, and after OK applies. */
  onClose: { type: Function as PropType<() => void>, required: true as const },
  /** `false`: only the default slot, with no default arrangement. */
  preset: { type: Boolean, default: true },
  /** A class name added to the dialog surface, as React's `className` (`class` works too). */
  className: { type: String, default: undefined },
  /** Styles added to the dialog surface. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
} as const;

/**
 * Props of `BulletsAndNumberingDialog`, the twin of React's `BulletsAndNumberingDialogProps`.
 * @public
 */
export type BulletsAndNumberingDialogProps = ExtractPublicPropTypes<
  typeof bulletsAndNumberingDialogProps
>;

const BulletsAndNumberingDialogOpen = defineComponent({
  name: 'BulletsAndNumberingDialogOpen',
  inheritAttrs: false,
  props: bulletsAndNumberingDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const idBase = useId();
    // Undefined: the selected paragraphs differ; null: they have no list.
    const current = useEditorState((s) => s.format?.bulletFormat);
    const shared = useEditorState((s) => s.format?.bulletFormatMixed);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const opened = current.value;
    const section = takeDialogSection(editor.value, 'bulletsAndNumbering');
    const seed = seedBulletFields(
      opened,
      section === 'bullet' || section === 'number' ? section : undefined,
      shared.value
    );
    const draft = useFormDraft<BulletsDialogFields>({
      seed,
      problem: (values) => {
        const problem = bulletsFieldError(seed, values);
        return problem ? bulletsDialogError(t, problem, values.kind) : null;
      },
      command: (values) => {
        const format = bulletCommandOf(seed, values);
        return format === undefined ? null : { type: 'setBulletFormat', format };
      },
    });
    const state = formDialogState({ draft, isEnabled: editable, onClose: () => props.onClose() });
    const set = draft.setValue;
    const fields = (kind: BulletsDialogTab, d: BulletsDialogFields, mixed: string): VNode => {
      // While a tab's style is None, its size, color, and start are dimmed.
      const off = !bulletFieldsEnabled({ ...d, kind });
      const fontMixed = bulletFontMixed(seed, d);
      return (
        <div class="pptx-dialog__grid" {...fieldPart(kind)}>
          <SelectField
            {...fieldPart(kind === 'bullet' ? 'style' : 'numberStyle')}
            name={kind === 'bullet' ? 'style' : 'numberStyle'}
            label={t('bulletsAndNumberingDialog.style')}
            autoFocus
            mixed={mixed}
            value={d[kind]}
            options={bulletStyleOptions(t, kind, seed, d, opened)}
            onChange={(v: string) =>
              kind === 'bullet' ? draft.update((cur) => pickBulletStyle(cur, v)) : set(kind, v)
            }
          />
          {kind === 'bullet'
            ? [
                <TextField
                  {...fieldPart('char')}
                  name="char"
                  label={t('bulletsAndNumberingDialog.character')}
                  disabled={off}
                  value={d.char}
                  font={d.font}
                  // A symbol font's character also shows as its Unicode twin.
                  preview={
                    d.char && customBulletPreview(d) !== d.char ? customBulletPreview(d) : undefined
                  }
                  onChange={(v: string) =>
                    draft.update((cur) => editBulletGlyph(cur, 'char', v, seed))
                  }
                />,
                <SelectField
                  {...fieldPart('font')}
                  name="font"
                  label={t('bulletsAndNumberingDialog.font')}
                  disabled={off}
                  value={d.font || (fontMixed ? '' : BULLET_TEXT_FONT)}
                  mixed={fontMixed ? mixed : undefined}
                  options={bulletFontOptions(t, seed, editor.value?.fontFamilies() ?? [])}
                  onChange={(v: string) =>
                    draft.update((cur) =>
                      editBulletGlyph(cur, 'font', v === BULLET_TEXT_FONT ? '' : v, seed)
                    )
                  }
                />,
              ]
            : null}
          <NumberField
            {...fieldPart(kind === 'bullet' ? 'sizePercent' : 'numberSizePercent')}
            name={kind === 'bullet' ? 'sizePercent' : 'numberSizePercent'}
            label={t('bulletsAndNumberingDialog.size')}
            unit={t('bulletsAndNumberingDialog.ofText')}
            step="5"
            min="25"
            max="400"
            disabled={off}
            placeholder={seed.sizePercent === '' ? mixed : undefined}
            value={d.sizePercent}
            onChange={(v: string) => set('sizePercent', v)}
          />
          <ColorField
            {...fieldPart(kind === 'bullet' ? 'color' : 'numberColor')}
            name={kind === 'bullet' ? 'color' : 'numberColor'}
            label={t('bulletsAndNumberingDialog.color')}
            disabled={off}
            value={d.color}
            seed={seed.color}
            onChange={(v: ColorPick) => set('color', v)}
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
              disabled={off}
              placeholder={seed.startAt === '' ? mixed : undefined}
              value={d.startAt}
              onChange={(v: string) => set('startAt', v)}
            />
          ) : null}
        </div>
      ) as VNode;
    };
    return () => {
      const d = draft.values.value;
      const mixed = t('bulletsAndNumberingDialog.mixed');
      return (
        <FormDialog
          {...attrs}
          name="bulletsAndNumbering"
          className={props.className}
          style={props.style}
          preset={props.preset}
          parts={parts}
          state={state as UseDialogReturn<object>}
          title={t('bulletsAndNumberingDialog.title')}
          okLabel={t('bulletsAndNumberingDialog.ok')}
          cancelLabel={t('bulletsAndNumberingDialog.cancel')}
          closeLabel={t('bulletsAndNumberingDialog.close')}
        >
          {{
            default: () => [
              <FormTabs
                label={t('bulletsAndNumberingDialog.title')}
                idBase={idBase}
                tabs={TABS.map((id) => [id, t(`bulletsAndNumberingDialog.tabs.${id}`)] as const)}
                value={d.kind}
                onChange={(kind: BulletsDialogTab) => set('kind', kind)}
              />,
              // Both tabs share one cell, so the dialog keeps its size when the tab changes.
              <div class="pptx-dialog__panels">
                {TABS.map((kind) => (
                  <FormTabPanel key={kind} idBase={idBase} tab={kind} inactive={d.kind !== kind}>
                    {{ default: () => <>{fields(kind, d, mixed)}</> }}
                  </FormTabPanel>
                ))}
              </div>,
            ],
            host: () => slots.default?.(),
          }}
        </FormDialog>
      );
    };
  },
});

/**
 * Format > Bullets and Numbering: the bullet or numbering style, size, color, and start number
 * of the selected paragraphs, applied as one undo step, with replaceable parts (`Header`,
 * `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`), as React's
 * `BulletsAndNumberingDialog`. Renders nothing while `open` is false. @public
 */
export const BulletsAndNumberingDialog = Object.assign(
  defineComponent({
    name: 'BulletsAndNumberingDialog',
    inheritAttrs: false,
    props: bulletsAndNumberingDialogProps,
    setup(props, { attrs, slots }) {
      return () =>
        props.open ? h(BulletsAndNumberingDialogOpen, { ...attrs, ...props }, slots) : null;
    },
  }),
  {
    Header: parts.Header,
    Title: parts.Title,
    Body: parts.Body,
    Footer: parts.Footer,
    Field: parts.Field,
    Error: parts.Error,
    Apply: parts.Apply,
    Cancel: parts.Cancel,
  }
);

/**
 * Read the draft of the enclosing Bullets and Numbering dialog, for a host's own field inside
 * it. It does not open the dialog: `useEditorDialog('bulletsAndNumbering').open()` does.
 * @public
 */
export function useBulletsAndNumberingDialog(): UseBulletsAndNumberingDialogReturn {
  return parts.useState();
}
