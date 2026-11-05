/** @jsxImportSource vue */
// The Font dialog (Format > Font…, the text context menu's Font…, or Character Spacing > More
// Spacing…): the Vue twin of the React adapter's `FontDialog`. The Font tab (Latin font, style,
// size; font color, underline style and underline color; effects with the offset) and the
// Character Spacing tab (spacing and kerning). The fields seed from the selected text when the
// dialog opens; a field the selected runs disagree about opens blank, and a checkbox opens
// mixed. OK sends only the fields that changed as ONE `setFontFormat` command, so the dialog is
// one undo step. The form model is core's (`font-dialog-fields.ts`). The host owns visibility
// (`open`, `onClose`).

import {
  defineComponent,
  h,
  ref,
  useId,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  UNDERLINE_STYLES,
  fontCommandOf,
  fontDialogError,
  fontDialogFamilies,
  fontFieldError,
  seedFontFields,
  takeDialogSection,
  withExclusive,
  withScript,
  withSpacing,
  type ColorPick,
  type FontDialogFields,
  type FontDialogTab,
  type FontStyleChoice,
  type SpacingChoice,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { createDialogParts, type UseDialogReturn } from './dialog-parts.ts';
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
  type FormDialogParts,
} from './dialog-form.tsx';
import { useFormDraft } from './form-draft.ts';

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
  UseFontDialogReturn,
  FontDialogFieldName
>('FontDialog');

const STYLES: readonly FontStyleChoice[] = ['regular', 'italic', 'bold', 'boldItalic'];
const SPACINGS: readonly SpacingChoice[] = ['normal', 'expanded', 'condensed'];

const fontDialogProps = {
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

/** Props of `FontDialog`, the twin of React's `FontDialogProps`. @public */
export type FontDialogProps = ExtractPublicPropTypes<typeof fontDialogProps>;

const FontDialogOpen = defineComponent({
  name: 'FontDialogOpen',
  inheritAttrs: false,
  props: fontDialogProps,
  setup(props, { attrs, slots }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const idBase = useId();
    const current = useEditorState((s) => s.format?.font ?? null);
    const textColor = useEditorState((s) => s.format?.color ?? null);
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    // Character Spacing > More Spacing… opens the dialog on its spacing tab.
    const tab = ref<FontDialogTab>(
      takeDialogSection(editor.value, 'font') === 'spacing' ? 'spacing' : 'font'
    );
    const seed = seedFontFields(current.value);
    const draft = useFormDraft<FontDialogFields>({
      seed,
      problem: (values) => {
        const problem = fontFieldError(seed, values);
        if (!problem) return null;
        // Show the field's tab first, so the error can mark and focus the field.
        tab.value = problem.tab;
        return fontDialogError(t, problem);
      },
      command: (values) => {
        const command = fontCommandOf(seed, values);
        return command ? { ...command, type: 'setFontFormat' } : null;
      },
    });
    const state = formDialogState({ draft, isEnabled: editable, onClose: () => props.onClose() });
    const set = draft.setValue;
    return () => {
      const d = draft.values.value;
      const mixed = t('fontDialog.mixed');
      const pt = t('fontDialog.points');
      const families = fontDialogFamilies(d.fontFamily, editor.value?.fontFamilies() ?? []);
      const blank = (k: 'size' | 'offset' | 'by' | 'kerningAt') =>
        seed[k] === '' ? mixed : undefined;
      return (
        <FormDialog
          {...attrs}
          name="font"
          className={props.className}
          style={props.style}
          preset={props.preset}
          parts={parts}
          state={state as UseDialogReturn<object>}
          title={t('fontDialog.title')}
          okLabel={t('fontDialog.ok')}
          cancelLabel={t('fontDialog.cancel')}
          closeLabel={t('fontDialog.close')}
        >
          {{
            default: () => [
              <FormTabs
                label={t('fontDialog.title')}
                idBase={idBase}
                tabs={(['font', 'spacing'] as const).map(
                  (id) => [id, t(`fontDialog.tabs.${id}`)] as const
                )}
                value={tab.value}
                onChange={(next: FontDialogTab) => {
                  tab.value = next;
                  draft.clearErrors();
                }}
              />,
              // Both tabs share one cell, so the dialog keeps its size when the tab changes.
              <div class="pptx-dialog__panels">
                <FormTabPanel idBase={idBase} tab="font" inactive={tab.value !== 'font'}>
                  {{
                    default: () => (
                      <>
                        <div class="pptx-dialog__grid">
                          <SelectField
                            {...fieldPart('fontFamily')}
                            name="fontFamily"
                            label={t('fontDialog.fields.fontFamily')}
                            autoFocus
                            mixed={mixed}
                            value={d.fontFamily}
                            options={families.map((f) => [f, f] as const)}
                            onChange={(v: string) => set('fontFamily', v)}
                          />
                          <SelectField
                            {...fieldPart('style')}
                            name="style"
                            label={t('fontDialog.fields.style')}
                            mixed={mixed}
                            value={d.style}
                            options={STYLES.map((s) => [s, t(`fontDialog.styles.${s}`)] as const)}
                            onChange={(v: FontDialogFields['style']) => set('style', v)}
                          />
                          <NumberField
                            {...fieldPart('size')}
                            name="size"
                            label={t('fontDialog.fields.size')}
                            unit={pt}
                            step="1"
                            min="1"
                            placeholder={blank('size')}
                            value={d.size}
                            onChange={(v: string) => set('size', v)}
                          />
                        </div>
                        <FormSection {...fieldPart('allText')} title={t('fontDialog.allText')}>
                          {{
                            default: () => (
                              <>
                                <div class="pptx-dialog__grid">
                                  <ColorField
                                    {...fieldPart('color')}
                                    name="color"
                                    label={t('fontDialog.fields.color')}
                                    value={d.color}
                                    seed={seed.color}
                                    current={textColor.value}
                                    onChange={(v: ColorPick) => set('color', v)}
                                  />
                                  <SelectField
                                    {...fieldPart('underline')}
                                    name="underline"
                                    label={t('fontDialog.fields.underline')}
                                    mixed={mixed}
                                    value={d.underline}
                                    options={UNDERLINE_STYLES.map(
                                      (u) => [u, t(`fontDialog.underlines.${u}`)] as const
                                    )}
                                    onChange={(v: FontDialogFields['underline']) =>
                                      set('underline', v)
                                    }
                                  />
                                  <ColorField
                                    {...fieldPart('underlineColor')}
                                    name="underlineColor"
                                    label={t('fontDialog.fields.underlineColor')}
                                    value={d.underline === 'none' ? 'auto' : d.underlineColor}
                                    seed={seed.underlineColor}
                                    disabled={d.underline === 'none'}
                                    onChange={(v: ColorPick) => set('underlineColor', v)}
                                  />
                                </div>
                              </>
                            ),
                          }}
                        </FormSection>
                        <FormSection {...fieldPart('effects')} title={t('fontDialog.effects')}>
                          {{
                            default: () => (
                              <>
                                {/* Two columns: the strike and script pairs, then the caps pair. */}
                                <div class="pptx-dialog__columns">
                                  <div class="pptx-dialog__rows">
                                    <CheckField
                                      {...fieldPart('strike')}
                                      name="strike"
                                      label={t('fontDialog.strikethrough')}
                                      value={d.strike}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) => withExclusive(cur, 'strike', on))
                                      }
                                    />
                                    <CheckField
                                      {...fieldPart('doubleStrike')}
                                      name="doubleStrike"
                                      label={t('fontDialog.doubleStrikethrough')}
                                      value={d.doubleStrike}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) =>
                                          withExclusive(cur, 'doubleStrike', on)
                                        )
                                      }
                                    />
                                    <CheckField
                                      {...fieldPart('superscript')}
                                      name="superscript"
                                      label={t('fontDialog.superscript')}
                                      value={d.superscript}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) => withScript(cur, 'superscript', on))
                                      }
                                    />
                                    <CheckField
                                      {...fieldPart('subscript')}
                                      name="subscript"
                                      label={t('fontDialog.subscript')}
                                      value={d.subscript}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) => withScript(cur, 'subscript', on))
                                      }
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
                                      disabled={d.superscript !== true && d.subscript !== true}
                                      value={d.offset}
                                      onChange={(v: string) => set('offset', v)}
                                    />
                                  </div>
                                  <div class="pptx-dialog__rows">
                                    <CheckField
                                      {...fieldPart('smallCaps')}
                                      name="smallCaps"
                                      label={t('fontDialog.smallCaps')}
                                      value={d.smallCaps}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) => withExclusive(cur, 'smallCaps', on))
                                      }
                                    />
                                    <CheckField
                                      {...fieldPart('allCaps')}
                                      name="allCaps"
                                      label={t('fontDialog.allCaps')}
                                      value={d.allCaps}
                                      onChange={(on: boolean) =>
                                        draft.update((cur) => withExclusive(cur, 'allCaps', on))
                                      }
                                    />
                                  </div>
                                </div>
                              </>
                            ),
                          }}
                        </FormSection>
                      </>
                    ),
                  }}
                </FormTabPanel>
                <FormTabPanel idBase={idBase} tab="spacing" inactive={tab.value !== 'spacing'}>
                  {{
                    default: () => (
                      <>
                        <div class="pptx-dialog__inline">
                          <SelectField
                            {...fieldPart('spacing')}
                            name="spacing"
                            label={t('fontDialog.fields.spacing')}
                            autoFocus
                            mixed={mixed}
                            value={d.spacing}
                            options={SPACINGS.map(
                              (s) => [s, t(`fontDialog.spacings.${s}`)] as const
                            )}
                            onChange={(v: SpacingChoice) =>
                              draft.update((cur) => withSpacing(cur, v))
                            }
                          />
                          <NumberField
                            {...fieldPart('by')}
                            name="by"
                            label={t('fontDialog.fields.by')}
                            unit={pt}
                            step="0.1"
                            min="0"
                            placeholder={blank('by')}
                            disabled={d.spacing === 'normal' || d.spacing === ''}
                            value={d.by}
                            onChange={(v: string) => set('by', v)}
                          />
                        </div>
                        <div class="pptx-dialog__inline">
                          <CheckField
                            {...fieldPart('kerning')}
                            name="kerning"
                            label={t('fontDialog.kerning')}
                            value={d.kerning}
                            onChange={(on: boolean) => set('kerning', on)}
                          />
                          <NumberField
                            {...fieldPart('kerningAt')}
                            name="kerningAt"
                            label={t('fontDialog.fields.kerningAt')}
                            unit={pt}
                            step="1"
                            placeholder={blank('kerningAt')}
                            disabled={d.kerning !== true}
                            value={d.kerningAt}
                            onChange={(v: string) => set('kerningAt', v)}
                          />
                        </div>
                      </>
                    ),
                  }}
                </FormTabPanel>
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
 * Format > Font: the font, style, size, underline, effects, character spacing, and kerning of
 * the selected text, applied as one undo step, with replaceable parts (`Header`, `Title`,
 * `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`), as React's `FontDialog`. Renders
 * nothing while `open` is false. @public
 */
export const FontDialog = Object.assign(
  defineComponent({
    name: 'FontDialog',
    inheritAttrs: false,
    props: fontDialogProps,
    setup(props, { attrs, slots }) {
      return () => (props.open ? h(FontDialogOpen, { ...attrs, ...props }, slots) : null);
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
 * Read the draft of the enclosing Font dialog, for a host's own field inside it. It does not
 * open the dialog: `useEditorDialog('font').open()` does. @public
 */
export function useFontDialog(): UseFontDialogReturn {
  return parts.useState();
}
