/** @jsxImportSource vue */
// The font family picker: the Vue twin of the React adapter's `FontFamilyControl.tsx`, after DOCX
// Editor's `FontFamily` toolbar compound (docx-editor/packages/vue/src/editor/toolbar/
// FontFamily.tsx): a root that holds the picker state in a part-level context (provide/inject),
// and the `Trigger`, `Content`, and `Item` parts that read it. With no slot content the root
// renders a trigger and a search field over a list box of groups (the slide's theme fonts,
// labeled Headings and Body; the fonts in this presentation; all fonts), each name drawn in its
// own face, the current family marked and scrolled into view when the list opens. With slot
// content, a host composes the parts itself. Core's `font-picker.ts` builds the groups and rows.

import {
  defineComponent,
  inject,
  nextTick,
  normalizeClass,
  provide,
  ref,
  useId,
  watch,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type Ref,
} from 'vue';
import {
  firstFontPickerFamily,
  focusFontOption,
  fontPickerGroups,
  fontPickerRow,
  isCurrentFontRow,
  revealCurrentFont,
  warnInvalidFontItem,
  warnPartOutsideCompound,
  type ChromeControl,
  type ChromeSlotId,
  type FontPickerGroup,
  type FontPickerRow,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, focusText, keepFocus, useAriaDisabled } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

/** The font the text uses: its family name and the theme font it comes from, if any. */
interface CurrentFont {
  readonly family: string | null;
  readonly theme: 'major' | 'minor' | null;
}

/** The picker state the parts share; every read is current in the reader's render. */
interface FontFamilyState {
  readonly slot: () => ChromeSlotId;
  readonly label: () => string;
  readonly open: Ref<boolean>;
  readonly query: Ref<string>;
  readonly groups: () => readonly FontPickerGroup[];
  readonly current: () => CurrentFont;
  /** Apply a row's value (a theme font reference or a family) and close the list. */
  readonly pick: (value: string) => void;
  /** The row a host's `Item value` stands for, or null for a value that is not a font name. */
  readonly rowFor: (value: string) => FontPickerRow | null;
  readonly anchor: Ref<HTMLButtonElement | null>;
  readonly list: Ref<HTMLElement | null>;
  readonly trigger: ReturnType<typeof useTriggerAttributes>;
  readonly slotState: ReturnType<typeof useSlot>;
  readonly off: ReturnType<typeof useAriaDisabled>;
}

const fontFamilyKey: InjectionKey<FontFamilyState> = Symbol('pptxFontFamily');

/** The shared state, or null (with a warning) outside `PptxEditorToolbar.FontFamily`. */
function useFontFamilyState(part: string): FontFamilyState | null {
  const state = inject(fontFamilyKey, null);
  if (!state)
    warnPartOutsideCompound(`PptxEditorToolbar.FontFamily.${part}`, 'PptxEditorToolbar.FontFamily');
  return state;
}

const fontFamilyPartProps = {
  /** A class name added to the part's element, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

const fontFamilyItemProps = {
  ...fontFamilyPartProps,
  /**
   * What the item applies: a family name, or a theme font reference (`+mj-lt` for the headings
   * font, `+mn-lt` for the body font), so the text follows a later theme change.
   */
  value: { type: String, required: true as const },
} as const;

/**
 * Props of `PptxEditorToolbar.FontFamily.Trigger` and `.Content`, as React's. The default slot
 * replaces the part's default content. @public
 */
export type FontFamilyPartProps = ExtractPublicPropTypes<typeof fontFamilyPartProps>;

/** Props of `PptxEditorToolbar.FontFamily.Item`, as React's. @public */
export type FontFamilyItemProps = ExtractPublicPropTypes<typeof fontFamilyItemProps>;

/**
 * The button that opens the list: the current family, or the label when the selection has no
 * one family. The default slot replaces its content.
 */
const FontFamilyTrigger = defineComponent({
  name: 'FontFamilyTrigger',
  inheritAttrs: false,
  props: fontFamilyPartProps,
  setup(props, { attrs, slots }) {
    const state = useFontFamilyState('Trigger');
    return () => {
      if (!state) return null;
      const { cmd, label, tip } = state.slotState;
      const trigger = state.trigger();
      return [
        <button
          ref={state.anchor}
          type="button"
          {...trigger}
          class={[trigger.class, props.className, attrs.class]}
          data-slot={state.slot()}
          aria-label={label.value}
          aria-haspopup="dialog"
          aria-expanded={state.open.value}
          {...tip.value}
          {...state.off.attrs.value}
          onMousedown={keepFocus}
          onClick={() => {
            state.query.value = '';
            state.open.value = !state.open.value;
          }}
        >
          {slots.default?.() ?? [
            <span class="pptx-toolbar__family-label">{cmd.value.value ?? label.value}</span>,
            <Icon name="expand_more" className="pptx-icon pptx-toolbar__caret" />,
          ]}
        </button>,
        state.off.note(),
      ];
    };
  },
});

/** One family row: an `option` with the name in its own face and the theme role, if any. */
const FontRow = defineComponent({
  name: 'PptxFontRow',
  props: {
    row: { type: Object as PropType<FontPickerRow>, required: true as const },
    rowClass: { type: String, default: undefined },
  },
  setup(props, { slots }) {
    const { t } = useTranslation();
    const state = inject(fontFamilyKey, null);
    return () => {
      if (!state) return null;
      const row = props.row;
      const now = state.current();
      const current = !row.disabled && isCurrentFontRow(row, now.family, now.theme);
      return (
        <button
          type="button"
          role="option"
          aria-selected={current}
          tabindex={-1}
          disabled={row.disabled}
          class={['pptx-menu__item pptx-font-picker__option', props.rowClass]}
          data-family={row.family || undefined}
          data-theme={row.theme ?? undefined}
          onClick={() => state.pick(row.value)}
        >
          {slots.default?.() ?? [
            // The face is a quoted, validated family name (core's `fontFaceValue`), set as a
            // custom property through a style object, never a CSS string.
            row.family ? (
              <span
                class="pptx-font-picker__face"
                style={row.face ? { '--pptx-font-face': row.face } : undefined}
              >
                {row.family}
              </span>
            ) : null,
            row.roleKey ? <span class="pptx-font-picker__role">{t(row.roleKey)}</span> : null,
          ]}
          {current ? <Icon name="check" className="pptx-icon pptx-menu__check" /> : null}
        </button>
      );
    };
  },
});

/**
 * One option of the list: it applies `value` and carries the check mark while the text uses
 * that font. The default slot replaces the name and role. A theme font that the slide's theme
 * leaves empty shows its role, disabled. A value that is not a font name renders nothing, with
 * a development warning.
 */
const FontFamilyItem = defineComponent({
  name: 'FontFamilyItem',
  inheritAttrs: false,
  props: fontFamilyItemProps,
  setup(props, { attrs, slots }) {
    const state = useFontFamilyState('Item');
    return () => {
      if (!state) return null;
      const row = state.rowFor(props.value);
      if (!row) {
        warnInvalidFontItem('PptxEditorToolbar.FontFamily.Item', props.value);
        return null;
      }
      return (
        <FontRow row={row} rowClass={normalizeClass([props.className, attrs.class]) || undefined}>
          {slots.default ? { default: slots.default } : undefined}
        </FontRow>
      );
    };
  },
});

/** One labeled group of the list box. */
const FontGroup = defineComponent({
  name: 'PptxFontGroup',
  props: {
    group: { type: Object as PropType<FontPickerGroup>, required: true as const },
  },
  setup(props) {
    const { t } = useTranslation();
    const id = useId();
    return () => (
      <div class="pptx-font-picker__group" role="group" aria-labelledby={id}>
        <div id={id} role="presentation" class="pptx-menu__heading pptx-font-picker__heading">
          {t(props.group.labelKey)}
        </div>
        {props.group.rows.map((row) => (
          <FontRow key={row.value} row={row} />
        ))}
      </div>
    );
  },
});

/**
 * The list, while it is open. With no slot content: a search field over the theme,
 * presentation, and all-fonts groups. With slot content: a list box of the host's `Item` parts.
 */
const FontFamilyContent = defineComponent({
  name: 'FontFamilyContent',
  inheritAttrs: false,
  props: fontFamilyPartProps,
  setup(props, { attrs, slots }) {
    const state = useFontFamilyState('Content');
    const { t } = useTranslation();
    return () => {
      if (!state || !state.open.value) return null;
      const groups = state.groups();
      const custom = slots.default;
      return (
        <Popover
          anchor={state.anchor.value}
          role="dialog"
          field
          onClose={() => (state.open.value = false)}
        >
          {{
            default: () => (
              <div class={['pptx-font-picker', props.className, attrs.class]}>
                {custom ? null : (
                  <input
                    class="pptx-font-picker__search"
                    type="search"
                    placeholder={t('toolbar.fontSearch')}
                    aria-label={t('toolbar.fontSearch')}
                    value={state.query.value}
                    onInput={(e: Event) =>
                      (state.query.value = (e.target as HTMLInputElement).value)
                    }
                    onKeydown={(e: KeyboardEvent) => {
                      if (e.key !== 'Enter') return;
                      // The key stops here, so it does not press the trigger and open the list
                      // again.
                      e.preventDefault();
                      const first = firstFontPickerFamily(groups);
                      if (first) state.pick(first);
                    }}
                  />
                )}
                <div
                  ref={state.list}
                  class="pptx-font-picker__list"
                  role="listbox"
                  aria-label={state.label()}
                >
                  {custom
                    ? custom()
                    : [
                        ...groups.map((g) => <FontGroup key={g.id} group={g} />),
                        groups.length === 0 ? (
                          <p class="pptx-font-picker__empty">{t('toolbar.fontNone')}</p>
                        ) : null,
                      ]}
                </div>
              </div>
            ),
          }}
        </Popover>
      );
    };
  },
});

/**
 * The font family picker: the toolbar's part for `font.family`, and the root of the
 * `PptxEditorToolbar.FontFamily` compound. Its default slot holds `Trigger` and `Content` parts
 * in place of the defaults. @internal
 */
export const FontFamilyControl = defineComponent({
  name: 'PptxFamilyControl',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    control: { type: Object as PropType<ChromeControl>, required: true as const },
  },
  setup(props, { slots }) {
    const slotState = useSlot(
      () => props.slot,
      () => props.control
    );
    const { cmd, label, reason } = slotState;
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const editor = usePptxEditor();
    const anchor = ref<HTMLButtonElement | null>(null);
    const list = ref<HTMLElement | null>(null);
    const open = ref(false);
    const query = ref('');
    // The groups change with the deck (embedded and host fonts) and with the slide's theme.
    const revision = useEditorState((s) => s.revision);
    const slide = useEditorState((s) => s.currentSlide);
    // The check mark follows what the text uses: a theme font marks its theme row only.
    const theme = useEditorState((s) => s.format?.fontFamilyTheme ?? null);
    const trigger = useTriggerAttributes(
      props.slot,
      'pptx-toolbar__field pptx-toolbar__field--family',
      () => ({ disabled: !cmd.isEnabled.value })
    );
    // Once, when the list opens: the current family comes into view.
    // After the popover took focus (its `onMounted`), a list with no search field focuses its
    // current option, so the arrow keys reach the host's `Item` parts.
    watch(open, async (isOpen) => {
      if (!isOpen) return;
      await nextTick();
      revealCurrentFont(list.value);
      focusFontOption(list.value);
    });
    provide(fontFamilyKey, {
      slot: () => props.slot,
      label: () => label.value,
      open,
      query,
      groups: () => {
        void revision.value;
        return open.value ? fontPickerGroups(editor.value, slide.value, query.value) : [];
      },
      current: () => ({
        family: typeof cmd.value.value === 'string' ? cmd.value.value : null,
        theme: theme.value,
      }),
      // A theme row applies its theme font reference (`+mj-lt`, `+mn-lt`), any other row its
      // family.
      pick: (value) => {
        cmd.execute(value);
        focusText(editor.value, anchor.value);
        open.value = false;
      },
      rowFor: (value) => {
        void revision.value;
        return fontPickerRow(editor.value, slide.value, value);
      },
      anchor,
      list,
      trigger,
      slotState,
      off,
    });
    return () => slots.default?.() ?? [<FontFamilyTrigger />, <FontFamilyContent />];
  },
});

/**
 * The sub-parts of `PptxEditorToolbar.FontFamily`, as DOCX Editor's `FontFamily` statics.
 * @internal
 */
export const FONT_FAMILY_PARTS = {
  Trigger: FontFamilyTrigger,
  Content: FontFamilyContent,
  Item: FontFamilyItem,
} as const;
