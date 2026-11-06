/** @jsxImportSource vue */
// The toolbar's controls, one per registry shape or adapter kind: the Vue twin of the controls
// in the React adapter's `Toolbar.tsx`. `ToolbarControl` picks the part by core's
// `toolbarPartKind` and by `SHAPED_PARTS`, so a slot added to the registry renders in both
// adapters without editing this file. The rules (dropdown rows, the split color, the font
// filter) are core's (`toolbar-controls.ts`).

import {
  Fragment,
  defineComponent,
  h,
  inject,
  ref,
  type Component,
  type PropType,
  type VNode,
  type ExtractPublicPropTypes,
} from 'vue';
import {
  colorMenuSelection,
  chromeOfferedControl,
  chromeSlotIsToggle,
  requestDialogSection,
  dispatchFindRequest,
  runChromeRow,
  surfaceInputIn,
  toolbarColorResetKey,
  toolbarDropdownRow,
  toolbarPartKind,
  toolbarRowSection,
  toolbarSplitDefaultColor,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
  type ToolbarPartKind,
  type ToolbarPickedColor,
} from '@pptx-editor.dev/core/editor';
import {
  Icon,
  MenuItem,
  Popover,
  ToolButton,
  focusText,
  keepFocus,
  useAriaDisabled,
} from './chrome.tsx';
import { ColorSwatches } from './ColorSwatches.tsx';
import {
  chromeActionHostKey,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from './composables.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';
import { SizeControl } from './toolbar-size.tsx';
import { FontFamilyControl } from './FontFamilyControl.tsx';
import { ImageControl, LinkControl } from './toolbar-file-controls.tsx';
import { FillControl, LineMenuControl, ShapeGalleryControl } from './ShapeControls.tsx';
import { EffectMenuControl, ShapeStyleGalleryControl } from './ShapeFormatControls.tsx';
import {
  LayoutControl,
  NewSlideControl,
  SectionAddControl,
  SectionRenameControl,
} from './SlideControls.tsx';
import { TableGridControl, TableStyleControl } from './TableControls.tsx';
import { CorrectionsControl } from './PictureControls.tsx';

const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** A toolbar button that runs the slot's command; a toggle shows `aria-pressed`. */
const IconControl = defineComponent({
  name: 'PptxIconControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    return () => (
      <ToolButton
        slot={props.slot}
        label={label.value}
        paths={props.control.paths}
        onClick={() => cmd.execute()}
        disabled={!cmd.isEnabled.value}
        disabledReason={reason.value}
        // The registry says which controls are toggles; `aria-pressed` goes on those only.
        pressed={chromeSlotIsToggle(props.slot) ? cmd.isActive.value : undefined}
      />
    );
  },
});

/** A value menu: a radio group of choices, or a menu of actions, with dialog rows. */
const DropdownControl = defineComponent({
  name: 'PptxDropdownControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const host = inject(chromeActionHostKey, null);
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    // A dropdown with no icon is the arrow half of a split button (Bullets and Numbering).
    const icon = (): readonly string[] | null | undefined =>
      props.control.options?.find((o) => o.value === cmd.value.value)?.paths ?? props.control.paths;
    const trigger = useTriggerAttributes(props.slot, '', () => ({
      disabled: !cmd.isEnabled.value,
    }));
    const openRow = (dialog: Parameters<NonNullable<typeof host>['openDialog']>[0]): void => {
      open.value = false;
      // Focus returns to the text when the dialog closes: the dialog host takes the element
      // that has focus as it opens.
      const text = anchor.value
        ?.closest('.pptx-editor')
        ?.querySelector<HTMLElement>('.pptx-surface__input');
      (text ?? anchor.value)?.focus({ preventScroll: true });
      const section = toolbarRowSection(props.slot);
      if (editor.value && section) requestDialogSection(editor.value, dialog, section);
      host?.openDialog(dialog);
    };
    return () => {
      const paths = icon();
      const base = `pptx-toolbar__button ${paths ? 'pptx-toolbar__button--wide' : 'pptx-toolbar__button--caret'}`;
      const attrs = trigger();
      // The choices this build offers: a choice whose result does not paint stays out.
      const options = chromeOfferedControl(props.slot, props.control).options ?? [];
      const glyphs = options.some((x) => x.glyph);
      return [
        <button
          ref={anchor}
          type="button"
          {...attrs}
          class={[base, attrs.class].filter(Boolean).join(' ')}
          data-slot={props.slot}
          aria-label={label.value}
          aria-haspopup="menu"
          aria-expanded={open.value}
          {...tip.value}
          {...off.attrs.value}
          onMousedown={keepFocus}
          onClick={() => (open.value = !open.value)}
        >
          {paths ? <Icon paths={paths} /> : null}
          <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
        </button>,
        off.note(),
        open.value ? (
          <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
            {{
              default: () =>
                options.map((o) => {
                  const row = toolbarDropdownRow(
                    editor.value,
                    host,
                    props.slot,
                    props.control,
                    o,
                    cmd.value.value
                  );
                  const dialog = row.dialog;
                  return (
                    <Fragment key={o.value}>
                      {dialog ? <div class="pptx-menu__sep" role="separator" /> : null}
                      <MenuItem
                        {...(row.checked !== undefined ? { checked: row.checked } : {})}
                        {...(row.refused
                          ? {
                              disabled: true,
                              ...(row.reason ? { reason: translateKeyOrText(t, row.reason) } : {}),
                            }
                          : {})}
                        onSelect={() => {
                          if (dialog) return openRow(dialog);
                          // A pick gives the caret back to the slide text, so the next key
                          // (typing, Cmd+Z) reaches the slide.
                          cmd.execute(o.value);
                          focusText(editor.value, anchor.value);
                          open.value = false;
                        }}
                      >
                        {{
                          default: () => [
                            o.paths ? (
                              <Icon paths={o.paths} />
                            ) : o.glyph || glyphs ? (
                              // A row with no preview keeps the column, so the labels line up.
                              <span class="pptx-menu__glyph" aria-hidden="true">
                                {o.glyph}
                              </span>
                            ) : null,
                            t(o.labelKey, o.labelParams),
                          ],
                        }}
                      </MenuItem>
                    </Fragment>
                  );
                }),
            }}
          </Popover>
        ) : null,
      ];
    };
  },
});

/** A split color button: the main half applies the last color, the arrow opens the swatches. */
const ColorControl = defineComponent({
  name: 'PptxColorControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason, tip } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const anchor = ref<HTMLButtonElement | null>(null);
    const open = ref(false);
    const text1 = useEditorState((s) => s.themeColors[0]?.[1]?.color);
    // The current color of the selection, marked in the menu.
    const selected = useEditorState((s) => colorMenuSelection(props.slot, s));
    const picked = ref<ToolbarPickedColor | null>(null);
    const off = useAriaDisabled(() => !cmd.isEnabled.value, reason);
    const close = (): void => {
      focusText(editor.value, anchor.value);
      open.value = false;
    };
    return () => {
      const last = picked.value ?? toolbarSplitDefaultColor(props.slot, text1.value);
      return (
        <span class="pptx-toolbar__group">
          <ToolButton
            slot={props.slot}
            label={label.value}
            paths={props.control.paths}
            disabled={!cmd.isEnabled.value}
            disabledReason={reason.value}
            asChild={false}
            onClick={() => cmd.execute(last.choice)}
          >
            {{
              default: () => (
                <span class="pptx-toolbar__swatch" style={{ background: last.color }} />
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
            <Popover anchor={anchor.value} onClose={() => (open.value = false)}>
              {{
                default: () => [
                  <ColorSwatches
                    selected={selected.value}
                    onPick={(choice, color) => {
                      picked.value = { choice, color };
                      cmd.execute(choice);
                      close();
                    }}
                  />,
                  <div class="pptx-menu__sep" />,
                  <MenuItem onSelect={() => (cmd.execute(null), close())}>
                    {{ default: () => t(toolbarColorResetKey(props.slot)) }}
                  </MenuItem>,
                ],
              }}
            </Popover>
          ) : null}
        </span>
      );
    };
  },
});

/** A control that opens a panel the adapter owns: the find bar, to find or to replace. */
const PanelControl = defineComponent({
  name: 'PptxPanelControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const host = inject(chromeActionHostKey, null);
    return () => {
      // A `popups.find` of `false`, the slide sorter, or no find bar turns the button off.
      // Read the command state before `off`: a render that stops at `off` would not track it,
      // and the control would stay off after the reason goes away.
      const enabled = cmd.isEnabled.value;
      const why = reason.value;
      const off = host?.unavailable?.(props.slot) ?? null;
      return (
        <ToolButton
          slot={props.slot}
          label={label.value}
          paths={props.control.paths}
          disabled={!!off || !enabled}
          disabledReason={off ? translateKeyOrText(t, off) : why}
          onClick={() => dispatchFindRequest(editor.value, props.slot === 'find.replace')}
        />
      );
    };
  },
});

/**
 * A control the adapter runs itself: a registry dialog (Page setup, Format background), the
 * clipboard, or the slide show. Core's `runChromeRow` picks the action.
 */
const ActionControl = defineComponent({
  name: 'PptxActionControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const host = inject(chromeActionHostKey, null);
    const wrap = ref<HTMLElement | null>(null);
    const run = (): void => {
      if (!host) return void cmd.execute();
      // A dialog's focus returns to this button: the dialog host takes the focused element.
      const button = wrap.value?.querySelector<HTMLElement>('[data-slot]') ?? null;
      runChromeRow(
        editor.value,
        {
          ...host,
          openDialog: (key) => {
            button?.focus({ preventScroll: true });
            host.openDialog(key);
          },
        },
        props.slot,
        undefined,
        surfaceInputIn(wrap.value?.closest('.pptx-editor'))
      );
    };
    return () => {
      // A `popups` entry of `false` turns the button off. A slot that also takes a value
      // (`image.altText`) keeps its command enabled, so the button asks the host here.
      // Read the command state before `off`: a render that stops at `off` would not track it,
      // and the control would stay off after the reason goes away.
      const enabled = cmd.isEnabled.value;
      const why = reason.value;
      const off = host?.unavailable?.(props.slot) ?? null;
      return (
        <span ref={wrap} class="pptx-toolbar__group">
          <ToolButton
            hasPopup={props.control.state.kind === 'dialog' ? 'dialog' : undefined}
            slot={props.slot}
            label={label.value}
            paths={props.control.paths}
            disabled={!!off || !enabled}
            disabledReason={off ? translateKeyOrText(t, off) : why}
            onClick={run}
          />
        </span>
      );
    };
  },
});

/**
 * Slots whose control is more than its registry shape: a searchable font list, a link editor,
 * a fill with gradients, and line menus that draw each option. As in DOCX Editor's
 * `SHAPED_PARTS`, this picks a renderer; it never decides which slots the toolbar shows.
 */
const SHAPED_PARTS: Partial<Record<ChromeSlotId, Component>> = {
  'font.family': FontFamilyControl,
  'text.link': LinkControl,
  'insert.line': LineMenuControl,
  'shape.fill': FillControl,
  'shape.lineStart': LineMenuControl,
  'shape.lineEnd': LineMenuControl,
  'shape.styles': ShapeStyleGalleryControl,
  'shape.effects': EffectMenuControl,
  'slide.add': NewSlideControl,
  'slide.layout': LayoutControl,
  'section.add': SectionAddControl,
  'section.rename': SectionRenameControl,
  'table.style': TableStyleControl,
  'image.corrections': CorrectionsControl,
};

/** The renderer for each part kind (core's `toolbarPartKind`), as React's `KIND_PARTS`. */
const KIND_PARTS: Record<ToolbarPartKind, Component> = {
  icon: IconControl,
  stepper: SizeControl,
  dropdown: DropdownControl,
  colorSplit: ColorControl,
  gallery: ShapeGalleryControl,
  tableGrid: TableGridControl,
  load: ImageControl,
  panel: PanelControl,
  action: ActionControl,
};

const toolbarControlProps = {
  group: { type: String, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** The props of `ToolbarControl`, the twin of React's `ToolbarControlProps`. @public */
export type ToolbarControlProps = ExtractPublicPropTypes<typeof toolbarControlProps>;

/** One registered control, by registry shape, as React's `PptxEditor.ToolbarControl`. @public */
export const ToolbarControl = defineComponent({
  name: 'ToolbarControl',
  props: toolbarControlProps,
  setup(props) {
    return (): VNode => {
      const slot = `${props.group}.${props.control.id}` as ChromeSlotId;
      const Part = SHAPED_PARTS[slot] ?? KIND_PARTS[toolbarPartKind(props.control)];
      return h(Part, { slot, control: props.control });
    };
  },
});
