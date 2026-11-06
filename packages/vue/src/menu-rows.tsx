/** @jsxImportSource vue */
// Menu rows over chrome slots, for the Vue menus and context menus: the Vue twin of the React
// adapter's `menu-rows.tsx`, after DOCX Editor's `Menu.Item` and `ContextMenu.Slot`. A row takes
// its label, icon, shortcut, enabled state, and action from the slot's registry entry and core's
// row rules (`resolveChromeMenuRow`, `chromeRowState`, `menuRowCheck`, `runChromeRow`), so a
// menu row and its toolbar twin cannot drift. A slot that collects a value (Insert > Shape,
// Insert > Line, Insert > Table, New Slide, Layout) opens a submenu of its choices.

import {
  defineComponent,
  inject,
  provide,
  type InjectionKey,
  type PropType,
  type VNode,
} from 'vue';
import {
  SHAPE_GALLERY,
  chromeControlForSlot,
  chromeOfferedControl,
  chromeRowState,
  editorCommandFor,
  menuEntryKey,
  menuRowCheck,
  resolveChromeMenuRow,
  runChromeRow,
  sameChromeCommandState,
  surfaceInputIn,
  tidySeparators,
  translateKeyOrText,
  warnMenuPart,
  type ChromeActionHost,
  type ChromeMenuEntry,
  type ChromeMenuItemEntry,
  type ChromeSlotId,
  type MenuPartNames,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import {
  chromeActionHostKey,
  useEditorState,
  usePptxEditor,
  useTranslation,
} from './composables.ts';
import { MenuRow, MenuSeparator, MenuSubmenu, menuCloseKey } from './menu-row.tsx';
import { LineGlyph, ShapeGrid, ShapeIcon } from './ShapeControls.tsx';
import { LayoutGallery } from './SlideControls.tsx';
import { MenuTableGrid } from './TableControls.tsx';

export { menuCloseKey, tidySeparators };

/** The Vue names of the menu bar and the root, for the menu warnings. */
export const VUE_MENU_PARTS: MenuPartNames = { bar: 'PptxEditorMenu', root: 'PptxEditorRoot' };

/** The close function of the enclosing menu, or a no-op. @internal */
export function useMenuClose(): () => void {
  return inject(menuCloseKey, null) ?? noop;
}

function noop(): void {}

/** The host overrides of one registry row. @internal */
export interface SlotRowOverrides {
  readonly labelKey?: string;
  readonly shortcutKey?: string;
  readonly className?: string;
  /** A Material Symbol name, or icon content in place of the registry icon. */
  readonly icon?: string | VNode;
}

/**
 * Run a row's slot. Outside a root there is no action host: the slot's own command runs, and an
 * adapter action (a dialog, a file pick) is refused with the reason, as `useEditorCommand`.
 */
function runRow(
  editor: PptxEditorInstance | null,
  host: ChromeActionHost | null,
  slot: ChromeSlotId,
  value: string | undefined,
  input?: HTMLElement | null
): void {
  if (host) runChromeRow(editor, host, slot, value, input);
  else editorCommandFor(editor, slot, null, 'disabled.noRoot').execute(value);
}

/** One registry row: a slot, or a slot with a value. @internal */
export const SlotMenuRow = defineComponent({
  name: 'PptxSlotMenuRow',
  props: {
    entry: { type: Object as PropType<ChromeMenuItemEntry>, required: true as const },
    overrides: { type: Object as PropType<SlotRowOverrides>, default: undefined },
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    // The host knows which `popups` entries are off: their rows show disabled with the reason.
    const host = inject(chromeActionHostKey, null);
    const close = useMenuClose();
    const state = useEditorState(
      () => chromeRowState(editor.value, props.entry.slot, props.entry.value, host ?? undefined),
      sameChromeCommandState
    );
    return () => {
      const { entry, overrides } = props;
      const row = resolveChromeMenuRow(entry);
      const control = chromeControlForSlot(entry.slot);
      if (!row || !control) {
        warnMenuPart('PptxEditorMenu.Item', 'unknownSlot', entry.slot, VUE_MENU_PARTS);
        return null;
      }
      const label = overrides?.labelKey
        ? translateKeyOrText(t, overrides.labelKey)
        : translateKeyOrText(t, row.labelKey, row.labelParams);
      const shortcutKey = overrides?.shortcutKey ?? row.shortcutKey;
      const reason = state.value.disabledReason
        ? translateKeyOrText(t, state.value.disabledReason)
        : null;
      // One rule for every menu part: a string is a Material Symbol name, a vnode is content.
      const named = typeof overrides?.icon === 'string' ? overrides.icon : undefined;
      // Insert > Line has no registry icon: the row draws a line, as its options do.
      const glyph =
        named !== undefined
          ? undefined
          : ((overrides?.icon as VNode | undefined) ??
            (row.paths === null && entry.slot === 'insert.line' ? (
              <LineGlyph value="line" />
            ) : undefined));
      const common = {
        label,
        disabled: !state.value.isEnabled,
        reason,
        ...(glyph !== undefined
          ? { glyph }
          : named !== undefined
            ? { icon: named, paths: null }
            : { paths: row.paths }),
        ...(overrides?.className ? { className: overrides.className } : {}),
      };
      // A slot that collects a value opens its choices, unless the row names one.
      if (entry.value === undefined && control.state.kind === 'value' && !entry.direct) {
        if (entry.slot === 'slide.add' || entry.slot === 'slide.layout')
          return <LayoutSubmenu slot={entry.slot} {...common} />;
        // Insert > Shape cascades its shape grid. Quick Styles is a gallery with no menu form
        // yet, so its row falls through to the disabled "pick in the toolbar" row below.
        if (entry.slot === 'insert.shape') return <GallerySubmenu slot={entry.slot} {...common} />;
        // Insert > Table: the size grid, as DOCX Editor's `Menu.TableGrid`.
        if (control.shape === 'tableGrid')
          return (
            <MenuSubmenu {...common}>
              {{
                default: () => (
                  <MenuTableGrid slot={entry.slot} label={label} autoFocus onInserted={close} />
                ),
              }}
            </MenuSubmenu>
          );
        // The choices this build offers: a choice whose result does not paint stays out.
        const offered = chromeOfferedControl(entry.slot, control);
        // Shape Effects: one submenu for each effect kind, as its toolbar menu has.
        if (offered.optionGroups?.length && offered.options?.length)
          return (
            <MenuSubmenu {...common}>
              {{
                default: () =>
                  offered.optionGroups!.map((group) => (
                    <MenuSubmenu key={group.id} label={t(group.labelKey)}>
                      {{
                        default: () =>
                          offered
                            .options!.filter((o) => o.group === group.id)
                            .map((o) => (
                              <SlotMenuRow
                                key={o.value}
                                entry={{ kind: 'item', slot: entry.slot, value: o.value }}
                              />
                            )),
                      }}
                    </MenuSubmenu>
                  )),
              }}
            </MenuSubmenu>
          );
        if (offered.options?.length)
          return (
            <MenuSubmenu {...common}>
              {{
                default: () =>
                  offered.options!.map((o) => (
                    <SlotMenuRow
                      key={o.value}
                      entry={{ kind: 'item', slot: entry.slot, value: o.value }}
                    />
                  )),
              }}
            </MenuSubmenu>
          );
        // A picker with no menu form (a color, a font, a size) cannot run from a row.
        if (entry.slot !== 'text.link') {
          warnMenuPart('PptxEditorMenu.Item', 'valueRow', entry.slot, VUE_MENU_PARTS);
          return <MenuRow {...common} disabled reason={t('menu.pickInToolbar')} slot={row.key} />;
        }
      }
      // A two-state row (Bold, Ruler) is a checkbox row and a choice row (View > Normal, a
      // Shadow preset) a radio row: `aria-checked` says whether it is on.
      const { checked, radio } = menuRowCheck(row, control, state.value, entry.value);
      const valueGlyph =
        entry.value !== undefined && entry.slot === 'insert.line' ? (
          <LineGlyph value={entry.value} />
        ) : undefined;
      return (
        <MenuRow
          {...common}
          {...(valueGlyph ? { glyph: valueGlyph } : {})}
          {...(shortcutKey ? { shortcut: translateKeyOrText(t, shortcutKey) } : {})}
          {...(checked !== undefined ? { checked } : {})}
          {...(radio ? { radio: true } : {})}
          slot={row.key}
          onSelect={() => {
            // A pointer press keeps the focus where it was (the slide's text input); a keyboard
            // user's focus is on the row, inside the editor scope. Either way the scope is found.
            const scope = document.activeElement?.closest('.pptx-editor') ?? null;
            close();
            runRow(editor.value, host, entry.slot, entry.value, surfaceInputIn(scope));
          }}
        />
      );
    };
  },
});

/** The row props a gallery submenu passes on to its `MenuSubmenu`. */
const submenuRowProps = {
  label: { type: String, required: true as const },
  disabled: { type: Boolean, default: false },
  reason: { type: String as PropType<string | null>, default: null },
  className: { type: String, default: undefined },
  /** A Material Symbol name from `Menu.Item icon`, in place of the registry icon. */
  icon: { type: String, default: undefined },
  paths: { type: Array as PropType<readonly string[] | null>, default: undefined },
  glyph: { type: Object as PropType<VNode>, default: undefined },
} as const;

/** The `MenuSubmenu` props of a gallery row: icon content, a named icon, or the paths. */
function submenuIcon(props: {
  readonly icon?: string | undefined;
  readonly paths?: readonly string[] | null | undefined;
  readonly glyph?: VNode | undefined;
}) {
  return props.glyph !== undefined
    ? { glyph: props.glyph }
    : props.icon !== undefined
      ? { icon: props.icon }
      : { paths: props.paths ?? null };
}

/**
 * Insert > Shape: one cascading grid per gallery category, drawn from the spec geometry. As in
 * React, a named `icon` is not shown on this row.
 */
const GallerySubmenu = defineComponent({
  name: 'PptxGallerySubmenu',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true as const },
    ...submenuRowProps,
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const host = inject(chromeActionHostKey, null);
    const close = useMenuClose();
    return () => (
      <MenuSubmenu
        label={props.label}
        disabled={props.disabled}
        reason={props.reason}
        {...(props.className ? { className: props.className } : {})}
        {...(props.glyph !== undefined ? { glyph: props.glyph } : { paths: props.paths ?? null })}
      >
        {{
          default: () =>
            SHAPE_GALLERY.map((category) => (
              <MenuSubmenu
                key={category.id}
                label={t(category.labelKey as TranslationKey)}
                glyph={<ShapeIcon preset={category.presets[0] ?? 'rect'} />}
                className="pptx-menu--shapes"
              >
                {{
                  default: () => (
                    <ShapeGrid
                      presets={category.presets}
                      onPick={(preset: string) => {
                        close();
                        runRow(editor.value, host, props.slot, preset);
                      }}
                    />
                  ),
                }}
              </MenuSubmenu>
            )),
        }}
      </MenuSubmenu>
    );
  },
});

/**
 * New Slide and Layout: the deck's layout gallery, each tile drawn by the engine. Layout marks
 * the current slide's layout.
 */
const LayoutSubmenu = defineComponent({
  name: 'PptxLayoutSubmenu',
  props: {
    slot: { type: String as PropType<'slide.add' | 'slide.layout'>, required: true as const },
    ...submenuRowProps,
  },
  setup(props) {
    const editor = usePptxEditor();
    const host = inject(chromeActionHostKey, null);
    const close = useMenuClose();
    const current = useEditorState((s) =>
      props.slot === 'slide.layout' ? (s.slides[s.currentSlide]?.layoutId ?? '') : ''
    );
    return () => (
      <MenuSubmenu
        label={props.label}
        disabled={props.disabled}
        reason={props.reason}
        {...(props.className ? { className: props.className } : {})}
        {...submenuIcon(props)}
      >
        {{
          default: () => (
            <LayoutGallery
              {...(props.slot === 'slide.layout' ? { value: current.value } : {})}
              onChange={(layoutId: string) => {
                close();
                runRow(editor.value, host, props.slot, layoutId);
              }}
            />
          ),
        }}
      </MenuSubmenu>
    );
  },
});

/**
 * The registry row a host row replaces in place: a host `MenuItem` there keeps the row's own
 * label, shortcut, and `direct` unless it sets its own. @internal
 */
export const rowDefaultsKey: InjectionKey<ChromeMenuItemEntry | null> = Symbol('pptxRowDefaults');

/** Provides the registry row a host row replaces, then renders the host row. */
const RowDefaults = defineComponent({
  name: 'PptxRowDefaults',
  props: { entry: { type: Object as PropType<ChromeMenuItemEntry>, required: true as const } },
  setup(props, { slots }) {
    provide(rowDefaultsKey, props.entry);
    return () => slots.default?.();
  },
});

/** Render a list of registry entries, with host overrides by row key. @internal */
export const MenuEntries = defineComponent({
  name: 'PptxMenuEntries',
  props: {
    entries: { type: Array as PropType<readonly ChromeMenuEntry[]>, required: true as const },
    /** Host rows that replace registry rows, by row key; `null` removes the row. */
    overrides: {
      type: Object as PropType<ReadonlyMap<string, (() => VNode | null) | null>>,
      default: undefined,
    },
  },
  setup(props) {
    const { t } = useTranslation();
    return () =>
      props.entries.map((entry, i) => {
        if (entry.kind === 'separator') return <MenuSeparator key={`sep-${i}`} />;
        if (entry.kind === 'submenu')
          return (
            <MenuSubmenu key={entry.id} label={t(entry.labelKey)} paths={entry.paths}>
              {{
                default: () => (
                  <MenuEntries
                    entries={entry.items}
                    {...(props.overrides ? { overrides: props.overrides } : {})}
                  />
                ),
              }}
            </MenuSubmenu>
          );
        const key = menuEntryKey(entry);
        if (props.overrides?.has(key)) {
          const render = props.overrides.get(key);
          return (
            <RowDefaults key={key} entry={entry}>
              {{ default: () => render?.() ?? null }}
            </RowDefaults>
          );
        }
        return <SlotMenuRow key={key} entry={entry} />;
      });
  },
});
