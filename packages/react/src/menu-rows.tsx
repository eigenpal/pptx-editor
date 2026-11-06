/** @jsxImportSource react */
// Menu rows over chrome slots, shared by `PptxEditor.Menu` and `PptxEditor.ContextMenu`, after
// DOCX Editor's `Menu.Item` and `ContextMenu.Slot` (docx-editor/packages/react/src/editor/menu/
// parts.tsx). A row takes its label, icon, shortcut, enabled state, and action from the slot's
// registry entry and core's row rules (`resolveChromeMenuRow`, `chromeRowState`,
// `runChromeRow`), so a menu row and its toolbar twin cannot drift. A slot that collects a
// value (Insert > Shape, Insert > Line, Insert > Table, New Slide, Layout) opens a submenu of its
// choices.

import { createContext, useContext, type ReactNode } from 'react';
import {
  SHAPE_GALLERY,
  chromeControlForSlot,
  chromeRowState,
  menuEntryKey,
  menuRowCheck,
  resolveChromeMenuRow,
  runChromeRow,
  sameChromeCommandState,
  surfaceInputIn,
  tidySeparators,
  translateKeyOrText,
  warnMenuPart,
  type ChromeCommandState,
  type ChromeMenuEntry,
  type ChromeMenuItemEntry,
  type ChromeSlotId,
  type EditorSnapshot,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { MenuCloseContext, MenuRow, MenuSeparator, MenuSubmenu } from './ContextMenu.tsx';
import { useEditorState, useOfferedControl, usePptxEditor, useTranslation } from './context.ts';
import { useChromeActionHost } from './chrome-actions.tsx';
import { LineGlyph, ShapeGrid, ShapeIcon } from './ShapeControls.tsx';
import { LayoutGallery } from './SlideControls.tsx';
import { MenuTableGrid } from './TableControls.tsx';

export { MenuCloseContext };

/** The close function of the enclosing menu, or a no-op. */
export function useMenuClose(): () => void {
  return useContext(MenuCloseContext) ?? noop;
}

function noop(): void {}

/** The host overrides of one registry row. @internal */
export interface SlotRowOverrides {
  readonly labelKey?: string;
  readonly shortcutKey?: string;
  readonly className?: string;
  /** Icon content in place of the registry icon. */
  readonly icon?: ReactNode;
}

/** A row's state, re-read when the editor changes. */
function useRowState(slot: ChromeSlotId, value: string | undefined): ChromeCommandState {
  const editor = usePptxEditor();
  // The host knows which `popups` entries are off: their rows show disabled with the reason.
  const host = useChromeActionHost();
  return useEditorState(
    (() => chromeRowState(editor, slot, value, host)) as unknown as (
      s: EditorSnapshot
    ) => ChromeCommandState,
    sameChromeCommandState as (a: ChromeCommandState, b: ChromeCommandState) => boolean
  );
}

/** The icon props of a row: exactly one of `icon`, `glyph`, or `paths`. */
type SlotRowIcon =
  | { readonly icon: ReactNode }
  | { readonly glyph: ReactNode }
  | { readonly paths: readonly string[] | null };

/** What every row kind takes from its registry row and the host overrides. */
type SlotRowCommon = {
  readonly label: string;
  readonly disabled: boolean;
  readonly reason: string | null;
  readonly className?: string;
} & SlotRowIcon;

/**
 * The one icon rule for every row kind, so no kind drops the host's icon. The host's `icon`
 * wins; `MenuRow` and `MenuSubmenu` read a string as a Material Symbol name and an element as
 * icon content. Insert > Line has no registry icon: its rows draw a line, as its options do.
 */
function slotRowIcon(
  entry: ChromeMenuItemEntry,
  paths: readonly string[] | null,
  override: ReactNode | undefined
): SlotRowIcon {
  if (override !== undefined) return { icon: override };
  if (entry.slot === 'insert.line' && (entry.value !== undefined || paths === null))
    return { glyph: <LineGlyph value={entry.value ?? 'line'} /> };
  return { paths };
}

/** One registry row: a slot, or a slot with a value. @internal */
export function SlotMenuRow({
  entry,
  overrides,
}: {
  entry: ChromeMenuItemEntry;
  overrides?: SlotRowOverrides;
}): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const host = useChromeActionHost();
  const close = useMenuClose();
  const state = useRowState(entry.slot, entry.value);
  const row = resolveChromeMenuRow(entry);
  const control = chromeControlForSlot(entry.slot);
  // The choices this build offers for the selection: a choice whose result does not paint
  // stays out, and a row that removes an effect the selection has stays in.
  // An option row (`entry.value`) draws no choices, so it reads none.
  const offeredControl = useOfferedControl(entry.slot, entry.value === undefined ? control : null);
  if (!row || !control) {
    warnMenuPart('Menu.Item', 'unknownSlot', entry.slot);
    return null;
  }
  const label = overrides?.labelKey
    ? translateKeyOrText(t, overrides.labelKey)
    : translateKeyOrText(t, row.labelKey, row.labelParams);
  const shortcutKey = overrides?.shortcutKey ?? row.shortcutKey;
  const reason = state.disabledReason ? translateKeyOrText(t, state.disabledReason) : null;
  // Every row kind (plain, toggle, submenu, gallery, value option) draws the same icon.
  const common: SlotRowCommon = {
    label,
    disabled: !state.isEnabled,
    reason,
    ...slotRowIcon(entry, row.paths, overrides?.icon),
    ...(overrides?.className ? { className: overrides.className } : {}),
  };
  // A slot that collects a value opens its choices, unless the row names one.
  if (entry.value === undefined && control.state.kind === 'value' && !entry.direct) {
    if (entry.slot === 'slide.add' || entry.slot === 'slide.layout')
      return <LayoutSubmenu slot={entry.slot} {...common} />;
    // Insert > Shape cascades its shape grid. Quick Styles is a gallery with no menu form yet,
    // so its row falls through to the disabled "pick in the toolbar" row below.
    if (entry.slot === 'insert.shape') return <GallerySubmenu slot={entry.slot} {...common} />;
    // Insert > Table: the size grid, as DOCX Editor's `Menu.TableGrid`.
    if (control.shape === 'tableGrid')
      return (
        <MenuSubmenu {...common}>
          <MenuTableGrid slot={entry.slot} label={label} autoFocus onInserted={close} />
        </MenuSubmenu>
      );
    const offered = offeredControl ?? control;
    // Shape Effects: one submenu for each effect kind, as its toolbar menu has.
    if (offered.optionGroups?.length && offered.options?.length)
      return (
        <MenuSubmenu {...common}>
          {offered.optionGroups.map((group) => (
            <MenuSubmenu key={group.id} label={t(group.labelKey)}>
              {offered
                .options!.filter((o) => o.group === group.id)
                .map((o) => (
                  <SlotMenuRow
                    key={o.value}
                    entry={{ kind: 'item', slot: entry.slot, value: o.value }}
                  />
                ))}
            </MenuSubmenu>
          ))}
        </MenuSubmenu>
      );
    if (offered.options?.length)
      return (
        <MenuSubmenu {...common}>
          {offered.options.map((o) => (
            <SlotMenuRow key={o.value} entry={{ kind: 'item', slot: entry.slot, value: o.value }} />
          ))}
        </MenuSubmenu>
      );
    // A picker with no menu form (a color, a font, a size) cannot run from a row.
    if (entry.slot !== 'text.link') {
      warnMenuPart('Menu.Item', 'valueRow', entry.slot);
      return <MenuRow {...common} disabled reason={t('menu.pickInToolbar')} slot={row.key} />;
    }
  }
  // A View setting (Smart guides) is a check row; a formatting toggle (Bold) shows as active,
  // as in a Format menu; an option of a choice slot is checked when in force.
  // A two-state row (Bold, Ruler) is a checkbox row and a choice row (View > Normal, a Shadow
  // preset) a radio row: `aria-checked` says whether it is on, as in DOCX Editor's menus.
  const { checked, radio } = menuRowCheck(row, control, state, entry.value);
  return (
    <MenuRow
      {...common}
      {...(shortcutKey ? { shortcut: translateKeyOrText(t, shortcutKey) } : {})}
      {...(checked !== undefined ? { checked } : {})}
      {...(radio ? { radio: true } : {})}
      slot={row.key}
      onSelect={() => {
        // A pointer press keeps the focus where it was (the slide's text input); a keyboard
        // user's focus is on the row, inside the editor scope. Either way the scope is found.
        const scope = document.activeElement?.closest('.pptx-editor') ?? null;
        close();
        runChromeRow(editor, host, entry.slot, entry.value, surfaceInputIn(scope));
      }}
    />
  );
}

/** Insert > Shape: one cascading grid per gallery category, drawn from the spec geometry. */
function GallerySubmenu({ slot, ...common }: { slot: ChromeSlotId } & SlotRowCommon): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const host = useChromeActionHost();
  const close = useMenuClose();
  return (
    <MenuSubmenu {...common}>
      {SHAPE_GALLERY.map((category) => (
        <MenuSubmenu
          key={category.id}
          label={t(category.labelKey as TranslationKey)}
          glyph={<ShapeIcon preset={category.presets[0] ?? 'rect'} />}
          className="pptx-menu--shapes"
        >
          <ShapeGrid
            presets={category.presets}
            onPick={(preset) => {
              close();
              runChromeRow(editor, host, slot, preset);
            }}
          />
        </MenuSubmenu>
      ))}
    </MenuSubmenu>
  );
}

/**
 * New Slide and Layout: the deck's layout gallery, each tile drawn by the engine. Layout marks
 * the current slide's layout.
 */
function LayoutSubmenu({
  slot,
  ...common
}: { slot: 'slide.add' | 'slide.layout' } & SlotRowCommon): ReactNode {
  const editor = usePptxEditor();
  const host = useChromeActionHost();
  const close = useMenuClose();
  const current = useEditorState((s) =>
    slot === 'slide.layout' ? (s.slides[s.currentSlide]?.layoutId ?? '') : ''
  );
  return (
    <MenuSubmenu {...common}>
      <LayoutGallery
        {...(slot === 'slide.layout' ? { value: current } : {})}
        onChange={(layoutId) => {
          close();
          runChromeRow(editor, host, slot, layoutId);
        }}
      />
    </MenuSubmenu>
  );
}

/** Render a list of registry entries, with host overrides by row key. @internal */
export function MenuEntries({
  entries,
  overrides,
}: {
  entries: readonly ChromeMenuEntry[];
  /** Host rows that replace registry rows, by row key; `null` removes the row. */
  overrides?: ReadonlyMap<string, ReactNode | null>;
}): ReactNode {
  const { t } = useTranslation();
  return entries.map((entry, i) => {
    if (entry.kind === 'separator') return <MenuSeparator key={`sep-${i}`} />;
    if (entry.kind === 'submenu')
      return (
        <MenuSubmenu key={entry.id} label={t(entry.labelKey)} paths={entry.paths}>
          <MenuEntries entries={entry.items} {...(overrides ? { overrides } : {})} />
        </MenuSubmenu>
      );
    const key = menuEntryKey(entry);
    if (overrides?.has(key))
      return (
        <RowDefaultsContext.Provider key={key} value={entry}>
          {overrides.get(key)}
        </RowDefaultsContext.Provider>
      );
    return <SlotMenuRow key={key} entry={entry} />;
  });
}

/**
 * The registry row a host row replaces in place: a `Menu.Item` there keeps the row's own label,
 * shortcut, and `direct` unless it sets its own. @internal
 */
export const RowDefaultsContext = createContext<ChromeMenuItemEntry | null>(null);

/**
 * Drop separators a hidden row left behind: at the start, at the end, and two in a row. The rule
 * lives in core (`chrome-forms.ts`), so the Vue menus drop the same ones. @internal
 */
export { tidySeparators };
