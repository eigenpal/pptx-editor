/** @jsxImportSource react */
// The font family picker, after DOCX Editor's `FontFamily` toolbar compound
// (docx-editor/packages/react/src/editor/toolbar/FontFamily.tsx): a root that holds the picker
// state in a part-level context, and the `Trigger`, `Content`, and `Item` parts that read it.
// With no children the root renders a trigger and a search field over a list box of groups (the
// slide's theme fonts, labeled Headings and Body; the fonts in this presentation; all fonts),
// each name drawn in its own face, the current family marked and scrolled into view when the
// list opens. With children, a host composes the parts itself. Core's `font-picker.ts` builds
// the groups and rows, and the Vue twin (`packages/vue/src/FontFamilyControl.tsx`) renders the
// same markup. DOCX Editor commit adf0e178.

import {
  createContext,
  useContext,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
  type RefObject,
} from 'react';
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
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';
import type { ToolbarSlotPartProps } from './toolbar-parts.tsx';

/** The font the text uses: its family name and the theme font it comes from, if any. */
interface CurrentFont {
  readonly family: string | null;
  readonly theme: 'major' | 'minor' | null;
}

/** The picker state the parts share. */
interface FontFamilyState {
  readonly slot: ChromeSlotId;
  readonly label: string;
  readonly open: boolean;
  readonly setOpen: (next: boolean) => void;
  readonly query: string;
  readonly setQuery: (query: string) => void;
  readonly groups: readonly FontPickerGroup[];
  readonly current: CurrentFont;
  /** Apply a row's value (a theme font reference or a family) and close the list. */
  readonly pick: (value: string) => void;
  /** The row a host's `Item value` stands for, or null for a value that is not a font name. */
  readonly rowFor: (value: string) => FontPickerRow | null;
  readonly anchor: RefObject<HTMLButtonElement | null>;
  readonly list: RefObject<HTMLDivElement | null>;
  readonly trigger: ReturnType<typeof useTriggerAttributes>;
  readonly slotState: ReturnType<typeof useSlot>;
  readonly off: ReturnType<typeof useAriaDisabled>;
}

const FontFamilyContext = createContext<FontFamilyState | null>(null);

/** The shared state, or null (with a warning) outside `Toolbar.FontFamily`. */
function useFontFamilyState(part: string): FontFamilyState | null {
  const state = useContext(FontFamilyContext);
  if (!state) warnPartOutsideCompound(`Toolbar.FontFamily.${part}`, 'Toolbar.FontFamily');
  return state;
}

const join = (...names: (string | undefined)[]): string => names.filter(Boolean).join(' ');

/** Props of `Toolbar.FontFamily.Trigger` and `Toolbar.FontFamily.Content`. @public */
export interface FontFamilyPartProps {
  /** A class name added to the part's element. */
  readonly className?: string;
  /** Content in place of the part's default content. */
  readonly children?: ReactNode;
}

/** Props of `Toolbar.FontFamily.Item`. @public */
export interface FontFamilyItemProps extends FontFamilyPartProps {
  /**
   * What the item applies: a family name, or a theme font reference (`+mj-lt` for the headings
   * font, `+mn-lt` for the body font), so the text follows a later theme change.
   */
  readonly value: string;
}

/**
 * The button that opens the list: the current family, or the label when the selection has no
 * one family. Children replace its content.
 */
function FontFamilyTrigger({ className, children }: FontFamilyPartProps): ReactNode {
  const state = useFontFamilyState('Trigger');
  if (!state) return null;
  const { cmd, label, tip } = state.slotState;
  return (
    <>
      <button
        ref={state.anchor}
        type="button"
        {...state.trigger}
        className={join(state.trigger.className, className)}
        data-slot={state.slot}
        aria-label={label}
        aria-haspopup="dialog"
        aria-expanded={state.open}
        {...tip}
        {...state.off.attrs}
        onMouseDown={keepFocus}
        onClick={() => {
          state.setQuery('');
          state.setOpen(!state.open);
        }}
      >
        {children ?? (
          <>
            <span className="pptx-toolbar__family-label">{cmd.value ?? label}</span>
            <Icon name="expand_more" className="pptx-icon pptx-toolbar__caret" />
          </>
        )}
      </button>
      {state.off.note}
    </>
  );
}

/** One family row: an `option` with the name in its own face and the theme role, if any. */
function FontRow({
  row,
  className,
  children,
}: {
  row: FontPickerRow;
  className?: string | undefined;
  children?: ReactNode;
}): ReactNode {
  const { t } = useTranslation();
  const state = useContext(FontFamilyContext);
  if (!state) return null;
  const current = !row.disabled && isCurrentFontRow(row, state.current.family, state.current.theme);
  // The face is a quoted, validated family name (core's `fontFaceValue`), set as a custom
  // property through a style object, never a CSS string.
  const face = row.face ? ({ '--pptx-font-face': row.face } as CSSProperties) : undefined;
  return (
    <button
      type="button"
      role="option"
      aria-selected={current}
      tabIndex={-1}
      disabled={row.disabled}
      className={join('pptx-menu__item pptx-font-picker__option', className)}
      data-family={row.family || undefined}
      data-theme={row.theme ?? undefined}
      onClick={() => state.pick(row.value)}
    >
      {children ?? (
        <>
          {row.family ? (
            <span className="pptx-font-picker__face" style={face}>
              {row.family}
            </span>
          ) : null}
          {row.roleKey ? <span className="pptx-font-picker__role">{t(row.roleKey)}</span> : null}
        </>
      )}
      {current ? <Icon name="check" className="pptx-icon pptx-menu__check" /> : null}
    </button>
  );
}

/**
 * One option of the list: it applies `value` and carries the check mark while the text uses
 * that font. Children replace the name and role. A theme font that the slide's theme leaves
 * empty shows its role, disabled. A value that is not a font name renders nothing, with a
 * development warning.
 */
function FontFamilyItem({ value, className, children }: FontFamilyItemProps): ReactNode {
  const state = useFontFamilyState('Item');
  if (!state) return null;
  const row = state.rowFor(value);
  if (!row) {
    warnInvalidFontItem('Toolbar.FontFamily.Item', value);
    return null;
  }
  return (
    <FontRow row={row} className={className}>
      {children}
    </FontRow>
  );
}

/** One labeled group of the list box. */
function FontGroup({ group }: { group: FontPickerGroup }): ReactNode {
  const { t } = useTranslation();
  const id = useId();
  return (
    <div className="pptx-font-picker__group" role="group" aria-labelledby={id}>
      <div id={id} role="presentation" className="pptx-menu__heading pptx-font-picker__heading">
        {t(group.labelKey)}
      </div>
      {group.rows.map((row) => (
        <FontRow key={row.value} row={row} />
      ))}
    </div>
  );
}

/**
 * The list, while it is open. With no children: a search field over the theme, presentation,
 * and all-fonts groups. With children: a list box of the host's `Item` parts.
 */
function FontFamilyContent({ className, children }: FontFamilyPartProps): ReactNode {
  const state = useFontFamilyState('Content');
  const { t } = useTranslation();
  if (!state || !state.open) return null;
  const { groups } = state;
  return (
    <Popover anchor={state.anchor.current} role="dialog" field onClose={() => state.setOpen(false)}>
      <div className={join('pptx-font-picker', className)}>
        {children === undefined ? (
          <input
            className="pptx-font-picker__search"
            type="search"
            placeholder={t('toolbar.fontSearch')}
            aria-label={t('toolbar.fontSearch')}
            value={state.query}
            autoFocus
            onChange={(e) => state.setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key !== 'Enter') return;
              // The key stops here, so it does not press the trigger and open the list again.
              e.preventDefault();
              const first = firstFontPickerFamily(groups);
              if (first) state.pick(first);
            }}
          />
        ) : null}
        <div
          ref={state.list}
          className="pptx-font-picker__list"
          role="listbox"
          aria-label={state.label}
        >
          {children !== undefined ? (
            children
          ) : (
            <>
              {groups.map((g) => (
                <FontGroup key={g.id} group={g} />
              ))}
              {groups.length === 0 ? (
                <p className="pptx-font-picker__empty">{t('toolbar.fontNone')}</p>
              ) : null}
            </>
          )}
        </div>
      </div>
    </Popover>
  );
}

/** Props of the font family root, as the toolbar renders it. @internal */
export interface FontFamilyControlProps {
  readonly slot: ChromeSlotId;
  readonly control: ChromeControl;
  /** `Trigger`, `Content`, and `Item` parts in place of the default trigger and list. */
  readonly children?: ReactNode;
}

/**
 * The font family picker: the toolbar's part for `font.family`, and the root of the
 * `Toolbar.FontFamily` compound. @internal
 */
export function FontFamilyControl({ slot, control, children }: FontFamilyControlProps) {
  const slotState = useSlot(slot, control);
  const { cmd, reason } = slotState;
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const editor = usePptxEditor();
  const anchor = useRef<HTMLButtonElement>(null);
  const list = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  // The groups change with the deck (embedded and host fonts) and with the slide's theme.
  useEditorState((s) => s.revision);
  const slide = useEditorState((s) => s.currentSlide);
  const groups = open ? fontPickerGroups(editor, slide, query) : [];
  // The check mark follows what the text uses: a theme font marks its theme row only.
  const theme = useEditorState((s) => s.format?.fontFamilyTheme ?? null);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__field pptx-toolbar__field--family', {
    disabled: !cmd.isEnabled,
  });
  // Once, when the list opens: the current family comes into view.
  useLayoutEffect(() => {
    if (open) revealCurrentFont(list.current);
  }, [open]);
  // After the popover took focus: a list with no search field focuses its current option, so
  // the arrow keys reach the host's `Item` parts.
  useEffect(() => {
    if (open) focusFontOption(list.current);
  }, [open]);
  const state: FontFamilyState = {
    slot,
    label: slotState.label,
    open,
    setOpen,
    query,
    setQuery,
    groups,
    current: { family: typeof cmd.value === 'string' ? cmd.value : null, theme },
    // A theme row applies its theme font reference (`+mj-lt`, `+mn-lt`), any other row its
    // family.
    pick: (value) => {
      cmd.execute(value);
      focusText(editor, anchor.current);
      setOpen(false);
    },
    rowFor: (value) => fontPickerRow(editor, slide, value),
    anchor,
    list,
    trigger,
    slotState,
    off,
  };
  return (
    <FontFamilyContext.Provider value={state}>
      {children ?? (
        <>
          <FontFamilyTrigger />
          <FontFamilyContent />
        </>
      )}
    </FontFamilyContext.Provider>
  );
}

/**
 * Props of `Toolbar.FontFamily`: a picker part's `className` and `hidden`, and children that
 * compose its `Trigger`, `Content`, and `Item` parts. @public
 */
export interface FontFamilyProps extends ToolbarSlotPartProps {
  /** `Toolbar.FontFamily.Trigger` and `Toolbar.FontFamily.Content`, in place of the defaults. */
  readonly children?: ReactNode;
}

/**
 * `Toolbar.FontFamily`: the `font.family` picker, as DOCX Editor's `FontFamily` compound. With
 * no children it renders the default trigger and list; with children, compose the `Trigger`,
 * `Content`, and `Item` parts around the shared state. @public
 */
export interface FontFamilyNamespace {
  /** Changes, hides, or recomposes the font family picker. */
  (props: FontFamilyProps): ReactNode;
  /** The slot the part stands for. The toolbar reads it to place the part. */
  readonly pptxSlot: ChromeSlotId;
  /** The button that opens the list. */
  readonly Trigger: (props: FontFamilyPartProps) => ReactNode;
  /** The list, while it is open. */
  readonly Content: (props: FontFamilyPartProps) => ReactNode;
  /** One option of the list. */
  readonly Item: (props: FontFamilyItemProps) => ReactNode;
}

/** The sub-parts of `Toolbar.FontFamily`, as DOCX Editor's `FontFamily` statics. @internal */
export const FONT_FAMILY_PARTS = {
  Trigger: FontFamilyTrigger,
  Content: FontFamilyContent,
  Item: FontFamilyItem,
} as const;
