/** @jsxImportSource react */
// The floating formatting toolbar. Its arrangement derives FROM `CHROME_GROUPS`; it is never
// hand-listed, so a slot added to the registry appears here without editing this file.

import type { ChromeTranslate } from '@pptx-editor.dev/core/editor';
import {
  Children,
  Fragment,
  isValidElement,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import {
  colorMenuSelection,
  arrangeToolbarGroups,
  chromeOfferedControl,
  createToolbarFitMemory,
  observeToolbarFit,
  requestDialogSection,
  runChromeRow,
  surfaceInputIn,
  chromeControlForSlot,
  chromeSlotIsToggle,
  isToolbarBarSlot,
  TOOLBAR_NAMED_PARTS,
  toolbarCollapsedIds,
  toolbarColorResetKey,
  toolbarDropdownRow,
  toolbarGroupLabel,
  toolbarGroupPriority,
  toolbarPartKind,
  toolbarRowSection,
  toolbarSplitDefaultColor,
  toolbarContextOf,
  toolbarGroupsFor,
  translateKeyOrText,
  warnDuplicateSlot,
  warnSlotOutsideBar,
  warnUnknownSlot,
  type ChromeControl,
  type ChromeGroup,
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
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useSlot, type SlotPartProps } from './slot.ts';
import { ColorSwatches } from './ColorSwatches.tsx';
import { FillControl, LineMenuControl, ShapeGalleryControl } from './ShapeControls.tsx';
import { EffectMenuControl, ShapeStyleGalleryControl } from './ShapeFormatControls.tsx';
import {
  LayoutControl,
  NewSlideControl,
  SectionAddControl,
  SectionRenameControl,
} from './SlideControls.tsx';
import { requestFindBar } from './FindBar.tsx';
import { ImageControl, LinkControl } from './toolbar-file-controls.tsx';
import { TableGridControl, TableStyleControl } from './TableControls.tsx';
import { CorrectionsControl } from './PictureControls.tsx';
import { usePopupEntry } from './popups.tsx';
import { useChromeActionHost } from './chrome-actions.tsx';
import {
  SlotControl,
  namedPartSlotProps,
  slotPartNames,
  type ToolbarGroupProps,
  type ToolbarOverrides,
  type ToolbarSlotProps,
} from './toolbar-parts.tsx';
import { useTriggerAttributes } from './slot-override.ts';
import { SizeControl } from './toolbar-size.tsx';
import { FontFamilyControl } from './FontFamilyControl.tsx';
import { useDialogOpener } from './popups.tsx';
import { useRovingToolbar } from './toolbar-measure.ts';
import { useNotesTarget } from './notes-target.ts';

/** The color a split color button applies before the user picks one (core's rule). */
function useDefaultColor(slot: ChromeSlotId): ToolbarPickedColor {
  const text1 = useEditorState((s) => s.themeColors[0]?.[1]?.color);
  return toolbarSplitDefaultColor(slot, text1);
}

function IconControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  return (
    <ToolButton
      slot={slot}
      label={label}
      paths={control.paths}
      onClick={() => cmd.execute()}
      disabled={!cmd.isEnabled}
      disabledReason={reason}
      // The registry says which controls are toggles; `aria-pressed` goes on those only.
      pressed={chromeSlotIsToggle(slot) ? cmd.isActive : undefined}
    />
  );
}

function DropdownControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const openDialog = useDialogOpener();
  const host = useChromeActionHost();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const current = control.options?.find((o) => o.value === cmd.value);
  const offered = chromeOfferedControl(slot, control);
  // A dropdown with no icon is the arrow half of a split button (Bullets and Numbering).
  const icon = current?.paths ?? control.paths;
  const trigger = useTriggerAttributes(
    slot,
    `pptx-toolbar__button ${icon ? 'pptx-toolbar__button--wide' : 'pptx-toolbar__button--caret'}`,
    { disabled: !cmd.isEnabled }
  );
  return (
    <>
      <button
        ref={anchor}
        type="button"
        {...trigger}
        data-slot={slot}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        {icon ? <Icon paths={icon} /> : null}
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          {/* The choices this build offers: a choice whose result does not paint stays out. */}
          {offered.options?.map((o) => {
            // A dialog row ("Line spacing options…") opens its registry dialog; core's rule says
            // which rows are checked or off, with the reason.
            const row = toolbarDropdownRow(editor, host, slot, control, o, cmd.value);
            const dialog = row.dialog;
            return (
              <Fragment key={o.value}>
                {dialog ? <div className="pptx-menu__sep" role="separator" /> : null}
                <MenuItem
                  {...(row.checked !== undefined ? { checked: row.checked } : {})}
                  {...(row.refused
                    ? {
                        disabled: true,
                        ...(row.reason ? { reason: translateKeyOrText(t, row.reason) } : {}),
                      }
                    : {})}
                  onSelect={() => {
                    // A pick gives the caret back to the slide text, so the
                    // next key (typing, Cmd+Z) reaches the slide.
                    if (!dialog) {
                      cmd.execute(o.value);
                      focusText(editor, anchor.current);
                      setOpen(false);
                      return;
                    }
                    setOpen(false);
                    // Focus returns to the text when the dialog closes.
                    const text = anchor.current
                      ?.closest('.pptx-editor')
                      ?.querySelector<HTMLElement>('.pptx-surface__input');
                    const section = toolbarRowSection(slot);
                    if (editor && section) requestDialogSection(editor, dialog, section);
                    openDialog(dialog, text ?? anchor.current);
                  }}
                >
                  {o.paths ? (
                    <Icon paths={o.paths} />
                  ) : o.glyph || offered.options?.some((x) => x.glyph) ? (
                    // A row with no preview keeps the column, so the labels line up.
                    <span className="pptx-menu__glyph" aria-hidden="true">
                      {o.glyph}
                    </span>
                  ) : null}
                  {t(o.labelKey, o.labelParams)}
                </MenuItem>
              </Fragment>
            );
          })}
        </Popover>
      ) : null}
    </>
  );
}

function ColorControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const fallback = useDefaultColor(slot);
  // The current color of the selection, marked in the menu.
  const selected = useEditorState((s) => colorMenuSelection(slot, s));
  const [picked, setPicked] = useState<ToolbarPickedColor | null>(null);
  const last = picked ?? fallback;
  // Shapes can be transparent; text returns to its inherited (automatic) color.
  const resetLabel = t(toolbarColorResetKey(slot));
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  return (
    <span className="pptx-toolbar__group">
      <ToolButton
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!cmd.isEnabled}
        disabledReason={reason}
        asChild={false}
        onClick={() => cmd.execute(last.choice)}
      >
        <span className="pptx-toolbar__swatch" style={{ background: last.color }} />
      </ToolButton>
      <button
        ref={anchor}
        type="button"
        className="pptx-toolbar__button pptx-toolbar__step"
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          <ColorSwatches
            selected={selected}
            onPick={(choice, color) => {
              setPicked({ choice, color });
              cmd.execute(choice);
              focusText(editor, anchor.current);
              setOpen(false);
            }}
          />
          <div className="pptx-menu__sep" />
          <MenuItem
            onSelect={() => (cmd.execute(null), focusText(editor, anchor.current), setOpen(false))}
          >
            {resetLabel}
          </MenuItem>
        </Popover>
      ) : null}
    </span>
  );
}

/**
 * The link control: asks `PptxEditor.Root` for the link popup (`popups.hyperlink`, or the
 * packaged `HyperlinkPopup`) under this button. Ctrl/Cmd+K and Insert > Link ask the same
 * root, so the popup opens with or without this control.
 */
type SlotPart = (props: SlotPartProps) => ReactNode;

/**
 * Slots whose control is more than its registry shape: a searchable font list, a link editor,
 * a fill with gradients, and line menus that draw each option. As in DOCX Editor's
 * `SHAPED_PARTS`, this picks a renderer; it never decides which slots the toolbar shows.
 */
const SHAPED_PARTS: Partial<Record<ChromeSlotId, SlotPart>> = {
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

/**
 * The renderer for each part kind (core's `toolbarPartKind`): a registry shape, or an adapter
 * part for a control that loads a file, opens a panel, or runs an adapter action.
 */
const KIND_PARTS: Record<ToolbarPartKind, SlotPart> = {
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

/** The props of `ToolbarControl`. @public */
export interface ToolbarControlProps {
  /** The id of the group the control belongs to, such as `text`. */
  group: string;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/** One control, by registry shape. @public */
export function ToolbarControl({ group, control }: ToolbarControlProps) {
  const slot = `${group}.${control.id}` as ChromeSlotId;
  const Part = SHAPED_PARTS[slot] ?? KIND_PARTS[toolbarPartKind(control)];
  return <Part slot={slot} control={control} />;
}

/** A control that opens a panel the adapter owns: the find bar, to find or to replace. */
function PanelControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const off = usePopupEntry('find') === false;
  return (
    <ToolButton
      slot={slot}
      label={label}
      paths={control.paths}
      disabled={off || !cmd.isEnabled}
      disabledReason={off ? t('find.unavailable') : reason}
      onClick={() => requestFindBar(editor, slot === 'find.replace')}
    />
  );
}

/**
 * A control the adapter runs itself: a registry dialog (Page setup, Format background), the
 * clipboard, or the slide show. Core's `runChromeRow` picks the action.
 */
function ActionControl({ slot, control }: SlotPartProps) {
  const { cmd, label, reason } = useSlot(slot, control);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const host = useChromeActionHost();
  // A `popups` entry of `false` turns the button off. A slot that also takes a value
  // (`image.altText`) keeps its command enabled, so the button asks the host here.
  const off = host.unavailable?.(slot) ?? null;
  const ref = useRef<HTMLSpanElement>(null);
  return (
    <span ref={ref} className="pptx-toolbar__group">
      <ToolButton
        hasPopup={control.state.kind === 'dialog' ? 'dialog' : undefined}
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!!off || !cmd.isEnabled}
        disabledReason={off ? translateKeyOrText(t, off) : reason}
        onClick={() =>
          runChromeRow(
            editor,
            host,
            slot,
            undefined,
            surfaceInputIn(ref.current?.closest('.pptx-editor'))
          )
        }
      />
    </span>
  );
}

/**
 * One registered control. Inside a `Toolbar` with the preset on, it declares an override;
 * anywhere else (`preset={false}`, a group) it renders the control.
 *
 * @example
 * ```tsx
 * <PptxEditor.Toolbar>
 *   <PptxEditor.Toolbar.Slot slotId="text.bold" icon={<BoldIcon />} className="acme-bold" />
 *   <PptxEditor.Toolbar.Slot slotId="history.undo" asChild>
 *     <AcmeButton>Undo</AcmeButton>
 *   </PptxEditor.Toolbar.Slot>
 *   <PptxEditor.Toolbar.Slot slotId="text.strike" hidden />
 * </PptxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function ToolbarSlot(props: ToolbarSlotProps): ReactNode {
  const control = chromeControlForSlot(props.slotId);
  if (!control) {
    warnUnknownSlot(props.slotId);
    return null;
  }
  const group = props.slotId.slice(0, props.slotId.indexOf('.'));
  return (
    <SlotControl props={props} render={() => <ToolbarControl group={group} control={control} />} />
  );
}

/**
 * A group of controls. Inside a `Toolbar` with the preset on, a built-in id changes that group
 * and any other id adds a host group; with `preset={false}`, it renders its children. A host
 * group's accessible name is its `label`, else its translated `labelKey`, else its `id`.
 *
 * @example
 * ```tsx
 * <PptxEditor.Toolbar>
 *   <PptxEditor.Toolbar.Group id="acme" label="Acme">
 *     <button type="button" onClick={askLegal}>Ask legal</button>
 *   </PptxEditor.Toolbar.Group>
 *   <PptxEditor.Toolbar.Group id="find" hidden />
 * </PptxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function ToolbarGroup(props: ToolbarGroupProps): ReactNode {
  const { t } = useTranslation();
  if (props.hidden) return null;
  return (
    <div
      className={['pptx-toolbar__group', props.className].filter(Boolean).join(' ')}
      role="group"
      aria-label={toolbarGroupLabel(props, (key) => translateKeyOrText(t, key))}
      data-group={props.id}
    >
      {props.children}
    </div>
  );
}

/** The slot a named part (`Toolbar.Bold`) stands for, read from its static. */
const namedPartSlot = (type: unknown): ChromeSlotId | undefined =>
  (type as { pptxSlot?: ChromeSlotId } | null)?.pptxSlot;

/** A named part's name by its slot, for warning text: `text.bold` → `Toolbar.Bold`. */
const NAMED_PART_BY_SLOT = new Map<string, string>(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => [slot, `Toolbar.${name}`])
);

/**
 * Read `Toolbar.Slot`, named part, and `Toolbar.Group` children as overrides of the preset.
 * Only direct children declare overrides. Other children (a loose `Toolbar.Action`, a host
 * element) go after the groups and never collapse, as in DOCX Editor.
 */
function collectToolbarOverrides(
  children: ReactNode,
  inBar: (slot: string) => boolean = isToolbarBarSlot
): ToolbarOverrides {
  const slots = new Map<string, ToolbarSlotProps>();
  const groups = new Map<string, ToolbarGroupProps>();
  const extras: ReactNode[] = [];
  // Fragments open (`{readOnly && <>…</>}`), and empty children (`''`, `false`) add nothing.
  for (const child of flattenChildren(children)) {
    if (child === '' || child === null || child === undefined || typeof child === 'boolean')
      continue;
    if (!isValidElement(child)) {
      extras.push(child);
      continue;
    }
    const named = namedPartSlot(child.type);
    if (child.type === ToolbarSlot || named) {
      const props = (
        named
          ? namedPartSlotProps(
              named,
              child.props as Record<string, unknown>,
              NAMED_PART_BY_SLOT.get(named)
            )
          : child.props
      ) as ToolbarSlotProps;
      const parts = slotPartNames(props);
      if (!chromeControlForSlot(props.slotId)) warnUnknownSlot(props.slotId, parts);
      else if (!inBar(props.slotId)) {
        // A registered control the preset bar does not draw (`Toolbar.SmartGuides`) goes after
        // the groups, as DOCX Editor appends it; a `Toolbar.Slot` says so (a named part does
        // not, as in DOCX Editor).
        // A hidden slot draws nothing, so it gets no warning.
        if (!named && !props.hidden) warnSlotOutsideBar(props.slotId, parts);
        if (!props.hidden) extras.push(child);
        continue;
      }
      if (slots.has(props.slotId)) warnDuplicateSlot(props.slotId, parts);
      slots.set(props.slotId, props);
    } else if (child.type === ToolbarGroup) {
      const props = child.props as ToolbarGroupProps;
      groups.set(props.id, props);
    } else extras.push(child);
  }
  return { slots, groups, extras };
}

/** Children with fragments opened, depth first. */
function flattenChildren(nodes: ReactNode): ReactNode[] {
  const out: ReactNode[] = [];
  Children.forEach(nodes, (node) => {
    if (isValidElement<{ children?: ReactNode }>(node) && node.type === Fragment)
      out.push(...flattenChildren(node.props.children));
    else out.push(node);
  });
  return out;
}

/** Whether `nodes` renders anything: not only `null`, `false`, `''`, or empty fragments. */
function hasContent(nodes: ReactNode): boolean {
  return flattenChildren(nodes).some(
    (node) => node !== '' && node !== null && node !== undefined && typeof node !== 'boolean'
  );
}

/** Props for `PptxEditor.Toolbar`. @public */
export interface ToolbarProps {
  /**
   * Your translation function for this compound and its parts, as DOCX Editor's `t`. Without
   * it, the root's `locale` gives the labels.
   */
  readonly t?: ChromeTranslate;
  /** A class name added to the bar. */
  readonly className?: string;
  /**
   * `true` (default): the registry's groups, changed by `Toolbar.Slot` and `Toolbar.Group`
   * children. `false`: only the children, in the order given.
   */
  readonly preset?: boolean;
  /**
   * `true` (default): groups that do not fit the width move into a More menu, lowest priority
   * first. `false`: the bar wraps to more rows instead. A `preset={false}` bar never measures
   * itself, so it always wraps, as in DOCX Editor.
   */
  readonly overflow?: boolean;
  /**
   * `true` (default): the bar shows only the groups that act on the selection (text, a shape,
   * a picture, a table, or nothing). `false`: every bar group shows for every selection,
   * disabled with the engine's reason where it cannot run, and the picture or table group
   * joins after them. Hiding a group never turns off its commands.
   */
  readonly contextual?: boolean;
  /** `Toolbar.Slot` and `Toolbar.Group` overrides, or the whole bar when `preset` is false. */
  readonly children?: ReactNode;
}

/** One group on the bar: built-in or host, already resolved against the overrides. */
interface BarItem {
  readonly id: string;
  readonly label: string | undefined;
  readonly content: ReactNode;
  readonly priority: number;
  readonly pinned: boolean;
  readonly className: string | undefined;
  /** The host's content for the group in the More menu. */
  readonly overflowContent: (() => ReactNode) | undefined;
  /**
   * What a built-in group shows in the More menu when it has `overflowContent`: its own
   * controls, then that content in place of the host's children, as DOCX Editor's
   * `hostPanelContent`. A host group's `overflowContent` replaces its controls.
   */
  readonly menuContent?: ReactNode;
}

/**
 * The toolbar. With the preset, its arrangement derives from `CHROME_GROUPS` (the
 * non-contextual groups, and the `image` group while every selected element is a picture); groups that do not fit the width move, in order from the end, into
 * a More menu. `Toolbar.Slot` and `Toolbar.Group` children change the preset; with
 * `preset={false}`, the children are the whole bar.
 *
 * @example
 * ```tsx
 * <PptxEditor.Toolbar preset={false}>
 *   <PptxEditor.Toolbar.Slot slotId="history.undo" />
 *   <PptxEditor.Toolbar.Group id="formatting" label="Formatting">
 *     <PptxEditor.Toolbar.Slot slotId="text.bold" />
 *     <PptxEditor.Toolbar.Slot slotId="text.italic" />
 *   </PptxEditor.Toolbar.Group>
 * </PptxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function Toolbar({
  className,
  preset = true,
  overflow = true,
  contextual = true,
  children,
}: ToolbarProps) {
  const { t } = useTranslation();
  // The bar shows the groups that act on the selection: text, a shape, a picture, a table, or
  // nothing (`toolbarContextOf`).
  const selection = useEditorState(toolbarContextOf);
  // A focused speaker notes field takes the text formatting (`NOTES_SLOTS`).
  const notes = useNotesTarget(usePptxEditor());
  const context = notes ? 'notes' : selection;
  if (!preset) return <HandBar className={className}>{children}</HandBar>;
  return (
    <PresetBar
      groups={toolbarGroupsFor(context, { contextual })}
      label={t('toolbar.ariaLabel')}
      className={className}
      overflow={overflow}
      inBar={(slot) => isToolbarBarSlot(slot, { contextual })}
    >
      {children}
    </PresetBar>
  );
}

/**
 * A preset bar over `groups`: the registry's controls with the host's `Toolbar.Slot`, named
 * part, and `Toolbar.Group` overrides, collapsed into a More menu by priority. The toolbar
 * draws the formatting groups; each ribbon tab draws its own groups. `inBar` says which slots
 * the bar draws (default: the slots of `groups`); another slot goes after the groups. @internal
 */
export function PresetBar({
  groups: barGroups,
  label,
  className,
  overflow,
  inBar,
  children,
}: {
  groups: readonly ChromeGroup[];
  label: string;
  className?: string | undefined;
  overflow: boolean;
  inBar?: (slot: string) => boolean;
  children?: ReactNode;
}) {
  const { t } = useTranslation();
  const own = new Set(barGroups.flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`)));
  const overrides = collectToolbarOverrides(children, inBar ?? ((slot) => own.has(slot)));
  const items: BarItem[] = [];
  for (const entry of arrangeToolbarGroups(overrides.groups, barGroups)) {
    const override = entry.override;
    const id = entry.kind === 'built-in' ? entry.group.id : entry.override.id;
    const placement = {
      priority: toolbarGroupPriority(id, override?.priority, barGroups),
      pinned: override?.pinned === true,
      className: override?.className,
      overflowContent: override?.overflowContent,
    };
    if (entry.kind === 'host') {
      // A group with nothing to render leaves the bar, as in DOCX Editor (`{ok && <Action/>}`).
      if (!hasContent(entry.override.children)) continue;
      items.push({
        ...placement,
        id: entry.override.id,
        label: toolbarGroupLabel(entry.override, (key) => translateKeyOrText(t, key)),
        content: entry.override.children,
      });
      continue;
    }
    const group = entry.group;
    const shown = group.controls.filter((c) => !overrides.slots.get(`${group.id}.${c.id}`)?.hidden);
    if (!shown.length && !hasContent(override?.children)) continue;
    const controls = shown.map((c) => (
      <SlotControl
        key={c.id}
        props={overrides.slots.get(`${group.id}.${c.id}`)}
        render={() => <ToolbarControl group={group.id} control={c} />}
      />
    ));
    items.push({
      ...placement,
      id: group.id,
      label: t(group.labelKey),
      content: (
        <>
          {controls}
          {override?.children}
        </>
      ),
      menuContent: placement.overflowContent ? (
        <>
          {controls}
          {placement.overflowContent()}
        </>
      ) : undefined,
    });
  }
  return (
    <PresetToolbar
      className={className}
      items={items}
      extras={overrides.extras}
      overflow={overflow}
      label={label}
    />
  );
}

/** A `preset={false}` bar: the host's children, wrapped, with roving focus. */
function HandBar({ className, children }: { className?: string; children?: ReactNode }) {
  const { t } = useTranslation();
  const ref = useRef<HTMLDivElement>(null);
  const roving = useRovingToolbar(ref);
  return (
    <div
      ref={ref}
      className={['pptx-toolbar', 'pptx-toolbar--wrap', className].filter(Boolean).join(' ')}
      role="toolbar"
      aria-label={t('toolbar.ariaLabel')}
      data-overflow="wrap"
      onMouseDown={keepFocus}
      {...roving}
    >
      {children}
    </div>
  );
}

function PresetToolbar({
  className,
  items,
  extras,
  overflow,
  label,
}: {
  className?: string | undefined;
  items: readonly BarItem[];
  extras: readonly ReactNode[];
  overflow: boolean;
  label: string;
}) {
  const { t } = useTranslation();
  useEditorState((s) => s.status);
  // Contextual controls change the width; refit when they appear or go.
  const context = useEditorState((s) => `${s.selection.kind}|${s.geometry?.lineHead ?? ''}`);
  const ref = useRef<HTMLDivElement>(null);
  const more = useRef<HTMLButtonElement>(null);
  const fixed = useRef<HTMLDivElement>(null);
  const panelId = useId();
  // Widths and the collapsed set are keyed by group id, so a group that comes or goes does
  // not hand its width to its neighbour.
  const memory = useRef(createToolbarFitMemory());
  const count = items.length;
  const [collapsed, setCollapsed] = useState<readonly string[]>([]);
  const [open, setOpen] = useState(false);
  const shape = items.map((g) => `${g.id}:${g.priority}:${g.pinned ? 1 : 0}`).join('|');
  const itemsRef = useRef(items);
  itemsRef.current = items;

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    if (!overflow) {
      setCollapsed((cur) => (cur.length ? [] : cur));
      return;
    }
    const fit = (): void => {
      const next = toolbarCollapsedIds(
        el,
        itemsRef.current,
        memory.current,
        more.current,
        fixed.current
      );
      setCollapsed((cur) => (cur.join() === next.join() ? cur : next));
    };
    fit();
    return observeToolbarFit(el, fit);
    // `collapsed` re-runs the fit after groups come back, so their widths are measured again.
  }, [context, collapsed, count, shape, overflow]);

  const group = (index: number, inMenu: boolean) => {
    const g = items[index]!;
    return (
      <div
        key={g.id}
        className={['pptx-toolbar__group', g.className].filter(Boolean).join(' ')}
        role="group"
        aria-label={g.label}
        data-group={g.id}
        data-group-index={inMenu ? undefined : index}
      >
        {/* The group's aria-label already names it: the heading is for sight only. */}
        {inMenu ? (
          <div className="pptx-toolbar__overflow-heading" aria-hidden="true">
            {g.label}
          </div>
        ) : null}
        {inMenu && g.overflowContent ? (g.menuContent ?? g.overflowContent()) : g.content}
      </div>
    );
  };
  const hiddenIds = new Set(collapsed);
  const visible = items.map((_, i) => i).filter((i) => !hiddenIds.has(items[i]!.id));
  const hidden = items.map((_, i) => i).filter((i) => hiddenIds.has(items[i]!.id));
  const roving = useRovingToolbar(ref);
  // The More menu closes for good when nothing is in it, so it does not reopen by itself.
  useEffect(() => {
    if (!hidden.length) setOpen(false);
  }, [hidden.length]);
  return (
    <div
      ref={ref}
      className={['pptx-toolbar', overflow ? '' : 'pptx-toolbar--wrap', className]
        .filter(Boolean)
        .join(' ')}
      role="toolbar"
      aria-label={label}
      data-overflow={overflow ? 'menu' : 'wrap'}
      onMouseDown={keepFocus}
      {...roving}
    >
      {visible.map((index, i) => (
        <Fragment key={items[index]!.id}>
          {i > 0 ? <span className="pptx-toolbar__sep" aria-hidden="true" /> : null}
          {group(index, false)}
        </Fragment>
      ))}
      {extras.length ? (
        <>
          {visible.length ? <span className="pptx-toolbar__sep" aria-hidden="true" /> : null}
          <div ref={fixed} className="pptx-toolbar__group pptx-toolbar__fixed" data-fixed="">
            {extras}
          </div>
        </>
      ) : null}
      {hidden.length ? (
        <>
          {visible.length || extras.length ? (
            <span className="pptx-toolbar__sep" aria-hidden="true" />
          ) : null}
          <button
            ref={more}
            type="button"
            className="pptx-toolbar__button"
            aria-label={t('toolbar.more')}
            data-pptx-tooltip={t('toolbar.more')}
            aria-haspopup="dialog"
            aria-expanded={open}
            data-active={open ? '' : undefined}
            aria-controls={open ? panelId : undefined}
            onMouseDown={keepFocus}
            onClick={() => setOpen((o) => !o)}
            onKeyDown={(event) => {
              // As DOCX Editor's More button: ArrowDown opens the panel at its first control.
              if (event.key !== 'ArrowDown') return;
              event.preventDefault();
              setOpen(true);
            }}
          >
            <Icon name="more_vert" />
          </button>
          {open ? (
            <Popover
              id={panelId}
              anchor={more.current}
              align="end"
              role="dialog"
              focusFirst
              onClose={() => setOpen(false)}
            >
              <div className="pptx-toolbar__overflow" onMouseDown={keepFocus}>
                {hidden.map((i) => group(i, true))}
              </div>
            </Popover>
          ) : null}
        </>
      ) : null}
    </div>
  );
}
