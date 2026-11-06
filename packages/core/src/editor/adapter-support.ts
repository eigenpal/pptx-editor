// Framework-neutral building blocks of the React and Vue adapters. Each adapter keeps only its
// framework glue (context or provide/inject, hooks or composables, elements or vnodes); the
// state they derive comes from here, so the two adapters cannot drift. After DOCX Editor,
// whose adapters share `toolbarCommandState` and the chrome registry the same way.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type { ChromeGroup, ChromeSlotId } from './chrome-controls.ts';
import {
  CHROME_GROUPS,
  PICTURE_BAR_GROUPS,
  commandForSlotValue,
  formattingBarChromeGroups,
  optionRefusal,
  toolbarCommandState,
} from './chrome-controls.ts';
import { NOTES_SLOTS } from './chrome-forms.ts';
import { isTableChromeSlot } from './chrome-table.ts';
import type {
  Editor,
  EditorCommand,
  EditorSnapshot,
  SlideSection,
  TableCellAddress,
  TableStyleChoice,
} from '../contracts/index.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { zoomStep } from './zoom.ts';
import { isTranslationKey, type Translate, type TranslationParams } from '@pptx-editor.dev/i18n';
import {
  warnBuiltInGroupSetting,
  warnUnknownGroupAnchor,
  type ToolbarPartNames,
} from './dev-warnings.ts';

/**
 * The snapshot adapters read before an editor mounts: an empty, editable document at the
 * default 16:9 size. @internal
 */
export const EMPTY_EDITOR_SNAPSHOT: EditorSnapshot = Object.freeze({
  status: 'empty',
  revision: 0,
  mode: 'edit',
  slides: [],
  currentSlide: 0,
  selectedSlides: [],
  slideNumbers: false,
  slideBackground: { kind: 'none', inherited: true },
  slideWidth: 12_192_000,
  slideHeight: 6_858_000,
  zoom: 1,
  fitZoom: true,
  insertions: 0,
  smartGuides: true,
  selection: { kind: 'none' },
  format: null,
  geometry: null,
  canUndo: false,
  canRedo: false,
  notes: '',
  embeddedFontFaces: 0,
  substitutedFonts: [],
  themeColors: [],
}) as EditorSnapshot;

/** The state of one chrome slot, as an adapter's command hook reports it. @internal */
export interface ChromeCommandState {
  /** Whether the command can run now. */
  readonly isEnabled: boolean;
  /** Whether a toggle control shows as pressed. */
  readonly isActive: boolean;
  /** A translation key, or the engine's own text, when the control is disabled. */
  readonly disabledReason: string | null;
  /** The value a value control shows (a font size, a family), when it has one. */
  readonly value?: string | number;
}

/** The state of `slot` now. One source of enabled state for every adapter. @internal */
export function chromeCommandState(
  editor: PptxEditorInstance | null,
  slot: ChromeSlotId
): ChromeCommandState {
  const s = toolbarCommandState(editor, slot);
  return {
    isEnabled: s.enabled,
    isActive: s.active,
    disabledReason: s.disabledReason,
    ...(s.value !== undefined ? { value: s.value } : {}),
  };
}

/** Whether two command states render the same. @internal */
export function sameChromeCommandState(
  a: ChromeCommandState | null,
  b: ChromeCommandState | null
): boolean {
  return (
    !!a &&
    !!b &&
    a.isEnabled === b.isEnabled &&
    a.isActive === b.isActive &&
    a.disabledReason === b.disabledReason &&
    a.value === b.value
  );
}

function detectApple(): boolean {
  if (typeof navigator === 'undefined') return false;
  const nav = navigator as Navigator & { userAgentData?: { platform?: string } };
  // An empty `userAgentData.platform` (some browsers report one) falls back to `platform`.
  return /Mac|iPhone|iPad|iPod/i.test(nav.userAgentData?.platform || nav.platform || '');
}

let apple: boolean | null = null;

/**
 * Shortcut text for the platform: `Ctrl+Shift+G` reads `⇧⌘G` on a Mac. The platform is read
 * once; tests pass `isApple`. Chrome building block, not a stable API. @internal
 */
export function platformShortcuts(text: string, isApple?: boolean): string {
  if (!(isApple ?? (apple ??= detectApple()))) return text;
  // macOS keeps ⌘H to hide the app: Replace is ⇧⌘H there (the find bar answers both).
  // On macOS, ⇧⌘↩ starts the show from the start and ⌘↩ from the current slide, because
  // F5 needs the fn key there; the editor answers both.
  return (
    text
      .replace(/Shift\+F5(?!\w)/g, '⌘↩')
      .replace(/(^|[^+\w])F5(?!\w)/g, '$1⇧⌘↩')
      .replace(/Ctrl\+H(?!\w)/g, '⇧⌘H')
      // Redo is ⇧⌘Z on macOS; Ctrl/Cmd+Y keeps working too.
      .replace(/Ctrl\+Y(?!\w)/g, '⇧⌘Z')
      // New Slide answers Control+M on every platform (a browser keeps ⇧⌘N and ⌘M), so it shows ⌃M.
      .replace(/Ctrl\+M(?!\w)/g, '⌃M')
      .replace(/Ctrl\+Shift\+/g, '⇧⌘')
      .replace(/Ctrl\+Alt\+/g, '⌥⌘')
      .replace(/Ctrl\+/g, '⌘')
      .replace(/Alt\+/g, '⌥')
  );
}

// --- Toolbar arrangement -------------------------------------------------------------------

/** The width the toolbar keeps for its More button, in CSS pixels. */
const MORE_WIDTH = 40;

/** The groups the preset toolbar draws: the registry's non-contextual groups. @internal */
export const TOOLBAR_BAR_GROUPS: readonly ChromeGroup[] = formattingBarChromeGroups();

/**
 * The groups the preset toolbar draws while every selected element is a picture: the bar and the
 * contextual `image` group. @internal
 */
export const TOOLBAR_PICTURE_GROUPS: readonly ChromeGroup[] = PICTURE_BAR_GROUPS;

/**
 * The table controls of the preset bar, as DOCX Editor's short table group: the style, cell
 * fill, borders, and the pen color and weight, then select, delete, insert, merge, and split.
 * The whole group would never fit beside the text controls; the style options, pen style,
 * distribute, and cell margins stay on the ribbon's Table Design and Table Layout tabs and in
 * `Toolbar.Slot`.
 */
const BAR_TABLE_CONTROLS: ReadonlySet<string> = new Set([
  'style',
  'cellFill',
  'borderTarget',
  'borderColor',
  'borderWidth',
  'select',
  'delete',
  'insertRowAbove',
  'insertRowBelow',
  'insertColumnLeft',
  'insertColumnRight',
  'mergeCells',
  'splitCell',
]);

/**
 * The groups the preset toolbar adds for the selection: the `table` group while it is in or
 * on a table (`tableChromeVisible`), with the table slots of `BAR_TABLE_CONTROLS`, as DOCX
 * Editor's table chrome. @internal
 */
export const TOOLBAR_CONTEXT_GROUPS: readonly ChromeGroup[] = CHROME_GROUPS.filter(
  (g) => g.id === 'table'
).map((g) => ({
  ...g,
  controls: g.controls.filter(
    (c) => isTableChromeSlot(`table.${c.id}` as ChromeSlotId) && BAR_TABLE_CONTROLS.has(c.id)
  ),
}));

/**
 * The slide controls the bar shows while nothing is selected: Format Background and Layout,
 * the two slide tasks of an empty selection.
 */
const TOOLBAR_SLIDE_GROUP: ChromeGroup = ((g) => ({
  ...g,
  controls: g.controls.filter((c) => c.id === 'background' || c.id === 'layout'),
}))(CHROME_GROUPS.find((g) => g.id === 'slide')!);

/**
 * What the selection is, for the groups the preset toolbar shows: nothing (or only slides in
 * the rail), text being edited, shapes, pictures only, the text or cells of a table, a table
 * selected as an object, or a focused speaker notes field. @internal
 */
export type ToolbarContext =
  'none' | 'text' | 'shape' | 'picture' | 'table' | 'tableFrame' | 'notes';

/**
 * The toolbar context of a snapshot. Adapters select it, so the bar changes only with it.
 * @internal
 */
export function toolbarContextOf(
  snapshot: Pick<EditorSnapshot, 'selection' | 'image' | 'table'>
): ToolbarContext {
  const sel = snapshot.selection;
  // A table selected as an object keeps Delete, Duplicate, and Arrange; a caret or a range of
  // cells in it shows the text and table groups.
  if (snapshot.table) return sel.kind === 'shapes' ? 'tableFrame' : 'table';
  if (snapshot.image) return 'picture';
  if (sel.kind === 'text') return 'text';
  if (sel.kind === 'shapes' && sel.nodeIds.length > 0) return 'shape';
  return 'none';
}

/**
 * The bar's groups for each context, in bar order. A group in no list stays composable
 * (`Toolbar.Slot`, `preset={false}`, menus, the ribbon); only the preset bar leaves it out.
 */
const CONTEXT_GROUP_IDS: Readonly<Record<ToolbarContext, readonly string[]>> = {
  none: ['history', 'insert', 'slide', 'find'],
  text: ['history', 'font', 'text', 'paragraph', 'list', 'insert', 'format'],
  shape: ['history', 'font', 'text', 'paragraph', 'insert', 'shape', 'format', 'arrange', 'more'],
  picture: ['history', 'insert', 'image', 'arrange', 'more'],
  table: ['history', 'font', 'text', 'paragraph', 'list', 'insert', 'format', 'table'],
  tableFrame: [
    'history',
    'font',
    'text',
    'paragraph',
    'insert',
    'format',
    'arrange',
    'more',
    'table',
  ],
  // Speaker notes take bold, italic, underline, and strikethrough (`NOTES_SLOTS`).
  notes: ['history', 'text', 'find'],
};

const CONTEXT_GROUPS: ReadonlyMap<string, ChromeGroup> = new Map(
  [
    ...TOOLBAR_BAR_GROUPS,
    TOOLBAR_SLIDE_GROUP,
    ...TOOLBAR_PICTURE_GROUPS.filter((g) => g.id === 'image'),
    ...TOOLBAR_CONTEXT_GROUPS,
  ].map((g) => [g.id as string, g])
);

const groupsOf = (ids: readonly string[]): readonly ChromeGroup[] =>
  ids.flatMap((id) => {
    const group = CONTEXT_GROUPS.get(id);
    return group && group.controls.length > 0 ? [group] : [];
  });

const GROUPS_BY_CONTEXT: Readonly<Record<ToolbarContext, readonly ChromeGroup[]>> = {
  none: groupsOf(CONTEXT_GROUP_IDS.none),
  text: groupsOf(CONTEXT_GROUP_IDS.text),
  shape: groupsOf(CONTEXT_GROUP_IDS.shape),
  picture: groupsOf(CONTEXT_GROUP_IDS.picture),
  table: groupsOf(CONTEXT_GROUP_IDS.table),
  tableFrame: groupsOf(CONTEXT_GROUP_IDS.tableFrame),
  notes: groupsOf(CONTEXT_GROUP_IDS.notes).map((g) =>
    g.id === 'text'
      ? { ...g, controls: g.controls.filter((c) => `text.${c.id}` in NOTES_SLOTS) }
      : g
  ),
};

/**
 * Every group the bar can show, in bar order: each context shows some of them in this order.
 * A host group anchored to a group that the selection does not show follows the nearest group
 * before it here, so it keeps its place as the selection changes.
 */
const BAR_ORDER: readonly string[] = [
  'history',
  'font',
  'text',
  'script',
  'paragraph',
  'list',
  'insert',
  'image',
  'shape',
  'format',
  'arrange',
  'more',
  'slide',
  'find',
  'table',
];

/** The order of a `contextual: false` bar: every bar group, then the picture and table groups. */
const FULL_BAR_ORDER: readonly string[] = [
  ...TOOLBAR_BAR_GROUPS.flatMap((g) => (g.id === 'text' ? ['text', 'script'] : [g.id])),
  'image',
  'table',
];

/** The anchor a host group uses on this bar: its `after`, or the nearest shown group before it. */
function shownAnchor(after: string | undefined, shown: ReadonlySet<string>): string | undefined {
  if (after === undefined || shown.has(after)) return after;
  // A bar with every bar group on it is a `contextual: false` bar, with its own order.
  const order = TOOLBAR_BAR_GROUPS.every((g) => shown.has(g.id)) ? FULL_BAR_ORDER : BAR_ORDER;
  const at = order.indexOf(after);
  if (at === -1) return after;
  for (let i = at - 1; i >= 0; i -= 1) if (shown.has(order[i]!)) return order[i];
  return after;
}

// The registry's bar groups stay built in, though `script` has no control on any bar now.
const BUILT_IN_GROUPS: ReadonlySet<string> = new Set([
  ...CONTEXT_GROUPS.keys(),
  ...CHROME_GROUPS.filter((g) => !g.contextual).map((g) => g.id),
]);
const slotsOf = (groups: Iterable<ChromeGroup>): ReadonlySet<string> =>
  new Set([...groups].flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`)));
const BAR_SLOTS = slotsOf(CONTEXT_GROUPS.values());
// A `contextual: false` bar has no slide group: a placed slide slot goes after the groups.
const FULL_BAR_SLOTS = slotsOf([...TOOLBAR_PICTURE_GROUPS, ...TOOLBAR_CONTEXT_GROUPS]);

/**
 * Whether the preset toolbar draws `slot` in one of its groups. Another slot that a host
 * places goes after the groups. @internal
 */
export function isToolbarBarSlot(
  slot: string,
  options: { readonly contextual?: boolean } = {}
): boolean {
  return (options.contextual === false ? FULL_BAR_SLOTS : BAR_SLOTS).has(slot);
}

/** What a host group override names. @internal */
export interface ToolbarGroupOverride {
  /** A built-in group id changes that group; any other id adds a host group. */
  readonly id: string;
  /** Remove the whole group. */
  readonly hidden?: boolean;
  /** The accessible name of a host group. Wins over `labelKey`. */
  readonly label?: string;
  /** A translation key for the name of a host group, used when there is no `label`. */
  readonly labelKey?: string;
  /**
   * The group a host group follows. Without it, a host group follows every built-in group, in
   * order of appearance. A built-in id ignores it, with a development warning.
   */
  readonly after?: string;
}

/** A host group's accessible name: the label, else the translated key, else the id. @internal */
export function toolbarGroupLabel(
  group: Pick<ToolbarGroupOverride, 'id' | 'label' | 'labelKey'>,
  translate: (key: string) => string
): string {
  if (group.label !== undefined) return group.label;
  return group.labelKey !== undefined ? translate(group.labelKey) : group.id;
}

/** One group of the arranged toolbar: a built-in group (with its override) or a host group. @internal */
export type ToolbarArrangedGroup<G extends ToolbarGroupOverride> =
  | { readonly kind: 'built-in'; readonly group: ChromeGroup; readonly override?: G }
  | { readonly kind: 'host'; readonly override: G };

/**
 * The bar order with the host's groups placed: copied from DOCX Editor's `arrangeToolbarGroups`
 * (packages/core/src/editor/toolbar-fit.ts at 398963c7),
 * renamed because `arrangeToolbarGroups` here returns the arranged groups themselves. A host group without `after` follows
 * every built-in group, in order of appearance. A group with `after` follows the group it
 * names, after any host groups already placed there, so several groups anchored to one place
 * keep their order. An anchor may be a host group declared later, or a hidden built-in group
 * (the host group then takes its place). An anchor that never resolves goes to the end.
 * @internal
 */
export function placeToolbarGroups(
  builtIn: readonly string[],
  hosts: readonly { readonly id: string; readonly after?: string | undefined }[]
): readonly string[] {
  const result = [...builtIn];
  // The last id placed after each anchor, so a second group anchored there follows the first.
  const tail = new Map<string, string>();
  let pending = hosts.filter((host) => !builtIn.includes(host.id));
  while (pending.length > 0) {
    const deferred: typeof pending = [];
    for (const host of pending) {
      const anchor = host.after;
      if (anchor === undefined || anchor === host.id) {
        result.push(host.id);
        continue;
      }
      if (!result.includes(anchor)) {
        deferred.push(host);
        continue;
      }
      const after = tail.get(anchor) ?? anchor;
      result.splice(result.indexOf(after) + 1, 0, host.id);
      tail.set(anchor, host.id);
    }
    if (deferred.length === pending.length) {
      // Nothing resolved in this pass: the rest name groups that do not exist.
      result.push(...deferred.map((host) => host.id));
      break;
    }
    pending = deferred;
  }
  return result;
}

/**
 * The built-in groups of the bar for the selection (`toolbarContextOf`): only the controls
 * that act on it. Text shows the font, text, paragraph, and list groups; a shape adds its fill,
 * line, and arrange groups; a picture shows the picture group; a table adds the table group
 * after the text groups, and a table selected as an object keeps Arrange, Duplicate, and
 * Delete too; a focused notes field shows the four notes formats and Find; an empty selection
 * shows Insert and the slide's Background and Layout. Undo and Redo stay first in each. With `contextual: false`, every bar group shows
 * for every selection, and the picture or table group joins after them. The same array comes
 * back for the same arguments. @internal
 */
export function toolbarGroupsFor(
  context: ToolbarContext,
  options: { readonly contextual?: boolean } = {}
): readonly ChromeGroup[] {
  if (options.contextual !== false) return GROUPS_BY_CONTEXT[context];
  // Every bar group for every selection, with the picture or table group after them.
  return context === 'picture'
    ? TOOLBAR_PICTURE_GROUPS
    : context === 'table' || context === 'tableFrame'
      ? TABLE_BAR_GROUPS
      : TOOLBAR_BAR_GROUPS;
}

const TABLE_BAR_GROUPS: readonly ChromeGroup[] = [...TOOLBAR_BAR_GROUPS, ...TOOLBAR_CONTEXT_GROUPS];

/**
 * The preset toolbar's groups after the host's group overrides: built-in groups in registry
 * order (the `table` group after them when `groups` comes from `toolbarGroupsFor` in a
 * table), host groups placed by `placeToolbarGroups`, then hidden groups left out. Both
 * adapters arrange the bar with this, so a group override means the same in React and Vue. A
 * setting the bar cannot use (`after` or a label on a built-in group, an anchor no group has)
 * prints a development warning with the adapter's part names. @internal
 */
export function arrangeToolbarGroups<G extends ToolbarGroupOverride>(
  overrides: ReadonlyMap<string, G>,
  groups: readonly ChromeGroup[] = TOOLBAR_BAR_GROUPS,
  parts?: ToolbarPartNames
): ToolbarArrangedGroup<G>[] {
  const builtIn = new Map(groups.map((g) => [g.id as string, g]));
  const hosts: G[] = [];
  for (const override of overrides.values()) {
    if (builtIn.has(override.id) || BUILT_IN_GROUPS.has(override.id)) {
      for (const setting of ['after', 'label', 'labelKey'] as const)
        if (override[setting] !== undefined) warnBuiltInGroupSetting(override.id, setting, parts);
    } else hosts.push(override);
  }
  // A built-in group that is not on the bar now (the contextual `table` group) is still an anchor.
  const known = [...new Set([...builtIn.keys(), ...BUILT_IN_GROUPS])].concat(
    hosts.map((h) => h.id)
  );
  for (const host of hosts)
    if (host.after !== undefined && !known.includes(host.after))
      warnUnknownGroupAnchor(host.id, host.after, known, parts);
  const out: ToolbarArrangedGroup<G>[] = [];
  const shown = new Set(builtIn.keys());
  const placed = hosts.map((h) => ({ id: h.id, after: shownAnchor(h.after, shown) }));
  for (const id of placeToolbarGroups([...builtIn.keys()], placed)) {
    const override = overrides.get(id);
    if (override?.hidden) continue;
    const group = builtIn.get(id);
    if (group)
      out.push(override ? { kind: 'built-in', group, override } : { kind: 'built-in', group });
    else if (override) out.push({ kind: 'host', override });
  }
  return out;
}

/**
 * Whether one option of a dropdown slot can run, and the engine's reason when it cannot: a
 * menu disables that option and shows the reason, never a click that does nothing (Split Cells
 * into 2 rows on a cell merged over 3 rows). The React toolbar reads it; the Vue pickers will
 * too. @internal
 */
export function slotOptionState(
  editor: Editor | null,
  slot: ChromeSlotId,
  value: unknown
): { readonly enabled: boolean; readonly reason: string | null } {
  const command = commandForSlotValue(slot, value);
  if (!editor || !command) return { enabled: false, reason: 'disabled.notWired' };
  // An option that would change nothing (Distribute on boxes already evenly spaced) stays
  // enabled; only a refusal disables it.
  const reason = optionRefusal(editor, command);
  return reason === null ? { enabled: true, reason: null } : { enabled: false, reason };
}

/** The table styles the Table Design gallery lists (`editor.tableStyles()`). @internal */
export function tableStyleChoices(
  editor: Pick<PptxEditorInstance, 'tableStyles'> | null
): readonly TableStyleChoice[] {
  return editor?.tableStyles() ?? [];
}

/**
 * How many groups fit a bar `inner` pixels wide, given each group's width in order. All fit
 * when their sum fits; otherwise the bar keeps the More button and as many leading groups as
 * fit beside it. @internal
 */
export function toolbarFitCount(
  widths: readonly number[],
  inner: number,
  moreWidth: number = MORE_WIDTH
): number {
  let total = 0;
  for (const w of widths) total += w;
  if (total <= inner) return widths.length;
  let used = moreWidth;
  let fits = 0;
  for (let i = 0; i < widths.length; i += 1) {
    used += widths[i];
    if (used > inner) break;
    fits = i + 1;
  }
  return fits;
}

// --- Chrome text and focus -----------------------------------------------------------------

/**
 * Translate a key that arrives as a plain string, such as an engine refusal reason. A string
 * that is not a catalog key shows as itself, as the translator shows any gap. Use it to show
 * `useEditorCommand(slotId).disabledReason` in your own chrome. @public
 */
export function translateKeyOrText(t: Translate, key: string, params?: TranslationParams): string {
  return isTranslationKey(key) ? t(key, params) : key;
}

/**
 * A key that changes only when the content of a locale catalog changes, not its identity: a
 * host that writes `locale={{ ... }}` inline makes a new object on each render, and the
 * adapters must not retranslate or repaint for it. @internal
 */
export function localeContentKey(locale: unknown): string {
  return locale === undefined || locale === null ? '' : JSON.stringify(locale);
}

/**
 * Whether a chrome mousedown on `target` should keep the editor's focus (call
 * `preventDefault()`): true for everything but form fields, which need focus to work.
 * @internal
 */
export function chromeKeepsFocus(target: EventTarget | null): boolean {
  const tag = (target as { tagName?: string } | null)?.tagName;
  return tag !== 'INPUT' && tag !== 'SELECT' && tag !== 'TEXTAREA';
}

// --- Sections ------------------------------------------------------------------------------

/**
 * The sections, when they cover every slide once and in order: the slide pane and the slide
 * sorter then group the slides under them. Null otherwise: the slides show without groups.
 * @internal
 */
export function groupSlideSections(
  sections: readonly SlideSection[],
  slideCount: number
): readonly SlideSection[] | null {
  let next = 0;
  for (const s of sections) {
    if (s.firstSlide !== next) return null;
    next += s.slideCount;
  }
  return sections.length > 0 && next === slideCount ? sections : null;
}

/** The section that holds slide `index`, or undefined. @internal */
export function sectionOfSlide(
  sections: readonly SlideSection[] | null,
  index: number
): SlideSection | undefined {
  return sections?.find((s) => index >= s.firstSlide && index < s.firstSlide + s.slideCount);
}

/** Where a one-step move of the selected slides lands: `moveSlides` `to` and `section`. @internal */
export interface SlideStep {
  readonly to: number;
  readonly section?: string;
}

/**
 * Where the selected slides land one step up or down (Ctrl/Cmd+Arrow, Move up, Move down), or
 * null at the top or bottom. `to` counts the slides that do not move. With sections that group
 * the slides (`groupSlideSections`), a step across a section header first changes the section
 * and keeps the place: Down from the last slide of a section
 * makes the block the first slides of the next section, and Up from the first slide makes it
 * the last of the section before. Inside a section, the step names that section. @internal
 */
export function slideStep(
  selected: readonly number[],
  count: number,
  step: -1 | 1,
  sections: readonly SlideSection[] | null = null
): SlideStep | null {
  if (selected.length === 0) return null;
  const n = selected.length;
  const first = selected[0]!;
  const last = selected[n - 1]!;
  const edge = step < 0 ? first : last;
  const own = sectionOfSlide(sections, edge);
  const k = own && sections ? sections.indexOf(own) : -1;
  const next = k >= 0 ? sections![k + step] : undefined;
  const atEdge =
    !!own && (step < 0 ? edge === own.firstSlide : edge === own.firstSlide + own.slideCount - 1);
  if (atEdge && next) return { to: step < 0 ? first : last - (n - 1), section: next.id };
  const to = step < 0 ? first - 1 : last - (n - 1) + 1;
  if (to < 0 || to > count - n) return null;
  return own ? { to, section: own.id } : { to };
}

/** The `moveSlides` command of a one-step move (`slideStep`), or null. @internal */
export function stepCommand(
  selected: readonly number[],
  to: SlideStep | null
): EditorCommand | null {
  return to && { type: 'moveSlides', indices: selected, ...to };
}

/**
 * Run a one-step move (`slideStep`). A step across a header into a collapsed section opens that
 * section, so the moved slides stay in view and keep focus. Collapse state is a view setting:
 * no undo step. Returns whether the move ran. @internal
 */
export function runSlideStep(
  editor: PptxEditorInstance | null,
  command: EditorCommand | null
): boolean {
  if (!editor || !command || !editor.exec(command).ok) return false;
  const section = command.type === 'moveSlides' ? command.section : undefined;
  if (section && (editor.snapshot().collapsedSections ?? []).includes(section))
    editor.exec({ type: 'setSectionsCollapsed', collapsed: false, sectionIds: [section] });
  return true;
}

// --- Zoom steps and color mode --------------------------------------------------------------

// The zoom range and its steps are `zoom.ts`'s, shared with the canvas.
export { ZOOM_LEVELS, zoomLimitReason, zoomStep } from './zoom.ts';

/**
 * Step `editor`'s zoom one level in (`1`) or out (`-1`) from the zoom in force now, so two
 * calls in a row step twice. Does nothing at the end of the range. `useZoom` calls it. @internal
 */
export function stepEditorZoom(
  editor: Pick<PptxEditorInstance, 'snapshot' | 'setZoom'> | null,
  direction: 1 | -1
): void {
  const next = editor ? zoomStep(editor.snapshot().zoom, direction) : null;
  if (next !== null) editor?.setZoom(next);
}
export { onEditorEvent, type PptxEditorEvents } from './editor-events.ts';

/** "Slide x of y" for the status bar: one-based, or null with no slides. @internal */
export function slideStatus(
  snapshot: Pick<EditorSnapshot, 'currentSlide' | 'slides'>
): { readonly current: number; readonly total: number } | null {
  const total = snapshot.slides.length;
  return total === 0 ? null : { current: snapshot.currentSlide + 1, total };
}

/** The chrome color mode: light, dark, or the operating system's setting. @public */
export type ColorMode = 'light' | 'dark' | 'system';

/** The theme a root resolved from `colorMode`, never `'system'`. DOCX Editor's type. @public */
export type ResolvedColorMode = 'light' | 'dark';

/** Whether the chrome is dark for `mode`, given the system setting. @internal */
export function resolveDarkChrome(mode: ColorMode | undefined, systemDark: boolean): boolean {
  return mode === 'dark' || (mode === 'system' && systemDark);
}

const DARK_QUERY = '(prefers-color-scheme: dark)';

/** Whether the operating system prefers a dark scheme now. False without `matchMedia`. @internal */
export function systemPrefersDark(): boolean {
  return typeof matchMedia === 'function' && matchMedia(DARK_QUERY).matches;
}

/** Be told when the system setting changes. Returns the unsubscribe function. @internal */
export function subscribeSystemDark(listener: () => void): () => void {
  if (typeof matchMedia !== 'function') return () => undefined;
  const query = matchMedia(DARK_QUERY);
  query.addEventListener('change', listener);
  return () => query.removeEventListener('change', listener);
}

/**
 * A host's translation function for the editor chrome, as DOCX Editor's `ChromeTranslate`: a
 * key of the locale catalog, and its parameters, to display text. @public
 */
export type ChromeTranslate = (key: string, params?: TranslationParams) => string;

/**
 * The caret: the text box that holds it, and a paragraph and an offset in that text box, as
 * DOCX Editor's `EditorCaret`. `cell` names the table cell when the text box is a table. @public
 */
export interface EditorCaret {
  readonly nodeId: string;
  readonly paragraph: number;
  readonly offset: number;
  readonly cell?: TableCellAddress;
}

/**
 * The caret of `snapshot`: the focus of a text selection, or null when no text is being
 * edited. `previous` comes back when the caret did not move, so a hook can compare by
 * identity. @internal
 */
export function editorCaretOf(
  snapshot: Pick<EditorSnapshot, 'selection'>,
  previous: EditorCaret | null
): EditorCaret | null {
  const selection = snapshot.selection;
  if (selection.kind !== 'text') return null;
  const { focus, nodeId, cell } = selection;
  if (
    previous &&
    previous.nodeId === nodeId &&
    previous.paragraph === focus.paragraph &&
    previous.offset === focus.offset &&
    previous.cell?.row === cell?.row &&
    previous.cell?.column === cell?.column
  )
    return previous;
  return {
    nodeId,
    paragraph: focus.paragraph,
    offset: focus.offset,
    ...(cell ? { cell: { row: cell.row, column: cell.column } } : {}),
  };
}
